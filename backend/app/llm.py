import httpx
import re
import json
from jsonschema import Draft202012Validator, ValidationError
from typing import Protocol, Any
from .config import Settings

class ServiceError(Exception):
    def __init__(self, code: str, message: str):
        self.code, self.message = code, message
        super().__init__(message)

class LLM(Protocol):
    async def chat(self, messages: list[dict], tools: list[dict]) -> dict: ...

def final_content(content: str) -> str:
    """Accept final text only; tolerate providers leaking a closing thinking marker."""
    if not isinstance(content, str):
        raise ServiceError('model_response_invalid', 'متن پاسخ مدل معتبر نیست.')
    parts = re.split(r'\\?</think\s*>', content, flags=re.IGNORECASE)
    result = parts[-1].strip()
    if re.search(r'<think(?:\s[^>]*)?>', result, re.IGNORECASE):
        raise ServiceError('model_reasoning_incomplete', 'مدل فقط متن پردازش ناقص برگرداند؛ پاسخ نهایی دریافت نشد.')
    return result

class Ollama:
    def __init__(self, settings: Settings, transport: Any = None):
        self.settings, self.transport = settings, transport

    async def request(self, path: str, payload: dict | None = None) -> dict:
        try:
            async with httpx.AsyncClient(base_url=self.settings.llm_url, trust_env=False,
                                         timeout=self.settings.timeout, transport=self.transport) as c:
                response = await (c.get(path) if payload is None else c.post(path, json=payload))
                response.raise_for_status()
                data = response.json()
                if not isinstance(data, dict): raise ValueError()
                return data
        except httpx.TimeoutException as exc:
            raise ServiceError('ollama_timeout', 'مهلت پاسخ Ollama پایان یافت؛ نتیجه‌ای تأیید نشد. اجرای مدل روی CPU ممکن است زمان بیشتری لازم داشته باشد.') from exc
        except (httpx.HTTPError, ValueError) as exc:
            raise ServiceError('ollama_unavailable', 'Ollama پاسخ معتبر نداد؛ سرویس و مهلت پاسخ را بررسی کنید.') from exc

    async def health(self) -> dict:
        data = await self.request('/api/tags')
        models = data.get('models', [])
        installed = any(m.get('name') == self.settings.model for m in models)
        if not installed:
            raise ServiceError('model_missing', f'مدل {self.settings.model} هنوز نصب کامل نشده است؛ دانلود جاری را کامل کنید. برنامه دانلودی آغاز نمی‌کند.')
        metadata = await self.request('/api/show', {'model': self.settings.model})
        capabilities = metadata.get('capabilities', [])
        if 'tools' not in capabilities or 'completion' not in capabilities:
            raise ServiceError('capability_missing', 'مدل نصب‌شده قابلیت completion و tools را در متادیتای Ollama تأیید نکرد.')
        return {'available': True, 'model': self.settings.model, 'capabilities': capabilities,
                'digest': next((m.get('digest') for m in models if m.get('name') == self.settings.model), None)}

    async def release_for_speech(self):
        """Unload only this app's configured model from RAM; never pull/delete models."""
        running=await self.request('/api/ps')
        if any(m.get('name')==self.settings.model for m in running.get('models', [])):
            await self.request('/api/generate', {'model':self.settings.model,'keep_alive':0})

    async def needs_speech_release(self, available_bytes):
        running=await self.request('/api/ps')
        if any(m.get('name')==self.settings.model for m in running.get('models',[])):return False
        installed=await self.request('/api/tags')
        size=next((m.get('size') for m in installed.get('models',[]) if m.get('name')==self.settings.model),None)
        return available_bytes is None or not isinstance(size,int) or available_bytes < size*1.4+512*1024**2

    async def route(self, text):
        data=await self.request('/api/chat',{'model':self.settings.model,'stream':False,'think':False,
            'format':{'type':'object','properties':{'route':{'enum':['chat','company']}},'required':['route'],'additionalProperties':False},
            'messages':[{'role':'system','content':'Classify ONLY the current user request. company means reading actual private company records/metrics. General conversation, capabilities, writing, explaining and summarizing provided text are chat. Return JSON only. /no_think'},
                        {'role':'user','content':text}],
            'options':{'temperature':0,'num_ctx':self.settings.context,'num_predict':40},'keep_alive':'2m'})
        try:
            route=json.loads(final_content(data['message']['content']))['route']
            if route not in {'chat','company'}:raise ValueError()
            return route
        except (ValueError,KeyError,TypeError) as exc:raise ServiceError('routing_failed','تشخیص مسیر درخواست کامل نشد؛ ابزاری اجرا نشد.') from exc

    async def select_tools(self,text,definitions):
        names=[d['function']['name'] for d in definitions]
        catalog=[{'name':d['function']['name'],'purpose':d['function'].get('description','').split('.')[0][:140]} for d in definitions]
        data=await self.request('/api/chat',{'model':self.settings.model,'stream':False,'think':False,
            'format':{'type':'object','properties':{'tools':{'type':'array','items':{'enum':names},'minItems':1,'maxItems':4,'uniqueItems':True}},'required':['tools'],'additionalProperties':False},
            'messages':[{'role':'system','content':'Select up to four relevant tools for the current request from this catalog. Return JSON only. /no_think\n'+json.dumps(catalog,ensure_ascii=False)}, {'role':'user','content':text}],
            'options':{'temperature':0,'num_ctx':self.settings.context,'num_predict':180},'keep_alive':'2m'})
        try:
            selected=json.loads(final_content(data['message']['content']))['tools']
            if not isinstance(selected,list) or not selected or len(selected)>4 or any(x not in names for x in selected):raise ValueError()
            return [d for d in definitions if d['function']['name'] in selected]
        except (ValueError,KeyError,TypeError) as exc:raise ServiceError('routing_failed','انتخاب ابزار معتبر کامل نشد؛ ابزاری اجرا نشد.') from exc

    async def chat(self, messages: list[dict], tools: list[dict]) -> dict:
        # Constrain the response protocol as well as advertising native tool schemas.
        # This prevents legacy templates from spending the output budget on prose
        # instead of producing a machine-readable call. All calls still go through
        # the orchestrator's allowlist, schema validation and real MCP execution.
        schema = {'oneOf': [{'type': 'object', 'properties': {
            'action': {'const': 'answer'}, 'text': {'type': 'string', 'minLength': 1}},
            'required': ['action', 'text'], 'additionalProperties': False}] + [
            {'type': 'object', 'properties': {'action': {'const': 'tool'},
                'name': {'const': t['function']['name']}, 'arguments': t['function']['parameters']},
             'required': ['action', 'name', 'arguments'], 'additionalProperties': False} for t in tools]}
        protocol = {'role': 'system', 'content':
            'Respond only with the required JSON object. To use an advertised tool return '
            '{"action":"tool","name":"tool_name","arguments":{...}}. '
            'When done return {"action":"answer","text":"پاسخ نهایی کوتاه فارسی"}. '
            'Never put analysis or English thinking in text. For company data call the appropriate tool first. '
            'For greetings, speech tests, or unclear messages, answer briefly or ask for clarification without company tools. '
            'Do not claim a tool result before receiving it. /no_think'}
        request_messages=messages
        if tools:
            if messages and messages[0].get('role')=='system':
                request_messages=[{'role':'system','content':messages[0]['content']+'\n'+protocol['content']},*messages[1:]]
            else: request_messages=[protocol,*messages]
        payload = {'model': self.settings.model,
            'messages': request_messages,
            'tools': tools, 'stream': False, 'think': False,
            'options': {'temperature': self.settings.temperature, 'num_ctx': self.settings.context,
                        'num_predict': 700}, 'keep_alive': '2m'}
        payload['format'] = schema
        data = await self.request('/api/chat', payload)
        message = data.get('message')
        if not isinstance(message, dict) or message.get('role') != 'assistant':
            raise ServiceError('model_response_invalid', 'ساختار پاسخ مدل معتبر نیست.')
        # Ollama's separate thinking field is intentionally never returned to the UI/history.
        message['content'] = final_content(message.get('content', ''))
        if not message.get('content') and not message.get('tool_calls'):
            raise ServiceError('model_response_empty', 'مدل پاسخ قابل نمایش یا فراخوانی ابزار برنگرداند.')
        if data.get('done_reason') == 'length':
            raise ServiceError('model_output_limit', 'پاسخ مدل به سقف خروجی رسید؛ پرسش را کوتاه‌تر کنید.')
        if not message.get('tool_calls'):
            try:
                decision = json.loads(message['content'])
                Draft202012Validator(schema).validate(decision)
            except (ValueError, ValidationError) as exc:
                raise ServiceError('model_response_invalid', 'مدل قرارداد پاسخ یا فراخوانی ابزار را رعایت نکرد؛ عملیاتی اجرا نشد.') from exc
            if decision['action'] == 'tool':
                message['content'] = ''
                message['tool_calls'] = [{'function': {'name': decision['name'], 'arguments': decision['arguments']}}]
            else:
                message['content'] = decision['text']
        return {k: message[k] for k in ('role', 'content', 'tool_calls') if k in message}
