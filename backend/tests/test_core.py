import asyncio
import json
import unittest
from contextlib import asynccontextmanager
from types import SimpleNamespace
import httpx
from fastapi.testclient import TestClient
from app.config import Settings
from app.llm import Ollama, ServiceError, final_content
from app.agent import Agent
from app.mcp_client import Tools
from app.main import create_app
from app.reporting import channel_percentage

class FakeSession:
    def __init__(self): self.calls = []
    async def list_tools(self, cursor=None):
        return SimpleNamespace(tools=[SimpleNamespace(name='sales_order_count', description='Count company sales orders', inputSchema={
            'type': 'object', 'properties': {'kind': {'type': 'string', 'enum': ['all', 'paid']}}})], nextCursor=None)
    async def call_tool(self, name, arguments):
        self.calls.append((name, arguments))
        return SimpleNamespace(isError=False, structuredContent={'count': 15000}, content=[])

class FakeLLM:
    def __init__(self, responses): self.responses, self.messages = responses, []
    async def health(self): return {'available': True}
    async def chat(self, messages, tools):
        self.messages.append(list(messages))
        return self.responses.pop(0)

class FakeMCP:
    @asynccontextmanager
    async def connect(self): yield Tools(FakeSession(), {'sales_order_count'})

class CoreTests(unittest.IsolatedAsyncioTestCase):
    async def test_percentage_uses_real_counts_without_currency(self):
        event={'tool':'sales_breakdown','arguments':{'group':'channel','limit':10},'result':{'data':[
            {'group':'online','orders':3635},{'group':'campaign','orders':3634},
            {'group':'call_center','orders':3634},{'group':'dealer','orders':3633}]}}
        answer=channel_percentage('درصد آنلاین',[event])
        self.assertIn('25.007',answer)
        self.assertIn('14,536',answer)
        self.assertNotIn('تومان',answer)
        event['arguments']['limit']=4
        self.assertNotIn('25.007',channel_percentage('درصد آنلاین',[event]))
    async def test_final_text_excludes_thinking(self):
        for prefix in ['<think>internal</think>', 'internal</think>', 'internal\\</think>']:
            self.assertEqual(final_content(prefix+' پاسخ فارسی'), 'پاسخ فارسی')
        with self.assertRaises(ServiceError): final_content('<think>unfinished')

    async def test_structured_protocol_preserves_native_tools(self):
        captured=[]
        def handler(request):
            if request.url.path == '/api/tags':
                return httpx.Response(200,json={'models':[{'name':'qwen3:4b'}]})
            if request.url.path == '/api/show':
                return httpx.Response(200,json={'capabilities':['completion','tools','thinking'],
                    'template':'{{- if and (ne .Role "assistant") $last }}<|im_start|>assistant\n<think>\n{{ end }}'})
            captured.append(json.loads(request.content))
            return httpx.Response(200,json={'message':{'role':'assistant','content':'',
                'thinking':'private', 'tool_calls':[{'function':{'name':'sales_breakdown','arguments':{'group':'channel','limit':10}}}]}})
        llm=Ollama(Settings(),httpx.MockTransport(handler))
        await llm.health()
        definitions=[{'type':'function','function':{'name':'sales_breakdown','parameters':{'type':'object'}}}]
        response=await llm.chat([{'role':'user','content':'پرسش'}],definitions)
        self.assertIn('format',captured[0])
        self.assertEqual(captured[0]['tools'],definitions)
        self.assertNotIn('thinking',response)
        self.assertEqual(response['tool_calls'][0]['function']['arguments']['limit'],10)

    async def test_structured_call_is_decoded_and_validated(self):
        defs=[{'type':'function','function':{'name':'sales_breakdown','parameters':{
            'type':'object','properties':{'group':{'type':'string'},'limit':{'type':'integer'}},
            'required':['group','limit'],'additionalProperties':False}}}]
        for decision,valid in [({'action':'tool','name':'sales_breakdown','arguments':{'group':'channel','limit':10}},True),
                               ({'action':'tool','name':'shell','arguments':{}},False)]:
            llm=Ollama(Settings(),httpx.MockTransport(lambda r:httpx.Response(200,json={
                'message':{'role':'assistant','content':json.dumps(decision)}})))
            if valid:
                result=await llm.chat([],defs)
                self.assertEqual(result['tool_calls'][0]['function']['arguments']['limit'],10)
            else:
                with self.assertRaises(ServiceError):await llm.chat([],defs)

    async def test_real_loop_passes_results_back(self):
        llm = FakeLLM([{'role':'assistant','content':'','tool_calls':[{'function':{'name':'sales_order_count','arguments':{}}}]},
                       {'role':'assistant','content':'تعداد سفارش‌ها ۱۵۰۰۰ است.'}])
        tools = Tools(FakeSession(), {'sales_order_count'})
        events = []
        async def emit(e): events.append(e)
        answer, additions, evidence = await Agent(Settings(),llm).run([], 'تعداد سفارش‌ها؟',tools,emit)
        self.assertIn('۱۵۰۰۰',answer)
        self.assertEqual(llm.messages[1][-1]['role'],'tool')
        self.assertIn('15000',llm.messages[1][-1]['content'])
        self.assertEqual(evidence[0]['result'],{'count':15000})
        self.assertEqual([e['stage'] for e in events], ['mcp','routing','model','model','tool','tool','model','model'])
        tool_events = [e for e in events if e['stage'] == 'tool']
        self.assertEqual(tool_events[0]['id'], tool_events[1]['id'])
        self.assertEqual(tool_events[0]['arguments'], {})
        self.assertEqual(tool_events[0]['operation'], 'tools/call')

    async def test_unauthorized_and_invalid_arguments(self):
        session = FakeSession(); tools = Tools(session, {'sales_order_count'})
        await tools.discover()
        for name,args in [('shell',{}),('sales_order_count',{'kind':'delete'}),('sales_order_count',{'sql':'DROP'}),
                          ('sales_order_count',{'limit':10.0}),('sales_order_count',{'limit':'10'})]:
            with self.assertRaises(ServiceError): await tools.execute(name,args)
        self.assertEqual(session.calls,[])

    async def test_iteration_limit(self):
        call = {'role':'assistant','content':'','tool_calls':[{'function':{'name':'sales_order_count','arguments':{}}}]}
        llm=FakeLLM([call,call])
        async def emit(e): pass
        with self.assertRaises(ServiceError) as cm:
            await Agent(Settings(max_steps=1),llm).run([],'تعداد',Tools(FakeSession(),{'sales_order_count'}),emit)
        self.assertEqual(cm.exception.code,'tool_limit')

    async def test_missing_model_no_pull(self):
        seen=[]
        def handler(request):
            seen.append(request.url.path)
            return httpx.Response(200,json={'models':[]})
        with self.assertRaises(ServiceError) as cm:
            await Ollama(Settings(),httpx.MockTransport(handler)).health()
        self.assertEqual(cm.exception.code,'model_missing')
        self.assertEqual(seen,['/api/tags'])

    async def test_metadata_not_name(self):
        def handler(request):
            return httpx.Response(200,json={'models':[{'name':'qwen3:4b'}]} if request.url.path=='/api/tags' else {'capabilities':['completion']})
        with self.assertRaises(ServiceError) as cm: await Ollama(Settings(),httpx.MockTransport(handler)).health()
        self.assertEqual(cm.exception.code,'capability_missing')

    async def test_empty_and_malformed_responses(self):
        for value in [{'message':{}},{'message':{'role':'assistant','content':''}}]:
            with self.assertRaises(ServiceError):
                await Ollama(Settings(),httpx.MockTransport(lambda r:httpx.Response(200,json=value))).chat([],[])

    async def test_tool_error(self):
        class Broken(FakeSession):
            async def call_tool(self,*args): return SimpleNamespace(isError=True)
        t=Tools(Broken(),{'sales_order_count'});await t.discover()
        with self.assertRaises(ServiceError): await t.execute('sales_order_count',{})

    async def test_tool_error_text_is_preserved(self):
        class Broken(FakeSession):
            async def call_tool(self,*args):
                return SimpleNamespace(isError=True,content=[SimpleNamespace(type='text',text='sales_breakdown failed: HTTP 422')])
        t=Tools(Broken(),{'sales_order_count'});await t.discover()
        with self.assertRaises(ServiceError) as cm: await t.execute('sales_order_count',{})
        self.assertIn('HTTP 422',cm.exception.message)

class APITests(unittest.TestCase):
    def test_repeated_question_is_not_accepted_as_report(self):
        question='تعداد سفارش‌های شرکت چقدر است؟'
        client=TestClient(create_app(Settings(),FakeLLM([{'role':'assistant','content':question}]),FakeMCP()))
        client.get('/');key=client.post('/api/conversations').json()['id']
        response=client.post('/api/chat',json={'conversation_id':key,'text':question})
        self.assertIn('answer_echo',response.text)
        self.assertNotIn('"stage": "final"',response.text)
    def setUp(self):
        self.client=TestClient(create_app(Settings(),FakeLLM([{'role':'assistant','content':'سلام'}]),FakeMCP()))
        self.client.get('/')
    def test_chat_and_history(self):
        self.assertEqual(self.client.get('/static/app.js').status_code,200)
        key=self.client.post('/api/conversations').json()['id']
        result=self.client.post('/api/chat',json={'conversation_id':key,'text':'سلام'})
        self.assertIn('"stage": "final"',result.text)
        events = [json.loads(line) for line in result.text.splitlines()]
        self.assertEqual([e['sequence'] for e in events], list(range(1,len(events)+1)))
        self.assertEqual(len({e['run_id'] for e in events}),1)
        self.assertTrue(all(e['elapsed_ms'] >= 0 for e in events))
        self.assertEqual(len(self.client.get('/api/conversations/'+key).json()['messages']),2)
    def test_isolation_origin_size(self):
        key=self.client.post('/api/conversations').json()['id']
        other=TestClient(self.client.app);other.get('/')
        self.assertEqual(other.get('/api/conversations/'+key).status_code,400)
        self.assertEqual(self.client.post('/api/conversations',headers={'origin':'http://evil.test'}).status_code,403)
        self.assertEqual(self.client.post('/api/chat',content='x'*20001).status_code,413)
    def test_mcp_failure_keeps_app_alive(self):
        class Broken:
            @asynccontextmanager
            async def connect(self):
                raise ServiceError('mcp_unavailable','خاموش')
                yield
        client=TestClient(create_app(Settings(),FakeLLM([]),Broken()));client.get('/')
        key=client.post('/api/conversations').json()['id']
        result=client.post('/api/chat',json={'conversation_id':key,'text':'تعداد سفارش‌های شرکت چقدر است؟'})
        self.assertIn('mcp_unavailable',result.text)
        self.assertEqual(client.get('/api/health').json()['application'],'ok')
    def test_token_required(self):
        client=TestClient(create_app(Settings(token='test-secret'),FakeLLM([]),FakeMCP()))
        self.assertEqual(client.get('/api/health').status_code,401)
    def test_reject_write_tool_configuration(self):
        with self.assertRaises(ValueError): Settings(allowlist=frozenset({'delete_customer'}))

if __name__=='__main__': unittest.main()
