"""Official MCP v1 SDK adapter; each run owns a session in the same async task."""
import asyncio
import json
import time
from contextlib import asynccontextmanager
import httpx
from jsonschema import Draft202012Validator, ValidationError
from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client
from .config import Settings
from .llm import ServiceError

class CompanyMCP:
    def __init__(self, settings: Settings):
        self.settings = settings

    @asynccontextmanager
    async def connect(self):
        try:
            def factory(headers=None, timeout=None, auth=None):
                return httpx.AsyncClient(headers=headers, timeout=timeout or 30, auth=auth, trust_env=False)
            async with streamablehttp_client(self.settings.mcp_url, httpx_client_factory=factory) as (r, w, _):
                async with ClientSession(r, w) as session:
                    await session.initialize()
                    yield Tools(session, self.settings.allowlist)
        except ServiceError:
            raise
        except Exception as exc:
            def known(error):
                if isinstance(error, ServiceError): return error
                for child in getattr(error, 'exceptions', ()):
                    found = known(child)
                    if found: return found
                return None
            original = known(exc)
            if original is not None: raise original from exc
            # ExceptionGroups from SDK task groups must also be sanitized.
            raise ServiceError('mcp_unavailable', 'ارتباط MCP کامل نشد؛ سرویس‌های شرکت را بررسی کنید. عملیات خودکار تکرار نشد.') from exc

class Tools:
    def __init__(self, session, allowlist, result_limit=128000):
        self.session, self.allowlist = session, allowlist
        self.result_limit = result_limit
        self.schemas: dict[str, dict] = {}
        self.descriptions: dict[str, str] = {}
        self.discovered_total = 0

    async def discover(self) -> list[dict]:
        cursor, seen = None, set()
        for _ in range(10):
            response = await asyncio.wait_for(self.session.list_tools(cursor=cursor), 20)
            self.discovered_total += len(response.tools)
            for tool in response.tools:
                if tool.name not in self.allowlist: continue
                schema = dict(tool.inputSchema)
                schema.setdefault('additionalProperties', False)
                Draft202012Validator.check_schema(schema)
                self.schemas[tool.name] = schema
                self.descriptions[tool.name] = (tool.description or tool.name)[:2000]
            cursor = response.nextCursor
            if not cursor: break
            if cursor in seen: raise ServiceError('mcp_pagination', 'صفحه‌بندی MCP معتبر نیست.')
            seen.add(cursor)
        else: raise ServiceError('mcp_tool_limit', 'فهرست ابزار بیش از حد مجاز است.')
        return [{'type': 'function', 'function': {'name': name,
                 'description': self.descriptions[name],
                 'parameters': schema}} for name, schema in self.schemas.items()]

    def validate(self, name: str, arguments: dict):
        if name not in self.allowlist or name not in self.schemas:
            raise ServiceError('tool_forbidden', 'ابزار درخواست‌شده مجاز نیست؛ اجرا نشد.')
        # JSON Schema regards 10.0 as an integer; HTTP query serialization does not.
        # Preserve strict Python integers for the company's row-limit query parameter.
        if 'limit' in arguments and type(arguments['limit']) is not int:
            raise ServiceError('tool_arguments_invalid', 'پارامتر limit باید عدد صحیح باشد، نه متن یا عدد اعشاری؛ ابزار اجرا نشد.')
        try: Draft202012Validator(self.schemas[name]).validate(arguments)
        except ValidationError as exc:
            raise ServiceError('tool_arguments_invalid', 'آرگومان ابزار مطابق قرارداد نیست؛ اجرا نشد.') from exc

    async def execute(self, name: str, arguments: dict) -> tuple[dict, int]:
        self.validate(name, arguments)
        started = time.monotonic()
        try: result = await asyncio.wait_for(self.session.call_tool(name, arguments), 30)
        except Exception as exc:
            raise ServiceError('tool_failed', 'ابزار پاسخ نداد؛ عملیات خودکار تکرار نشد.') from exc
        if result.isError:
            detail = '\n'.join(c.text for c in getattr(result, 'content', []) if c.type == 'text')[:2000]
            raise ServiceError('tool_failed', 'خطای ابزار ' + name + ': ' + (detail or 'متن خطا از MCP دریافت نشد.'))
        if result.structuredContent is not None:
            value = result.structuredContent
        else:
            text = '\n'.join(c.text for c in result.content if c.type == 'text')
            try:
                decoded=json.loads(text)
                value=decoded if isinstance(decoded,dict) else {'data':decoded} if isinstance(decoded,list) else {'text':text}
            except ValueError:value={'text':text}
        if len(json.dumps(value, ensure_ascii=False)) > self.result_limit:
            raise ServiceError('tool_result_limit', 'نتیجه بیش از حد بزرگ است؛ پرسش را محدودتر کنید.')
        return value, round((time.monotonic() - started) * 1000)
