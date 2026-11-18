import unittest
from contextlib import asynccontextmanager
from fastapi.testclient import TestClient
from app.main import create_app
from app.config import Settings, READ_ONLY
from app.routing import route_request, relevant_tools
from test_core import FakeLLM, FakeMCP

class RoutingTests(unittest.TestCase):
    def test_current_question_is_not_previous_sales_report(self):
        self.assertEqual(route_request('تو دیگه چی کارا میکنی'),'capabilities')
        self.assertEqual(route_request('سلام تعداد سفارش‌های شرکت چقدر است؟'),'company')
        self.assertEqual(route_request('خلاصه آمار فروش شرکت را بده'),'company')
        self.assertEqual(route_request('تفاوت درآمد و سود چیست؟'),'chat')
        self.assertEqual(route_request('این ویس را خلاصه کن\n\nمتن فایل صوتی:\nتعداد فروش شرکت چقدر است؟'),'chat')
        self.assertEqual(route_request('به سؤال ویس جواب بده\n\nمتن فایل صوتی:\nتعداد سفارش‌های شرکت چقدر است؟'),'company')
        self.assertEqual(len(READ_ONLY),46)
        self.assertEqual(len(Settings().allowlist),46)

    def test_capability_answer_needs_no_model_or_mcp(self):
        class NoMCP:
            @asynccontextmanager
            async def connect(self):
                raise AssertionError('MCP should not be contacted')
                yield
        with TestClient(create_app(Settings(),FakeLLM([]),NoMCP())) as c:
            c.get('/');key=c.post('/api/conversations').json()['id']
            r=c.post('/api/chat',json={'conversation_id':key,'text':'تو دیگه چی کارا میکنی'})
            self.assertIn('فایل صوتی',r.text)
            self.assertNotIn('sales_breakdown',r.text)
            self.assertNotIn('25.007',r.text)
            self.assertNotIn('"stage": "tool"',r.text)

    def test_chat_does_not_connect_to_mcp(self):
        class NoMCP:
            @asynccontextmanager
            async def connect(self):
                raise AssertionError('No company tools for chat')
                yield
        with TestClient(create_app(Settings(),FakeLLM([{'role':'assistant','content':'سلام!'}]),NoMCP())) as c:
            c.get('/');key=c.post('/api/conversations').json()['id']
            r=c.post('/api/chat',json={'conversation_id':key,'text':'سلام'})
            self.assertIn('سلام!',r.text)
            self.assertNotIn('"stage": "tool"',r.text)

if __name__=='__main__':unittest.main()
