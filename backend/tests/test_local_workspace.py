import tempfile
import unittest
from pathlib import Path
from datetime import timedelta,datetime
from app.local_workspace import LocalWorkspace,utcnow
from app.llm import ServiceError
from app.main import create_app
from app.config import Settings
from fastapi.testclient import TestClient
from test_core import FakeLLM,FakeMCP

EML=b'From: fixture@example.test\r\nSubject: Explicit local test invoice\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nPlease review invoice fixture-42. This is a test fixture, not Outlook.\r\n'

class WorkspaceTests(unittest.TestCase):
    def test_chat_registers_real_local_reminder_but_requires_outlook_connection(self):
        with tempfile.TemporaryDirectory() as temp:
            with TestClient(create_app(Settings(),FakeLLM([]),FakeMCP(),workspace_path=Path(temp)/'db.sqlite3')) as c:
                c.get('/');key=c.post('/api/conversations').json()['id']
                reminder=c.post('/api/chat',json={'conversation_id':key,'text':'۱۰ دقیقه دیگه یادم بنداز ایمیل را پیگیری کنم'})
                self.assertIn('یادآوری واقعاً ثبت شد',reminder.text)
                saved=c.get('/api/reminders').json()['items']
                self.assertEqual(len(saved),1)
                self.assertEqual(saved[0]['status'],'pending')
                self.assertEqual(saved[0]['title'],'ایمیل را پیگیری کنم')
                outlook=c.post('/api/chat',json={'conversation_id':key,'text':'ایمیل‌هام در Outlook را بخوان'})
                self.assertIn('outlook_not_connected',outlook.text)
    def test_reminder_survives_reopen_and_fires_once(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'workspace.sqlite3';one=LocalWorkspace(path);due=utcnow()+timedelta(seconds=30)
            result=one.add_reminder('owner','test reminder',due,'request-id-0001')
            self.assertEqual(one.add_reminder('owner','test reminder',due,'request-id-0001')['id'],result['id'])
            two=LocalWorkspace(path)
            self.assertEqual(two.reminders('owner')[0]['status'],'pending')
            self.assertEqual(two.tick(due+timedelta(seconds=1)),1)
            self.assertEqual(two.tick(due+timedelta(seconds=2)),0)
            self.assertEqual(two.reminders('owner')[0]['status'],'due')
            self.assertEqual(two.reminders('other'),[])
            with self.assertRaises(ServiceError):two.acknowledge('other',result['id'])
            two.acknowledge('owner',result['id'])
            self.assertEqual(two.reminders('owner')[0]['status'],'seen')

    def test_local_mail_is_real_parsed_content_and_owner_scoped(self):
        with tempfile.TemporaryDirectory() as temp:
            store=LocalWorkspace(Path(temp)/'db.sqlite3');mail=store.import_mail('a',EML)
            results=store.search_mail('a','fixture-42')
            self.assertEqual(results['items'][0]['id'],mail['id'])
            self.assertEqual(store.search_mail('b','fixture-42')['items'],[])
            with self.assertRaises(ServiceError):store.read_mail('b',mail['id'])
            with self.assertRaises(ServiceError):store.import_mail('a',b'<html>not email</html>')
            with self.assertRaises(ServiceError):store.add_reminder('b','cross user',utcnow()+timedelta(minutes=1),'cross-id-0001',mail['id'])
            self.assertEqual(store.search_mail('a','%')['items'],[])

    def test_naive_past_and_changed_retry_are_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            store=LocalWorkspace(Path(temp)/'db.sqlite3')
            for due in [datetime(2026,1,1),utcnow()-timedelta(minutes=1)]:
                with self.assertRaises(ServiceError):store.add_reminder('a','invalid',due,'invalid-key-0001')
            due=utcnow()+timedelta(minutes=1);store.add_reminder('a','first',due,'repeated-key-0001')
            with self.assertRaises(ServiceError):store.add_reminder('a','changed',due,'repeated-key-0001')

    def test_api_disabled_integrations_and_mail_reminder_link(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'db.sqlite3'
            with TestClient(create_app(Settings(),FakeLLM([]),FakeMCP(),workspace_path=path)) as c:
                c.get('/')
                status=c.get('/api/integrations').json()
                self.assertFalse(status['company_gpu']['connected'])
                self.assertFalse(status['outlook']['connected'])
                self.assertFalse(status['telegram']['connected'])
                mail=c.post('/api/local-mail',content=EML).json()
                payload={'title':'Review test email','due_at':(utcnow()+timedelta(minutes=1)).isoformat(),'request_key':'linked-key-0001','mail_id':mail['id']}
                response=c.post('/api/reminders',json=payload)
                self.assertEqual(response.status_code,200)
                self.assertEqual(response.json()['mail_id'],mail['id'])
                self.assertFalse(response.json()['telegram_sent'])
                cookie=c.cookies.get('agent_session')
            with TestClient(create_app(Settings(),FakeLLM([]),FakeMCP(),workspace_path=path)) as c:
                c.cookies.set('agent_session',cookie)
                self.assertEqual(len(c.get('/api/reminders').json()['items']),1)

if __name__=='__main__':unittest.main()
