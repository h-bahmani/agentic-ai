import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from fastapi.testclient import TestClient
from app.main import create_app
from app.llm import ServiceError

class LinkTests(unittest.TestCase):
    def test_inline_url_fetches_source_before_model(self):
        class Model:
            async def chat(self,messages,tools):
                assert 'ACTUAL PAGE CONTENT' in messages[-1]['content']
                return {'content':'توضیح از منبع واقعی'}
        with tempfile.TemporaryDirectory() as directory:
            with TestClient(create_app(llm=Model(),workspace_path=Path(directory)/'db')) as client:
                client.get('/');key=client.post('/api/conversations').json()['id']
                with patch('app.main.fetch_public_text',return_value=('ACTUAL PAGE CONTENT','https://docs.ollama.com/windows')) as fetch:
                    r=client.post('/api/chat',json={'conversation_id':key,'text':'این لینک چی میگه https://docs.ollama.com/windows'})
                    events=[json.loads(x) for x in r.text.splitlines()]
                    self.assertEqual(events[-1]['stage'],'final')
                    self.assertEqual(events[-1]['evidence'][0]['tool'],'read_public_link')
                    fetch.assert_called_once()
                with patch('app.main.fetch_public_text',side_effect=ServiceError('link_failed','خوانده نشد')):
                    r=client.post('/api/chat',json={'conversation_id':key,'text':'https://docs.ollama.com/windows'})
                    self.assertEqual(json.loads(r.text.splitlines()[-1])['code'],'link_failed')
