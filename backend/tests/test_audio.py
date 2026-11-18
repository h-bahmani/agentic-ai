import asyncio
import tempfile
import unittest
from pathlib import Path
from fastapi.testclient import TestClient
from app.main import create_app
from app.config import Settings
from app.speech import audio_signature
from app.llm import ServiceError
from test_core import FakeLLM, FakeMCP

class FakeSpeech:
    def __init__(self,root):self.upload_root=root;self.lock=asyncio.Lock();self.calls=[]
    def configuration(self):return {}
    def health(self):return {'available':True}
    async def close(self):pass
    async def transcribe(self,path):
        self.calls.append(path.read_bytes())
        return {'text':'تعداد سفارش‌ها چقدر است؟','language':'fa','model':'test-double'}

class AudioTests(unittest.TestCase):
    def test_low_ram_handoff_releases_only_configured_model(self):
        class LLM(FakeLLM):
            released=False
            async def release_for_speech(self):self.released=True
        class Speech(FakeSpeech):
            attempts=0
            closed=False
            async def transcribe(self,path):
                self.attempts+=1
                if self.attempts==1:raise ServiceError('stt_failed','RAM آزاد کافی نیست: 100 MB')
                return await super().transcribe(path)
            async def close(self):self.closed=True
        with tempfile.TemporaryDirectory() as temp:
            llm=LLM([]);speech=Speech(Path(temp)/'uploads')
            with TestClient(create_app(Settings(),llm,FakeMCP(),speech)) as c:
                c.get('/');key=c.post('/api/conversations').json()['id']
                r=c.post('/api/audio',params={'conversation_id':key,'extension':'.wav'},content=b'RIFF1234WAVEdata')
                self.assertEqual(r.status_code,200)
                self.assertEqual(r.json()['resource_mode'],'sequential_ram')
                self.assertTrue(llm.released)
                self.assertTrue(speech.closed)
                self.assertEqual(list(speech.upload_root.iterdir()),[])

    def test_signatures(self):
        self.assertTrue(audio_signature(b'RIFF1234WAVEdata','.wav'))
        self.assertTrue(audio_signature(b'OggS1234','.ogg'))
        self.assertTrue(audio_signature(b'ID3abc','.mp3'))
        self.assertFalse(audio_signature(b'<html>','.mp3'))

    def test_upload_access_validation_cleanup_and_transcript(self):
        with tempfile.TemporaryDirectory() as temp:
            speech=FakeSpeech(Path(temp)/'uploads')
            with TestClient(create_app(Settings(),FakeLLM([]),FakeMCP(),speech)) as c:
                c.get('/')
                key=c.post('/api/conversations').json()['id']
                url='/api/audio?conversation_id='+key+'&extension=.wav'
                self.assertEqual(c.post(url,content=b'fake').status_code,400)
                self.assertEqual(speech.calls,[])
                r=c.post(url,content=b'RIFF1234WAVEdata')
                self.assertEqual(r.json()['text'],'تعداد سفارش‌ها چقدر است؟')
                self.assertEqual(len(speech.calls),1)
                self.assertEqual(list(speech.upload_root.iterdir()),[])
                self.assertEqual(c.post(url.replace('.wav','.exe'),content=b'x').status_code,400)
                with TestClient(c.app) as other:
                    other.get('/')
                    self.assertEqual(other.post(url,content=b'RIFF1234WAVEdata').status_code,400)

    def test_audio_respects_origin_and_auth(self):
        with TestClient(create_app(Settings(token='secret'),FakeLLM([]),FakeMCP())) as c:
            self.assertEqual(c.post('/api/audio?conversation_id=123&extension=.wav',content=b'x').status_code,401)

if __name__=='__main__':unittest.main()
