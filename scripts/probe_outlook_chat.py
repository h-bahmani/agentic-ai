"""Real HTTP-handler → stdio MCP → Outlook integration probe; no email text logged."""
import json
import tempfile
import sys
from pathlib import Path
from fastapi.testclient import TestClient
from app.main import create_app

with tempfile.TemporaryDirectory() as directory:
    with TestClient(create_app(workspace_path=Path(directory)/'probe.sqlite3')) as client:
        client.get('/')
        connection=client.post('/api/outlook/connect')
        assert connection.status_code==200,connection.status_code
        assert connection.json()['connected'] is True
        key=client.post('/api/conversations').json()['id']
        for question in ['من ایمیلی امروز داشتم؟','خوب ایمیلی که در یه هفته پیش به من زده شده چی بوده','ایمیل‌های خوانده‌نشده من در ۷ روز اخیر را نمایش بده.']:
            response=client.post('/api/chat',json={'conversation_id':key,'text':question})
            events=[json.loads(line) for line in response.text.splitlines() if line]
            errors=[e for e in events if e['stage']=='error']
            assert not errors,[e['code'] for e in errors]
            final=next(e for e in events if e['stage']=='final')
            assert final['answer'] and final['evidence']
            result=final['evidence'][0]['result']
            payload=result.get('data',result)
            if 'خوانده' in question:
                assert payload['unread'] is True
                assert all(row['unread'] is True for row in payload['items'])
            assert all(len(row['id'])<15 for row in payload['items'])
            print({'question':question,'final':True,'actual_mcp':True,'folder':payload['folder'],
                   'returned':len(payload['items']),'matching_scanned':payload['matching_scanned'],
                   'scan_complete':payload['scan_complete'],'has_more':payload['possibly_more']})
        if '--with-model' in sys.argv:
            response=client.post('/api/chat',json={'conversation_id':key,'text':'آخرین ایمیل دریافتی را پیدا کن، متنش را بخوان و موضوعش را بگو.'})
            events=[json.loads(line) for line in response.text.splitlines() if line]
            errors=[e for e in events if e['stage']=='error']
            assert not errors,[e['code'] for e in errors]
            final=next(e for e in events if e['stage']=='final')
            reads=[e for e in final['evidence'] if e.get('tool')=='outlook_read_email']
            assert reads,'No real email body read'
            print({'real_read_workflow_and_mcp':True,'read_calls':len(reads),'final_received':bool(final['answer']),
                   'model_generation_required':False})
        client.post('/api/outlook/disconnect')
