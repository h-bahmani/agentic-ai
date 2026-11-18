"""Read-only real MCP probe. Never prints email body, sender or subject."""
import asyncio
from datetime import datetime,timedelta,timezone
from app.outlook import connect_outlook,recent_mail_range
from app.llm import ServiceError

async def main():
    async with connect_outlook() as tools:
        catalog=await tools.discover()
        try:
            await tools.execute('outlook_search_emails',{'limit':1000})
            raise AssertionError('invalid limit accepted')
        except ServiceError as exc:
            assert exc.code=='tool_arguments_invalid'
            print({'invalid_limit':exc.code})
        dated,_=await tools.execute('outlook_search_emails',{'query':'','limit':10,**recent_mail_range('یه هفته پیش')})
        print({'real_last_7_days_count':len(dated.get('data',dated)['items'])})
        status,_=await tools.execute('outlook_status',{})
        found,_=await tools.execute('outlook_search_emails',{'query':'','limit':2})
        payload=found.get('data',found)
        rows=payload.get('items',[])
        body_chars=0
        if rows:
            email,_=await tools.execute('outlook_read_email',{'message_id':rows[0]['id']})
            body_chars=len(email.get('data',email).get('body',''))
            first,_=await tools.execute('outlook_read_email',{'message_id':rows[0]['id'],'offset':0,'body_limit':80})
            p=first.get('data',first)
            assert p['body']==email.get('data',email)['body'][:80]
            if p['next_offset'] is not None:
                second,_=await tools.execute('outlook_read_email',{'message_id':rows[0]['id'],'offset':p['next_offset'],'body_limit':80})
                assert second.get('data',second)['body']==email.get('data',email)['body'][80:160]
            print({'real_body_pagination':True})
        page1,_=await tools.execute('outlook_search_emails',{'query':'','limit':2,'offset':0,'folder':'inbox'})
        page2,_=await tools.execute('outlook_search_emails',{'query':'','limit':2,'offset':2,'folder':'inbox'})
        ids1={r['id'] for r in page1.get('data',page1)['items']}
        ids2={r['id'] for r in page2.get('data',page2)['items']}
        assert not ids1.intersection(ids2)
        print({'real_search_pagination':True})
        now=datetime.now(timezone.utc)
        calendar,_=await tools.execute('outlook_calendar',{'start_utc':now.isoformat(),'end_utc':(now+timedelta(days=7)).isoformat(),'limit':20})
        print({'transport':'real stdio MCP','tools':len(catalog),'status':status,'search_count':len(rows),
               'read_body_characters':body_chars,'calendar_count':len(calendar.get('data',calendar).get('items',[]))})

if __name__=='__main__':asyncio.run(main())
