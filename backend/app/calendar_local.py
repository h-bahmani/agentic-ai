"""Strict, one-off iCalendar invitations. No dates guessed from email prose."""
import re
import hashlib
from datetime import datetime,timezone,timedelta
from zoneinfo import ZoneInfo
from .llm import ServiceError

def calendar_reminders(workspace,owner,content,mail_id=None):
    lines=re.sub(r'\r?\n[ \t]','',content).replace('\r','').split('\n')
    if any(line.strip() in {'METHOD:CANCEL','STATUS:CANCELLED','RRULE:'} or line.startswith('RRULE:') for line in lines):
        return [{'status':'not_scheduled','reason':'لغو یا تکرار تقویم در این نسخه خودکار پشتیبانی نمی‌شود.'}]
    blocks=re.findall(r'BEGIN:VEVENT\n(.*?)\nEND:VEVENT','\n'.join(lines),re.S)
    if len(blocks)>10:raise ServiceError('calendar_limit','حداکثر ۱۰ جلسه در یک فایل قابل بررسی است.')
    results=[]
    for block in blocks:
        fields={}
        for line in block.split('\n'):
            key,sep,value=line.partition(':')
            if sep:fields[key]=value
        start=next(((k,v) for k,v in fields.items() if k.split(';')[0]=='DTSTART'),None)
        uid=fields.get('UID');title=fields.get('SUMMARY')
        if not start or not uid or not title:
            results.append({'status':'not_scheduled','reason':'شناسه، عنوان یا زمان شروع جلسه کامل نیست.'});continue
        try:
            key,value=start
            if value.endswith('Z'):when=datetime.strptime(value,'%Y%m%dT%H%M%SZ').replace(tzinfo=timezone.utc)
            else:
                zone=re.search(r'TZID=([^;]+)',key)
                if not zone:raise ValueError()
                when=datetime.strptime(value,'%Y%m%dT%H%M%S').replace(tzinfo=ZoneInfo(zone[1]))
            now=datetime.now(timezone.utc);when=when.astimezone(timezone.utc)
            if when<=now:
                results.append({'status':'not_scheduled','reason':'این جلسه در گذشته است.','title':title});continue
            due=max(when-timedelta(minutes=15),now+timedelta(seconds=2))
            request='meeting-'+hashlib.sha256((uid+when.isoformat()).encode()).hexdigest()
            old=workspace.reminder_by_request(owner,request)
            saved=old or workspace.add_reminder(owner,('جلسه: '+title)[:300],due,request,mail_id)
            results.append({'status':'scheduled','title':saved['title'],'reminder_id':saved['id'],'meeting_start_utc':when.isoformat(),'reminder_at_utc':saved['due_at'],'destination':'chat'})
        except ServiceError as exc:
            results.append({'status':'not_scheduled','reason':exc.message,'title':title})
        except Exception:
            results.append({'status':'not_scheduled','reason':'زمان یا منطقه زمانی جلسه قابل تأیید نیست؛ زمانی حدس زده نشد.','title':title})
    return results
