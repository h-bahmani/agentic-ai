"""Durable local reminders and explicitly imported EML messages; no cloud access."""
import hashlib
import secrets
import sqlite3
from pathlib import Path
from datetime import datetime, timezone, timedelta
from email import policy
from email.parser import BytesParser
from contextlib import contextmanager
from .llm import ServiceError

def utcnow():return datetime.now(timezone.utc)
def owner_key(owner):return hashlib.sha256(owner.encode()).hexdigest()

class LocalWorkspace:
    def __init__(self,path):
        self.path=Path(path);self.path.parent.mkdir(parents=True,exist_ok=True)
        with self.connect() as db:
            db.executescript('''
              CREATE TABLE IF NOT EXISTS local_mail (
                id TEXT PRIMARY KEY, owner TEXT NOT NULL, subject TEXT NOT NULL,
                sender TEXT NOT NULL, sent_at TEXT NOT NULL, body TEXT NOT NULL);
              CREATE INDEX IF NOT EXISTS local_mail_owner ON local_mail(owner);
              CREATE TABLE IF NOT EXISTS reminders (
                id TEXT PRIMARY KEY, owner TEXT NOT NULL, title TEXT NOT NULL,
                due_at TEXT NOT NULL, status TEXT NOT NULL, created_at TEXT NOT NULL,
                request_key TEXT NOT NULL, mail_id TEXT,
                UNIQUE(owner,request_key));
              CREATE INDEX IF NOT EXISTS reminders_due ON reminders(status,due_at);
            ''')

    @contextmanager
    def connect(self):
        db=sqlite3.connect(self.path,timeout=5);db.row_factory=sqlite3.Row
        try:
            with db:yield db
        finally:db.close()

    def add_reminder(self,owner,title,due_at,request_key,mail_id=None):
        if due_at.tzinfo is None:raise ServiceError('reminder_timezone','زمان یادآوری باید منطقه زمانی داشته باشد.')
        owner=owner_key(owner)
        with self.connect() as db:
            existing=db.execute('SELECT * FROM reminders WHERE owner=? AND request_key=?',(owner,request_key)).fetchone()
            if existing:
                if existing['title']!=title or existing['due_at']!=due_at.astimezone(timezone.utc).isoformat() or existing['mail_id']!=mail_id:
                    raise ServiceError('idempotency_conflict','شناسه درخواست قبلاً با اطلاعات دیگری استفاده شده است.')
                return self.public(existing)
            if due_at.tzinfo is None:raise ServiceError('reminder_timezone','زمان یادآوری باید منطقه زمانی داشته باشد.')
            now=utcnow();due=due_at.astimezone(timezone.utc)
            if due<=now or due>now+timedelta(days=366):raise ServiceError('reminder_date','زمان باید در آینده و حداکثر یک سال بعد باشد.')
            if db.execute('SELECT COUNT(*) FROM reminders').fetchone()[0]>=1000:raise ServiceError('storage_limit','ظرفیت یادآوری‌های نسخه محلی پر شده است.')
            if mail_id and not db.execute('SELECT 1 FROM local_mail WHERE id=? AND owner=?',(mail_id,owner)).fetchone():
                raise ServiceError('mail_not_found','ایمیل مرتبط متعلق به این نشست نیست یا وجود ندارد.')
            key=secrets.token_urlsafe(18)
            db.execute('INSERT INTO reminders VALUES (?,?,?,?,?,?,?,?)',
                       (key,owner,title,due.isoformat(),'pending',now.isoformat(),request_key,mail_id))
            return self.public(db.execute('SELECT * FROM reminders WHERE id=?',(key,)).fetchone())

    @staticmethod
    def public(row):return {k:row[k] for k in row.keys() if k not in {'owner','request_key'}}

    def tick(self,now=None):
        with self.connect() as db:
            return db.execute("UPDATE reminders SET status='due' WHERE status='pending' AND due_at<=?",((now or utcnow()).astimezone(timezone.utc).isoformat(),)).rowcount

    def reminder_by_request(self,owner,key):
        with self.connect() as db:row=db.execute('SELECT * FROM reminders WHERE owner=? AND request_key=?',(owner_key(owner),key)).fetchone()
        return self.public(row) if row else None

    def sync_outlook_reminders(self,owner,meetings,now):
        key=owner_key(owner)
        with self.connect() as db:
            keys=['outlook:'+hashlib.sha256((m['id']+m['start'].isoformat()).encode()).hexdigest() for m in meetings]
            existing=db.execute("SELECT request_key FROM reminders WHERE owner=? AND request_key LIKE 'outlook:%' AND status IN ('pending','due')",(key,)).fetchall()
            for row in existing:
                if row['request_key'] not in keys:db.execute('DELETE FROM reminders WHERE owner=? AND request_key=?',(key,row['request_key']))
            for meeting in meetings:
                request_key='outlook:'+hashlib.sha256((meeting['id']+meeting['start'].isoformat()).encode()).hexdigest()
                due=max(meeting['due'],now)
                db.execute('INSERT OR IGNORE INTO reminders VALUES (?,?,?,?,?,?,?,?)',
                    (secrets.token_urlsafe(18),key,'جلسه Outlook: '+meeting['title'][:260],
                     due.isoformat(),'pending',now.isoformat(),request_key,None))

    def import_calendar(self,owner,text,mail_id=None):
        from .calendar_local import calendar_reminders
        return calendar_reminders(self,owner,text,mail_id)

    def reminders(self,owner):
        with self.connect() as db:
            return [self.public(r) for r in db.execute("SELECT * FROM reminders WHERE owner=? ORDER BY CASE status WHEN 'due' THEN 0 WHEN 'pending' THEN 1 ELSE 2 END,due_at ASC LIMIT 100",(owner_key(owner),))]

    def acknowledge(self,owner,key):
        with self.connect() as db:
            row=db.execute('SELECT * FROM reminders WHERE id=? AND owner=?',(key,owner_key(owner))).fetchone()
            if not row:raise ServiceError('reminder_not_found','یادآوری پیدا نشد.')
            if row['status']!='due':raise ServiceError('reminder_not_due','یادآوری هنوز موعدش نرسیده یا قبلاً دیده شده است.')
            db.execute("UPDATE reminders SET status='seen' WHERE id=? AND owner=?",(key,owner_key(owner)))
            return {'id':key,'status':'seen'}

    def import_mail(self,owner,content):
        if not content or len(content)>512*1024:raise ServiceError('mail_size','حداکثر حجم فایل EML محلی ۵۱۲ کیلوبایت است.')
        try:
            mail=BytesParser(policy=policy.default).parsebytes(content)
            subject=str(mail.get('subject','')).strip();sender=str(mail.get('from','')).strip()
            part=mail.get_body(preferencelist=('plain',))
            if part is None:raise ValueError()
            body=part.get_content()
            if not isinstance(body,str) or not body.strip() or not sender or not subject:raise ValueError()
            if len(body)>100000 or len(subject)>500 or len(sender)>500:raise ValueError()
        except Exception as exc:raise ServiceError('mail_invalid','فایل باید EML معتبر با موضوع، فرستنده و بدنه متن ساده باشد؛ HTML و پیوست اجرا نمی‌شوند.') from exc
        with self.connect() as db:
            if db.execute('SELECT COUNT(*) FROM local_mail').fetchone()[0]>=500:raise ServiceError('storage_limit','ظرفیت ایمیل محلی پر است.')
            key=secrets.token_urlsafe(18)
            db.execute('INSERT INTO local_mail VALUES (?,?,?,?,?,?)',(key,owner_key(owner),subject,sender,str(mail.get('date',''))[:200],body))
        meetings=[]
        for part in mail.walk():
            if part.get_content_type()=='text/calendar':
                try:meetings+=self.import_calendar(owner,part.get_content(),key)
                except Exception:meetings.append({'status':'not_scheduled','reason':'دعوت‌نامه قابل پردازش نیست.'})
        return {'id':key,'subject':subject,'sender':sender,'source':'local_eml','meeting_results':meetings}

    def search_mail(self,owner,query):
        escaped=query.replace('!','!!').replace('%','!%').replace('_','!_')
        value='%'+escaped+'%'
        with self.connect() as db:
            rows=db.execute("SELECT * FROM local_mail WHERE owner=? AND (subject LIKE ? ESCAPE '!' OR sender LIKE ? ESCAPE '!' OR body LIKE ? ESCAPE '!') LIMIT 21",(owner_key(owner),value,value,value)).fetchall()
            return {'items':[{'id':r['id'],'subject':r['subject'],'sender':r['sender'],'sent_at':r['sent_at'],'excerpt':r['body'][:700],'source':'local_eml'} for r in rows[:20]],
                    'truncated':len(rows)>20,'scope':'فقط فایل‌های EML واردشده توسط این نشست؛ جستجوی واژه‌ای، نه صندوق Outlook'}

    def read_mail(self,owner,key):
        with self.connect() as db:
            row=db.execute('SELECT * FROM local_mail WHERE owner=? AND id=?',(owner_key(owner),key)).fetchone()
            if not row:raise ServiceError('mail_not_found','ایمیل پیدا نشد.')
            return {**self.public(row),'source':'local_eml'}
