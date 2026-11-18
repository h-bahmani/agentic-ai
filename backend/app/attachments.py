"""Owner-scoped attachments and bounded public HTTPS text retrieval."""
import json
import re
import socket
import ssl
import http.client
import ipaddress
import zipfile
from io import BytesIO
from pathlib import Path
from urllib.parse import urlsplit,urljoin,quote
from html.parser import HTMLParser
from xml.etree import ElementTree
from .local_workspace import owner_key
from .llm import ServiceError
import secrets

class TextHTML(HTMLParser):
    def __init__(self):super().__init__();self.parts=[];self.skip=0
    def handle_starttag(self,tag,attrs):
        if tag in {'script','style'}:self.skip+=1
    def handle_endtag(self,tag):
        if tag in {'script','style'}:self.skip=max(0,self.skip-1)
    def handle_data(self,data):
        if not self.skip and data.strip():self.parts.append(data.strip())

def public_address(host):
    addresses=socket.getaddrinfo(host,443,type=socket.SOCK_STREAM)
    ips={a[4][0] for a in addresses}
    if not ips or any(not ipaddress.ip_address(ip).is_global for ip in ips):
        raise ServiceError('link_address','لینک محلی، خصوصی یا داخلی برای واکشی وب مجاز نیست.')
    return sorted(ips)[0]

def fetch_public_text(url):
    for _ in range(4):
        parts=urlsplit(url)
        if parts.scheme!='https' or not parts.hostname or parts.username or parts.password or parts.port not in (None,443):
            raise ServiceError('link_url','فقط لینک عمومی HTTPS بدون رمز و با پورت استاندارد پذیرفته می‌شود.')
        host=parts.hostname.encode('idna').decode('ascii');ip=public_address(host)
        class PinnedConnection(http.client.HTTPSConnection):
            def connect(self):
                raw=socket.create_connection((ip,443),timeout=8)
                try:self.sock=self._context.wrap_socket(raw,server_hostname=host)
                except BaseException:raw.close();raise
        connection=PinnedConnection(host,timeout=8,context=ssl.create_default_context())
        try:
            target=quote(parts.path or '/',safe='/%:@')
            if parts.query:target+='?'+quote(parts.query,safe='=&%:+,;/@?')
            connection.request('GET',target,headers={'User-Agent':'MadarLocalAgent/0.1','Accept':'text/html,text/plain,application/json','Accept-Encoding':'identity'})
            response=connection.getresponse()
            if response.status in (301,302,303,307,308):
                location=response.getheader('Location')
                if not location:raise ServiceError('link_failed','هدایت لینک مقصد ندارد.')
                url=urljoin(url,location);continue
            if response.status!=200:raise ServiceError('link_failed','سایت پاسخ موفق نداد؛ محتوایی حدس زده نشد.')
            mime=response.getheader('Content-Type','').split(';')[0].lower()
            if mime not in {'text/html','text/plain','application/json'} or response.getheader('Content-Encoding','identity') not in ('identity',''):
                raise ServiceError('link_type','فقط متن و HTML معمولی قابل خواندن است؛ سایت نیازمند ورود یا اجرای JavaScript پشتیبانی نمی‌شود.')
            content=response.read(1024*1024+1)
            if len(content)>1024*1024:raise ServiceError('link_size','صفحه از سقف یک مگابایت بیشتر است.')
            charset=response.headers.get_content_charset() or 'utf-8'
            text=content.decode(charset,errors='strict')
            if mime=='text/html':parser=TextHTML();parser.feed(text);text='\n'.join(parser.parts)
            if not text.strip() or len(text)>100000:raise ServiceError('link_text','متن صفحه خالی یا بیش از ظرفیت است.')
            return text,url
        finally:connection.close()
    raise ServiceError('link_redirects','تعداد هدایت‌های لینک بیش از حد مجاز است.')

class Attachments:
    def __init__(self,workspace):
        self.workspace=workspace
        with workspace.connect() as db:
            db.execute('CREATE TABLE IF NOT EXISTS attachments (id TEXT PRIMARY KEY,owner TEXT NOT NULL,name TEXT NOT NULL,kind TEXT NOT NULL,text TEXT NOT NULL,payload BLOB NOT NULL,mime TEXT NOT NULL)')
            if 'metadata' not in {row[1] for row in db.execute('PRAGMA table_info(attachments)')}:
                db.execute("ALTER TABLE attachments ADD COLUMN metadata TEXT NOT NULL DEFAULT '{}'")

    def add(self,owner,name,content,link_text=None):
        name=name[:200] if link_text is not None else Path(name.replace('\\','/')).name[:200];extension=Path(name).suffix.lower()
        if len(content)>5*1024**2 or (not content and link_text is None):raise ServiceError('attachment_size','حداکثر حجم هر پیوست ۵ مگابایت است و فایل نباید خالی باشد.')
        text='';kind='text';mime='application/octet-stream';metadata={}
        if link_text is not None:text=link_text;kind='link'
        elif extension in {'.txt','.md','.csv','.json','.log','.ics'}:
            try:text=content.decode('utf-8-sig')
            except UnicodeError as exc:raise ServiceError('attachment_encoding','فایل متنی باید UTF-8 باشد.') from exc
            if extension=='.ics':
                metadata['meetings']=self.workspace.import_calendar(owner,text);kind='calendar'
        elif extension=='.docx':
            try:
                with zipfile.ZipFile(BytesIO(content)) as archive:
                    if sum(z.file_size for z in archive.infolist())>10*1024**2:raise ValueError()
                    xml=archive.read('word/document.xml')
                    text='\n'.join(ElementTree.fromstring(xml).itertext())
            except Exception as exc:raise ServiceError('attachment_invalid','فایل DOCX معتبر و در محدوده حجم لازم است.') from exc
        elif extension=='.eml':
            imported=self.workspace.import_mail(owner,content);mail=self.workspace.read_mail(owner,imported['id'])
            text=f"موضوع: {mail['subject']}\nفرستنده: {mail['sender']}\n{mail['body']}"
            metadata['mail_id']=mail['id'];kind='mail'
            if imported.get('meeting_results'):metadata['meetings']=imported['meeting_results'];kind='calendar'
        elif extension in {'.png','.jpg','.jpeg','.webp'}:
            valid=(content.startswith(b'\x89PNG\r\n\x1a\n') if extension=='.png' else content.startswith(b'\xff\xd8\xff') if extension in {'.jpg','.jpeg'} else content.startswith(b'RIFF') and content[8:12]==b'WEBP')
            if not valid:raise ServiceError('attachment_invalid','امضای فایل تصویر معتبر نیست.')
            kind='image';mime={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.webp':'image/webp'}[extension]
        elif extension=='.pdf' and content.startswith(b'%PDF-'):kind='unavailable';text='PDF دریافت شد، اما استخراج متن PDF در محیط فعلی نصب نیست.'
        else:raise ServiceError('attachment_type','فرمت‌های فعلی: TXT، MD، CSV، JSON، DOCX، EML، ICS، PDF و PNG/JPEG/WebP.')
        if len(text)>100000:raise ServiceError('attachment_text_size','متن پیوست بیش از ظرفیت است؛ چیزی پنهانی حذف نمی‌شود.')
        with self.workspace.connect() as db:
            if db.execute('SELECT COUNT(*) FROM attachments').fetchone()[0]>=100 or db.execute('SELECT COALESCE(SUM(LENGTH(payload)),0) FROM attachments').fetchone()[0]+len(content)>50*1024**2:
                raise ServiceError('attachment_storage','ظرفیت پیوست محلی پر است.')
            key=secrets.token_urlsafe(18)
            db.execute('INSERT INTO attachments (id,owner,name,kind,text,payload,mime,metadata) VALUES (?,?,?,?,?,?,?,?)',(key,owner_key(owner),name,kind,text,content,mime,json.dumps(metadata,ensure_ascii=False)))
        return {'id':key,'name':name,'kind':kind,'mime':mime,'text_characters':len(text),'metadata':metadata}

    def get(self,owner,key):
        with self.workspace.connect() as db:row=db.execute('SELECT * FROM attachments WHERE id=? AND owner=?',(key,owner_key(owner))).fetchone()
        if not row:raise ServiceError('attachment_not_found','پیوست متعلق به این مرورگر نیست یا وجود ندارد.')
        result=dict(row);result['metadata']=json.loads(result['metadata']);return result
