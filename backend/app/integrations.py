"""Future integration slots. Missing setup never produces mocked execution."""
from typing import Protocol

class RemoteModelProvider(Protocol):
    async def health(self) -> dict: ...
    async def chat(self,messages:list[dict],tools:list[dict]) -> dict: ...

class MailProvider(Protocol):
    async def search(self,user_id:str,query:str) -> list[dict]: ...
    async def read(self,user_id:str,message_id:str) -> dict: ...

class NotificationChannel(Protocol):
    async def deliver(self,user_id:str,reminder_id:str,text:str) -> str: ...

def integration_status():
    return {
      'company_gpu':{'state':'not_configured','connected':False,'required':['base_url','model','authentication','allowed_host'],'message':'محل اتصال API شرکت آمادهٔ توسعه است؛ adapter اجرایی GPU هنوز متصل نیست.'},
      'outlook':{'state':'not_configured','connected':False,'required':['Microsoft365_or_Exchange','OAuth_registration','per_user_consent'],'message':'هیچ حساب Outlook متصل نیست. ایمیل محلی فقط EML واردشده توسط کاربر است.'},
      'telegram':{'state':'not_configured','connected':False,'required':['bot_token_secret','authorized_chat_binding'],'message':'ارسال Telegram پیاده‌سازی/فعال نشده؛ یادآوری فعلاً داخل MobinAid نمایش داده می‌شود.'},
      'local_reminders':{'state':'enabled','connected':True,'delivery':'in_app','persistence':'dedicated SQLite','message':'در زمان اجرای سرویس فعال است؛ پس از خاموشی، موارد عقب‌افتاده در راه‌اندازی بعدی موعدرسیده می‌شوند.'}}
