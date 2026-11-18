"""Current-request routing; previous tool results never determine a new intent."""
import re

def instruction(text: str) -> str:
    parts=text.split('\n\nمتن فایل صوتی:\n',1)
    head=parts[0].strip()
    if len(parts)==1:return head
    if re.search(r'خلاصه|ترجمه|بازنویسی|ابزار.*اجرا نکن|بدون.*ابزار',head):return head
    return head+'\n'+parts[1]

def route_request(text: str) -> str:
    value=instruction(text).lower().replace('ي','ی').replace('ك','ک')
    provided_text='\n\nمتن فایل صوتی:\n' in text
    head=text.split('\n\nمتن فایل صوتی:\n',1)[0].lower()
    if re.search(r'(ابزار|mcp).*(اجرا|استفاده).*نکن|بدون.*ابزار',head):return 'chat'
    if (provided_text and re.search(r'خلاصه|بازنویسی|ترجمه|توضیح|متن',head)) or re.search(r'(ترجمه|بازنویسی|متن.*ویس|متن.*صوت|این (ویس|صوت|متن).*خلاصه)',value):return 'chat'
    if re.search(r'\b(?:sales|customer|finance|database|quality|support|marketing|subscriptions)_\w+',value):return 'company'
    if re.search(r'(تو.*چی.*کار|چه کار.*می.?کن|چه کاری.*می.?کن|چه قابلیت|چه امکانات|چه ابزار|کی هستی|به چی وصل|چیکار.*می.?کنی|چی کار.*می.?کنی)',value):return 'capabilities'
    if re.search(r'چیست|چیه|تفاوت|فرق|توضیح بده',value) and not re.search(r'شرکت|داده|سامانه|سیستم',value):return 'chat'
    if re.search(r'(فروش|سفارش|مشتری|درآمد|پرداخت|اشتراک|کمپین|تیکت|پایگاه داده|دیتابیس|جدول|ستون|رکورد|داده.?های شرکت|آمار شرکت)',value):return 'company'
    if re.search(r'(سلام|درود|ممنون|تشکر|ویس تست|وویس تست|تست.*صوت|تستی.*می.?گیر|لطیفه|شعر)',value):return 'chat'
    return 'auto'

CAPABILITIES = ('می‌توانم گفت‌وگو کنم، متن را توضیح بدهم، خلاصه یا ترجمه کنم و فایل صوتی را با Whisper به متن فارسی تبدیل کنم. '
                'برای پرسش‌های مربوط به اطلاعات شرکت از ابزارهای مجاز MCP در حوزه‌های فروش، مشتریان، مالی، اشتراک‌ها، پشتیبانی، بازاریابی و کیفیت داده استفاده می‌کنم. '
                'برای گفت‌وگوی معمولی یا توضیح و خلاصه‌سازی، ابزار شرکت اجرا نمی‌کنم. می‌توانی فایل صوتی را همراه با توضیح دلخواهت بفرستی. '
                'در بخش «اتصال‌ها» می‌توانی Outlook کلاسیک همین رایانه را متصل کنی؛ پس از اتصال، جست‌وجوی ایمیل و تقویم و یادآوری محلی جلسه‌ها فعال می‌شود. Telegram متصل نیست.')

def relevant_tools(text: str, definitions: list[dict]) -> list[dict]:
    value=instruction(text).lower()
    explicit={d['function']['name'] for d in definitions if d['function']['name'] in value}
    if explicit:
        return [d for d in definitions if d['function']['name'] in explicit]
    domains=[('sales_',r'فروش|سفارش|درآمد'),('customer_',r'مشتری'),
             ('finance_',r'پرداخت|مالی|وصول'),('subscriptions_',r'اشتراک|تمدید'),
             ('support_',r'تیکت|پشتیبانی|رسیدگی'),('marketing_',r'کمپین|بازاریابی|تبلیغ'),
             ('quality_',r'کیفیت|تکراری|تهی|یتیم'),('database_',r'جدول|ستون|دیتابیس|پایگاه|رابطه|رکورد|اسکیما')]
    prefixes=[p for p,pattern in domains if re.search(pattern,value)]
    selected=[d for d in definitions if any(d['function']['name'].startswith(p) for p in prefixes)]
    return selected
