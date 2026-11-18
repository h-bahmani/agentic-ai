"""Fail-closed request scope and numeric provenance checks, outside model prompts."""
import json
import re
from decimal import Decimal, InvalidOperation
from datetime import date
from .llm import ServiceError
from .routing import instruction

DIGITS=str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩','01234567890123456789')

def city_constraint(text):
    hint=re.search(r'\nشرط مکانی سؤال: شهر ([^\n]+)$',text)
    if hint:return hint[1]
    value=instruction(text)
    for match in re.finditer(r'(?:شهر|ساکن)\s+([\u0621-\u06ff]+)',value):
        if re.search(r'(?:کدام|کدوم|چه|هر|همه|تمام)\s*$',value[:match.start()]):continue
        if match[1] in {'بیشترین','کمترین','بالاترین','به','از','را','با','ها','های'}:continue
        return match[1]
    return None

async def verified_request_scope(text,tools,catalog,emit):
    """Resolve city names from live company dimensions, including implicit names."""
    names={d['function']['name'] for d in catalog}
    if not {'business_semantics','database_distinct'}<=names:return text,[]
    evidence=[]
    async def call(name,args):
        tools.validate(name,args);identity='scope-'+str(len(evidence)+1)
        await emit({'stage':'tool','id':identity,'status':'running','tool':name,'arguments':args,'operation':'validate request dimensions'})
        result,duration=await tools.execute(name,args)
        event={'stage':'tool','id':identity,'status':'success','tool':name,'arguments':args,'result':result,'duration_ms':duration}
        evidence.append(event);await emit(event)
        return result.get('data',result)
    semantics=await call('business_semantics',{})
    if re.search(r'مشتری(?:ان)?\s+فعال\b',text) and semantics.get('active_customer') is None:
        raise ServiceError('business_definition_missing','تعریف مشتری فعال در شرکت تنظیم نشده است؛ آن را حدس نمی‌زنم.')
    try:region=semantics['entities']['regions'];schema=semantics['default_schema'];citycol=region['columns']['city']
    except (KeyError,TypeError) as exc:raise ServiceError('scope_unverified','نگاشت شهر شرکت برای کنترل دامنه آماده نیست.') from exc
    values=await call('database_distinct',{'schema':schema,'name':region['table'],'column':citycol,'limit':200})
    if not isinstance(values,list) or len(values)>=200:raise ServiceError('scope_unverified','کامل‌بودن فهرست شهرها تأیید نشد؛ شرط شهر حدس زده نمی‌شود.')
    cities=[row.get('value') for row in values if isinstance(row,dict) and isinstance(row.get('value'),str)]
    normalized=instruction(text).replace('ي','ی').replace('ك','ک')
    hits=[city for city in cities if re.search(r'(?<!\w)'+re.escape(city.replace('ي','ی').replace('ك','ک'))+r'(?!\w)',normalized)]
    explicit=city_constraint(text)
    if explicit and not hits:raise ServiceError('unknown_city','نام شهر در دامنه واقعی شرکت تأیید نشد؛ گزارش صفر یا گزارش کل به جای آن داده نمی‌شود.')
    if len(hits)>1:return text+'\nشهرهای تأییدشده سؤال: '+json.dumps(hits,ensure_ascii=False),evidence
    if hits:return text+'\nشرط مکانی سؤال: شهر '+hits[0],evidence
    return text,evidence

def scoped_sales_request(text):
    value=instruction(text)
    city=city_constraint(value)
    if not city or not re.search(r'فروش|سفارش|خرید|درآمد',value):return None
    if not re.search(r'پرداخت.?شده|موفق|paid',value):return None
    if re.search(r'مقایسه|نسبت|درصد|تفکیک|بیشترین|کمترین|رتبه',value):return None
    metric='average' if re.search(r'میانگین|متوسط',value) else 'sum' if re.search(r'مبلغ|درآمد|مجموع|جمع',value) else 'count' if 'تعداد' in value else None
    if metric is None:return None
    return checked_report_plan(text,{'kind':'sales_summary','city':city,'metric':metric})

def checked_report_plan(text,plan):
    value=instruction(text).split('\nشرط مکانی سؤال:',1)[0]
    if len(re.findall(r'(?:شهر|ساکن)\s+[\u0621-\u06ff]+',value))>1 or re.search(r'لغو|ناموفق|refund|cancel|محصول|به.?تفکیک|مقایسه|نسبت|حداقل|حداکثر|بیشتر از|بیش از|کمتر از|دارای|بدون|به.?جز|غیر از',value,re.I):
        raise ServiceError('request_scope_unsupported','این گزارش شرطی دارد که طرح اجرایی فعلی آن را پوشش نمی‌دهد؛ شرط حذف و نتیجه کلی جایگزین نمی‌شود.')
    times=re.findall(r'\b\d{4}-\d{2}-\d{2}\b',value.translate(DIGITS))
    for item in times:
        try:date.fromisoformat(item)
        except ValueError as exc:raise ServiceError('request_date_invalid','تاریخ درخواست معتبر نیست.') from exc
    if len(times) not in (0,2):raise ServiceError('request_scope_unsupported','برای گزارش بازه‌دار، شروع و پایان را با تاریخ میلادی YYYY-MM-DD مشخص کنید؛ شرط زمانی حذف نمی‌شود.')
    if not times and re.search(r'امروز|دیروز|ماه|سال|هفته|تاریخ|بازه|از\s+\d',value.translate(DIGITS)):
        raise ServiceError('request_scope_unsupported','بازه زمانی این درخواست هنوز به تاریخ دقیق تبدیل نشده است؛ گزارش کل دوره جای آن داده نمی‌شود.')
    if times and times[0]>times[1]:raise ServiceError('request_date_invalid','شروع بازه از پایان آن بزرگ‌تر است.')
    channel=None
    if re.search(r'آنلاین|online',value,re.I):channel='online'
    elif re.search(r'کانال|نمایندگ|تلفنی|campaign|dealer|call_center',value,re.I):
        raise ServiceError('request_scope_unsupported','کانال این درخواست هنوز به مقدار تأییدشده نگاشت نشده است؛ شرط کانال حذف نمی‌شود.')
    # This verified report path accepts a controlled vocabulary. Unknown modifiers
    # must not be silently dropped merely because city/metric were recognized.
    covered=value.translate(DIGITS).replace('ي','ی').replace('ك','ک').replace('\u200c',' ')
    if plan.get('city'):covered=covered.replace(plan['city'],' ')
    covered=re.sub(r'\d{4}-\d{2}-\d{2}|\d+',' ',covered)
    allowed=set('کدام کدوم چه دارد شهرها شهرهای مجموع جمع کل مبلغ فروش فروشهای درآمد میانگین متوسط تعداد سفارش سفارشهای خرید خریدهای موفق موفقهای پرداخت شده paid آنلاین online در شهر ساکن از تا بین طی به بازه تاریخ میلادی چقدر چقدره چند چندتا چنده است هست هستند مشتری مشتریان را رو که با و یا برای هر شناسه نام کامل دارند داشته بیشترین برترین بالاترین کمترین رتبه ترتیب نزولی صعودی معرفی کن کنید نمایش بده بدهید نشان لطفا لطفاً یک دو سه چهار پنج شش هفت هشت نه ده های ها ی سوال سؤال ویس وویس فایل صوت متن پاسخ جواب'.split())
    unknown=set(re.findall(r'[^\W\d_]+',covered,re.UNICODE))-allowed
    if unknown:raise ServiceError('request_scope_unsupported','این درخواست واژه یا شرطی خارج از قرارداد گزارش تأییدشده دارد؛ برای جلوگیری از حذف شرط، پاسخی قطعی صادر نشد.')
    return {**plan,'dates':times,'channel':channel}

def validate_tool_scope(text,name,args):
    if city_constraint(text) and name not in {'business_semantics','database_relationship_graph','database_catalog','database_describe'}:
        raise ServiceError('unverified_scope','این ابزار شرط شهرِ درخواست را تضمین نمی‌کند؛ نتیجه کل شرکت به نام شهر نمایش داده نمی‌شود. برای این درخواست هنوز طرح اجراییِ تأییدشده نداریم.')
    if name.startswith('sales_'):
        value=instruction(text).translate(DIGITS)
        dates=re.findall(r'\b\d{4}-\d{2}-\d{2}\b',value)
        if len(dates)==2 and (args.get('start')!=dates[0] or args.get('end')!=dates[1]):
            raise ServiceError('unverified_scope','بازه زمانی سؤال در فراخوانی ابزار اعمال نشده است؛ اجرای پاسخ کلی مسدود شد.')
        if not dates and re.search(r'امروز|دیروز|این ماه|این سال|هفته گذشته|ماه گذشته|سال گذشته',value):
            raise ServiceError('unverified_scope','زمان نسبی هنوز به بازه تأییدشده تبدیل نشده است؛ گزارش کل دوره جایگزین نمی‌شود.')
        if re.search(r'آنلاین|online|کانال',value,re.I) and name!='sales_breakdown':
            raise ServiceError('unverified_scope','ابزار انتخاب‌شده شرط کانال فروش را اعمال نمی‌کند؛ اجرای نتیجه کلی مسدود شد.')
        if re.search(r'موفق|پرداخت.?شده|paid',value) and name=='sales_order_count' and args.get('kind')!='paid':
            raise ServiceError('unverified_scope','وضعیت موفق در شمارش سفارش‌ها اعمال نشده است.')

def validate_final(text,answer,events):
    if not events:
        if re.search(r'\d',answer.translate(DIGITS)):
            raise ServiceError('answer_unverified','مدل بدون نتیجه ابزار عددی اعلام کرد؛ پاسخ به عنوان نتیجه تأییدشده پذیرفته نشد.')
        return
    permitted=set()
    def collect(value):
        if isinstance(value,bool):return
        if isinstance(value,(int,float)):
            permitted.add(Decimal(str(value)))
        elif isinstance(value,dict):
            for child in value.values():collect(child)
        elif isinstance(value,list):
            for child in value:collect(child)
        elif isinstance(value,str):
            try:collect(json.loads(value))
            except (ValueError,RecursionError):pass
    for event in events:collect(event.get('result',{}))
    for token in re.findall(r'\d[\d,٬]*(?:[.٫]\d+)?',answer.translate(DIGITS)):
        try:number=Decimal(token.replace(',','').replace('٬','').replace('٫','.'))
        except InvalidOperation:continue
        if number not in permitted:
            raise ServiceError('answer_number_unverified','مدل عددی خارج از شواهد ابزار نوشت؛ پاسخ متوقف شد و مقدار حدسی نمایش داده نمی‌شود.')
    currency=None
    for event in events:
        if event.get('tool')=='business_semantics':
            value=event['result'];value=value.get('data',value)
            currency=value.get('revenue',{}).get('currency')
    for unit in ('تومان','ریال'):
        if unit in answer and (not currency or unit not in str(currency)):
            raise ServiceError('answer_currency_unverified','مدل واحد پول تأییدنشده اضافه کرد؛ مبلغ به ریال یا تومان نسبت داده نمی‌شود.')
