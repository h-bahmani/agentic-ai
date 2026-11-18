"""Dimension rankings and zero-history queries with data-derived boundaries."""
import re
import json
from decimal import Decimal
from datetime import date
from .grounding import checked_report_plan, city_constraint, DIGITS
from .llm import ServiceError

def set_report_request(text):
    normalized=text.translate(DIGITS)
    multiple=re.search(r'\nشهرهای تأییدشده سؤال: (.+)$',text)
    if multiple:
        cities=json.loads(multiple[1]);value=text.split('\nشهرهای تأییدشده سؤال:',1)[0]
        if not re.search(r'فروش|سفارش|خرید',value) or not re.search(r'موفق|پرداخت.?شده',value):raise ServiceError('scope_unsupported','شاخص یا وضعیت مقایسه چند شهر باید مشخص باشد.')
        metric='orders' if 'تعداد' in value and not re.search(r'مبلغ|درآمد',value) else 'revenue'
        if re.search(r'ماه|سال|هفته|تاریخ|آنلاین|محصول',value):raise ServiceError('scope_unsupported','فیلتر اضافی این مقایسه هنوز قابل تأیید نیست؛ شرطی حذف نمی‌شود.')
        return {'kind':'city_ranking','city':None,'cities':cities,'metric':metric,'direction':'min' if 'کمترین' in value else 'max','limit':1,'dates':[],'channel':None}
    if re.search(r'(کدام|کدوم|چه)(?:\s+\d+)?\s+شهر|\d+\s+شهر',normalized) and re.search(r'بیشترین|بالاترین|کمترین',text) and re.search(r'مبلغ|درآمد',text) and re.search(r'پرداخت.?شده|موفق',text):
        if city_constraint(text):raise ServiceError('scope_unsupported','مقایسه شهرها همراه فیلتر یک شهر مبهم است.')
        count=re.search(r'(\d+)\s+شهر|شهر\s+(\d+)',normalized)
        limit=int(next(v for v in count.groups() if v)) if count else 1
        if not 1<=limit<=200:raise ServiceError('report_limit','تعداد شهر باید بین ۱ و ۲۰۰ باشد.')
        return checked_report_plan(text,{'kind':'city_ranking','city':None,'limit':limit,'direction':'min' if 'کمترین' in text and 'بیشترین' not in text else 'max'})
    if re.search(r'مشتری',text) and re.search(r'سفارش|خرید',text) and re.search(r'پرداخت.?شده|موفق',text) and re.search(r'ندار|نداشت|بدون|هیچ',text):
        if city_constraint(text) or re.search(r'ماه|سال|هفته|تاریخ|بازه|\d{4}-',text):
            raise ServiceError('scope_unsupported','فهرست بدون سابقه خرید با فیلتر شهر یا زمان هنوز تأیید نشده است؛ شرط حذف نمی‌شود.')
        normalized=text.replace('\u200c',' ').replace('ي','ی').replace('ك','ک')
        allowed=set('کدام کدوم چه مشتری مشتریان ها های ی سفارش سفارشهای خرید خریدهای پرداخت شده موفق هیچ ندارد ندارند نداشته داشته اند هرگز تاکنون بدون هستند است را رو نشان نمایش بده بدهید معرفی کن کنید لطفا لطفاً سابقه سفارشات'.split())
        if re.search(r'\d',normalized) or set(re.findall(r'[^\W\d_]+',normalized))-allowed:
            raise ServiceError('scope_unsupported','این درخواست شرط اضافه‌ای دارد که گزارش بدون سابقه خرید آن را پوشش نمی‌دهد؛ شرط حذف نمی‌شود.')
        return {'kind':'customers_without_paid'}
    return None

async def build_set_report(plan,tools,semantics,emit):
    events=[]
    async def call(name,args):
        tools.validate(name,args);identity='set-'+str(len(events)+1)
        await emit({'stage':'tool','id':identity,'status':'running','tool':name,'arguments':args,'operation':'tools/call'})
        result,duration=await tools.execute(name,args)
        event={'stage':'tool','id':identity,'status':'success','tool':name,'arguments':args,'result':result,'duration_ms':duration}
        events.append(event);await emit(event)
        return result.get('data',result)
    if plan['kind']=='city_ranking':
        if plan.get('channel'):raise ServiceError('scope_unsupported','ابزار تجمیع شهر فیلتر کانال ندارد؛ شرط کانال حذف نمی‌شود.')
        args={'group':'city','limit':200}
        if plan.get('dates'):args.update(start=plan['dates'][0],end=plan['dates'][1])
        rows=await call('sales_breakdown',args)
        if not isinstance(rows,list) or len(rows)>=200:raise ServiceError('report_incomplete','کامل‌بودن فهرست شهرها تأیید نشد.')
        if not rows:return 'فروش پرداخت‌شده‌ای در دامنه درخواست وجود ندارد.',events
        values=[]
        for row in rows:
            if not isinstance(row.get('group'),str) or type(row.get('orders')) is not int:raise ServiceError('report_invalid','ساختار تجمیع شهر معتبر نیست.')
            value=Decimal(str(row['revenue']))
            if not value.is_finite():raise ServiceError('report_invalid','مبلغ معتبر نیست.')
            values.append((row['group'],row['orders'],value))
        if plan.get('cities'):
            by_city={r[0]:r for r in values};values=[by_city.get(city,(city,0,Decimal(0))) for city in plan['cities']]
        metric_index=1 if plan.get('metric')=='orders' else 2
        extreme=(min if plan['direction']=='min' else max)(r[metric_index] for r in values)
        winners=[r for r in values if r[metric_index]==extreme]
        if plan.get('limit',1)>1:
            ranked=sorted(values,key=lambda r:((-r[2] if plan['direction']=='max' else r[2]),r[0]))[:plan['limit']]
            lines=[f'{len(ranked)} شهر به ترتیب مبلغ فروش پرداخت‌شده:', '', '| شهر | مجموع مبلغ | سفارش پرداخت‌شده |','| --- | --- | --- |']
            lines += [f'| {r[0]} | {r[2]:,f} | {r[1]} |' for r in ranked]
            lines+=['','منبع: sales_breakdown با group=city؛ واحد پول: '+str(semantics.get('revenue',{}).get('currency') or 'تنظیم نشده')]
            return '\n'.join(lines),events
        label='کمترین' if plan['direction']=='min' else 'بیشترین'
        extreme=Decimal(extreme)
        formatted=f'{extreme:,.0f}' if extreme==extreme.to_integral() else f'{extreme:,f}'
        metric_label='تعداد سفارش پرداخت‌شده' if metric_index==1 else 'مجموع مبلغ فروش پرداخت‌شده'
        result=f'{label} {metric_label}: '+formatted+'.\n'
        result+='\n'.join(f'{r[0]} — تعداد سفارش پرداخت‌شده: {r[1]:,}' for r in winners)
        result+='\nواحد پول: '+str(semantics.get('revenue',{}).get('currency') or 'تنظیم نشده؛ واحدی فرض نشده است.')
        result+='\nمنبع: sales_breakdown با group=city؛ تمام شهرهای برگشتی مقایسه شدند.'
        return result,events
    entities=semantics['entities'];schema=semantics['default_schema'];sales=entities['sales'];customers=entities['customers']
    stats=await call('database_column_stats',{'schema':schema,'name':sales['table'],'column':sales['columns']['date']})
    if stats['total']!=stats['non_null']:raise ServiceError('report_unverified','بعضی سفارش‌ها تاریخ ندارند؛ نبود خرید موفق از روی آخرین تاریخ قابل اثبات نیست.')
    boundary=stats['min'] if stats['total'] else date.min.isoformat()
    date.fromisoformat(boundary)
    rows=await call('customer_inactive',{'before_date':boundary,'limit':200})
    if not isinstance(rows,list) or len(rows)>=200:raise ServiceError('report_incomplete','فهرست ممکن است بریده شده باشد؛ کامل بودن آن تأیید نشد.')
    if any(row.get('last_purchase') is not None for row in rows):raise ServiceError('report_unverified','ابزار مشتری با تاریخ خرید برگرداند؛ نتیجه به عنوان بدون سفارش موفق پذیرفته نشد.')
    if not rows:return 'هیچ مشتریِ بدون سفارش پرداخت‌شده پیدا نشد. همه مشتریان برگشتی از دامنه کامل این بررسی، سابقه سفارش موفق دارند.\nمنبع: حداقل تاریخ واقعی سفارش‌ها و customer_inactive؛ تاریخ دلخواه استفاده نشد.',events
    ids=[row['customer_id'] for row in rows];found={};ct=customers['table'];cc=customers['columns']
    for start in range(0,len(ids),100):
        spec={'schema':schema,'root':ct,'columns':[ct+'.'+cc['id'],ct+'.'+cc['name']],
              'filters':[{'column':ct+'.'+cc['id'],'op':'in','value':ids[start:start+100]}],'limit':200,'offset':0}
        data=await call('database_safe_join',{'spec':spec})
        if data.get('truncated'):raise ServiceError('report_incomplete','مشخصات مشتریان کامل نیست.')
        for row in data['rows']:found[row[ct+'.'+cc['id']]]=row[ct+'.'+cc['name']]
    if set(found)!=set(ids):raise ServiceError('report_unverified','شناسه‌های مشتری با مشخصات برگشتی تطبیق ندارند.')
    lines=[f'{len(ids)} مشتری بدون سفارش پرداخت‌شده:', '', '| شناسه | نام |', '| --- | --- |']
    lines += [f'| {cid} | {str(found[cid]).replace("|","／")} |' for cid in sorted(ids)]
    lines+=['','منبع: customer_inactive با مرز پایین برگرفته از حداقل تاریخ واقعی سفارش‌ها؛ همه last_purchaseها NULL هستند.']
    return '\n'.join(lines),events
