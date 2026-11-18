"""Verified ranking over the existing globally sorted paid-customer prefix."""
import re
from decimal import Decimal, InvalidOperation
from .llm import ServiceError
from .grounding import city_constraint

def ranking_request(text):
    if not re.search(r'مشتری',text) or not re.search(r'بیشترین|برترین|بالاترین|رتبه',text):return None
    if not re.search(r'خرید|مبلغ|فروش|پرداخت',text):return None
    city=city_constraint(text)
    if not city:return None
    digits=text.translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩','01234567890123456789'))
    count=re.search(r'(\d+)\s*(?:تا\s*)?مشتری',digits)
    words={'یک':1,'دو':2,'سه':3,'چهار':4,'پنج':5,'شش':6,'هفت':7,'هشت':8,'نه':9,'ده':10}
    number=int(count[1]) if count else next((n for w,n in words.items() if re.search(r'\b'+w+r'\s+مشتری',text)),None)
    if number is None:return None
    if not 1<=number<=50:raise ServiceError('report_limit','تعداد مشتری باید بین ۱ تا ۵۰ باشد.')
    return {'city':city,'limit':number}

def normalized(value):return str(value).strip().replace('ي','ی').replace('ك','ک')

def ranking_report(plan, result, fetched_limit=50):
    data=result.get('data',result)
    rows=data.get('customers') if isinstance(data,dict) else None
    if not isinstance(rows,list):raise ServiceError('report_invalid','ابزار فهرست معتبر مشتریان برنگرداند؛ گزارشی ساخته نشد.')
    parsed=[]
    ids=set()
    for row in rows:
        try:
            cid=row['customer_id'];orders=row['paid_orders'];name=row['name'];city=row['city']
            amount=Decimal(str(row['lifetime_value']))
            if type(cid) is not int or type(orders) is not int or orders<0 or not isinstance(name,str) or not isinstance(city,str) or not amount.is_finite() or amount<0 or cid in ids:raise ValueError()
            ids.add(cid);parsed.append((cid,name,city,orders,amount))
        except (KeyError,ValueError,InvalidOperation,TypeError) as exc:
            raise ServiceError('report_invalid','ستون‌های مشتری یا مبلغ معتبر نیست؛ هیچ مقداری حدس زده نشد.') from exc
    # Reviewed endpoint orders globally by paid value DESC, customer ID ASC.
    if parsed!=sorted(parsed,key=lambda r:(-r[4],r[0])):
        raise ServiceError('report_order_unverified','ترتیب خروجی ابزار با رتبه‌بندی مورد انتظار سازگار نیست؛ رتبه‌ها تأیید نشدند.')
    matching=[r for r in parsed if normalized(r[2])==normalized(plan['city'])]
    if fetched_limit is not None and len(matching)<plan['limit'] and len(rows)>=fetched_limit:
        raise ServiceError('city_ranking_incomplete','در ۵۰ مشتری اول کل شرکت، تعداد کافی از شهر درخواستی وجود نداشت. ابزار فعلی فیلتر شهر ندارد؛ رتبه‌بندی کامل این شهر تأیید نشد.')
    chosen=matching[:plan['limit']]
    if not chosen:return 'در فهرست کامل برگشتی، مشتری با خرید موفق برای شهر '+plan['city']+' پیدا نشد.'
    def cell(value):return str(value).replace('|','／').replace('\n',' ')
    lines=[f"{len(chosen)} مشتری شهر {plan['city']} با بیشترین مجموع مبلغ خرید موفق، به ترتیب نزولی مبلغ:",
           '', '| شناسه | نام کامل | شهر | سفارش موفق | مجموع مبلغ خرید |', '| --- | --- | --- | --- | --- |']
    def money(value):return f'{value:,.0f}' if value==value.to_integral() else f'{value:,f}'
    lines.extend('| '+' | '.join(cell(v) for v in (r[0],r[1],r[2],r[3],money(r[4])))+' |' for r in chosen)
    semantics=data.get('semantics',{})
    currency=semantics.get('revenue',{}).get('currency')
    source=('منبع: database_safe_join؛ شهر و وضعیت موفق در خود ابزار فیلتر شدند، تمام صفحات خوانده شد و مجموع مبلغ بر مبنای شناسهٔ یکتای سفارش محاسبه شد.' if data.get('scope')=='city_complete' else 'منبع: customer_overview؛ صدر فهرست جهانی پس از تأیید ترتیب، بر اساس شهر فیلتر شد.')
    lines+=['', 'واحد پول: '+(str(currency) if currency else 'در تنظیمات شرکت مشخص نشده است؛ ریال یا تومان فرض نشده است.'), source,
            'خرید موفق در این گزارش یعنی سفارش با وضعیت '+str(semantics.get('successful_sale',{}).get('status','نامشخص'))+'، نه لزوماً وجه وصول‌شده.']
    if len({r[1] for r in chosen})<len(chosen):lines+=['نام‌ها در دادهٔ شرکت تکراری‌اند؛ مشتریان با شناسه‌های متفاوت، جداگانه حساب شده‌اند.']
    return '\n'.join(lines)

async def build_customer_ranking(plan,tools,settings,emit):
    evidence=[]
    async def call(name,arguments):
        tools.validate(name,arguments)
        identity='tool-'+str(len(evidence)+1)
        await emit({'stage':'tool','id':identity,'status':'running','tool':name,'arguments':arguments,'operation':'tools/call'})
        value,duration=await tools.execute(name,arguments)
        event={'stage':'tool','id':identity,'status':'success','tool':name,'arguments':arguments,'result':value,'duration_ms':duration}
        evidence.append(event);await emit(event)
        return value.get('data',value)
    semantic=await call('business_semantics',{})
    if semantic.get('validation_errors'):raise ServiceError('report_mapping_invalid','نگاشت کسب‌وکار شرکت خطا دارد؛ گزارش قابل تأیید نیست.')
    try:
        entities=semantic['entities'];schema=semantic['default_schema']
        sales=entities['sales'];customers=entities['customers'];regions=entities['regions']
        st,ct,rt=sales['table'],customers['table'],regions['table']
        sc,cc,rc=sales['columns'],customers['columns'],regions['columns']
        paid=semantic['successful_sale']['status']
        if not paid:raise ValueError()
    except (KeyError,TypeError,ValueError) as exc:raise ServiceError('report_mapping_invalid','تعریف جدول‌ها و خرید موفق کامل نیست؛ گزارشی حدس زده نشد.') from exc
    graph=await call('database_relationship_graph',{'schema':schema})
    edges=[]
    for source,target,source_column,target_column in [(st,ct,sc['customer_id'],cc['id']),(ct,rt,cc['region_id'],rc['id'])]:
        matches=[e['id'] for e in graph['edges'] if e['from_table']==source and e['to_table']==target and e['from_columns']==[source_column] and e['to_columns']==[target_column]]
        if len(matches)!=1:raise ServiceError('report_relationship_invalid','رابطه یکتای اعلام‌شده میان فروش، مشتری و شهر پیدا نشد؛ join حدس زده نشد.')
        edges.append(matches[0])
    columns=[st+'.'+sc['id'],st+'.'+sc['customer_id'],st+'.'+sc['amount'],ct+'.'+cc['name'],rt+'.'+rc['city']]
    if plan.get('channel'):columns.append(st+'.'+sc['channel'])
    accumulated={};seen=set();offset=0
    for page in range(settings.report_max_pages):
        spec={'schema':schema,'root':st,'edges':edges,'columns':columns,
              'filters':[{'column':rt+'.'+rc['city'],'op':'eq','value':plan['city']},{'column':st+'.'+sc['status'],'op':'eq','value':paid}],
              'order_by':[{'column':st+'.'+sc['id'],'direction':'asc'}],'limit':200,'offset':offset}
        if plan.get('dates'):
            spec['filters'] += [{'column':st+'.'+sc['date'],'op':'ge','value':plan['dates'][0]},
                                {'column':st+'.'+sc['date'],'op':'le','value':plan['dates'][1]}]
        if plan.get('channel'):spec['filters'].append({'column':st+'.'+sc['channel'],'op':'eq','value':plan['channel']})
        data=await call('database_safe_join',{'spec':spec})
        rows=data.get('rows')
        if not isinstance(rows,list) or type(data.get('truncated')) is not bool:raise ServiceError('report_invalid','صفحه‌بندی نتیجه ابزار معتبر نیست.')
        for row in rows:
            try:
                sid,cid,amount,name,city=[row[k] for k in columns[:5]]
                if plan.get('channel') and row[columns[5]]!=plan['channel']:raise ValueError()
                amount=Decimal(str(amount))
                if type(sid) is not int or sid in seen or type(cid) is not int or not amount.is_finite() or amount<0 or normalized(city)!=normalized(plan['city']):raise ValueError()
                seen.add(sid)
                previous=accumulated.setdefault(cid,{'customer_id':cid,'name':name,'city':city,'paid_orders':0,'lifetime_value':Decimal(0)})
                if previous['name']!=name or previous['city']!=city:raise ValueError()
                previous['paid_orders']+=1;previous['lifetime_value']+=amount
            except (KeyError,TypeError,ValueError,InvalidOperation) as exc:raise ServiceError('report_invalid','دادهٔ تکراری یا ناسازگار در سفارش‌ها دیده شد؛ جمع مبلغ قابل تأیید نیست.') from exc
        if not data['truncated']:break
        next_offset=data.get('next_offset')
        if type(next_offset) is not int or next_offset<=offset:raise ServiceError('report_incomplete','ابزار ادامهٔ معتبر صفحه‌بندی را برنگرداند؛ رتبه‌بندی نهایی نشد.')
        offset=next_offset
    else:raise ServiceError('report_incomplete','سقف صفحات گزارش رسید و هنوز داده باقی است؛ رتبه‌بندی ناقص به‌عنوان نتیجه نهایی نمایش داده نمی‌شود.')
    ranked=sorted(accumulated.values(),key=lambda r:(-r['lifetime_value'],r['customer_id']))
    if plan.get('kind')=='sales_summary':
        total=sum((r['lifetime_value'] for r in ranked),Decimal(0));count=len(seen)
        metric=plan['metric']
        if metric=='average' and not count:raise ServiceError('report_empty','سفارش موفقی در دامنه درخواست پیدا نشد؛ میانگین قابل محاسبه نیست.')
        amount=total/count if metric=='average' else total
        number=f'{amount:,.0f}' if amount==amount.to_integral() else f'{amount:,.3f}'
        label={'sum':'مجموع مبلغ','average':'میانگین مبلغ','count':'تعداد'}[metric]
        result=f"{label} فروش‌های موفق در شهر {plan['city']}: "+(f'{count:,}' if metric=='count' else number)+'.'
        result+=f'\nتعداد سفارش‌های یکتای پرداخت‌شده در همین دامنه: {count:,}.'
        if metric!='count':result+='\nواحد پول: '+str(semantic.get('revenue',{}).get('currency') or 'در تنظیمات شرکت مشخص نشده است؛ هیچ واحدی فرض نشده است.')
        if plan.get('dates'):result+='\nبازه میلادی: '+plan['dates'][0]+' تا '+plan['dates'][1]
        if plan.get('channel'):result+='\nکانال: '+plan['channel']
        result+='\nشهر و وضعیت موفق در ابزار فیلتر شدند؛ همه صفحات خوانده و مبلغ با Decimal جمع زده شد. منبع: database_safe_join و تعاریف شرکت.'
        return result,evidence
    result=ranking_report(plan,{'data':{'customers':ranked,'semantics':semantic,'scope':'city_complete'}},fetched_limit=None)
    result+='\nداده‌ها در چند خواندن فقط‌خواندنی دریافت شده‌اند؛ تغییر هم‌زمان داده‌ها ممکن است روی سازگاری گزارش اثر بگذارد.'
    return result,evidence
