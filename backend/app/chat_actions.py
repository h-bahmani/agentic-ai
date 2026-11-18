"""Deterministic chat actions over user-supplied facts and explicit reminder times."""
import re
from decimal import Decimal
from datetime import datetime,timezone,timedelta
from .grounding import DIGITS

def supplied_comparison(text):
    value=text.translate(DIGITS)
    if not re.search(r'کمترین|بیشترین|کوچکترین|بزرگترین',value):return None
    matches=re.findall(r'(?:^|[،,؛;\n]|\sو\s)\s*([\u0621-\u06ff][\u0621-\u06ff ]{0,30}?)\s+(\d(?:[\d٬,]*\d)?(?:\.\d+)?)',value)
    if len(matches)<2:return None
    rows=[]
    for label,number in matches:
        label=re.sub(r'^شهر\s+','',label.strip())
        rows.append((label,Decimal(number.replace('٬','').replace(',',''))))
    if len({r[0] for r in rows})!=len(rows):return None
    minimum=bool(re.search(r'کمترین|کوچکترین',value))
    result=(min if minimum else max)(r[1] for r in rows)
    winners=[label for label,number in rows if number==result]
    number=f'{result:,.0f}' if result==result.to_integral() else f'{result:,f}'
    return (' و '.join(winners)+(' کمترین' if minimum else ' بیشترین')+' مقدار را دارند: '+number+('.\nچند مورد برابرند و همه نمایش داده شدند.' if len(winners)>1 else '.')+'\nمبنای محاسبه: فقط اعدادی که خودت در همین پیام داده‌ای؛ MCP اجرا نشده است.')

def relative_reminder(text):
    if not re.search(r'یادم\s*(?:بنداز|بینداز)|یادآوری',text):return None
    normalized=text.translate(DIGITS)
    duration=re.search(r'(\d+)\s*(ثانیه|دقیقه|ساعت|روز)\s*(?:دیگه|دیگر|بعد)',normalized)
    if not duration:return None
    seconds=int(duration[1])*{'ثانیه':1,'دقیقه':60,'ساعت':3600,'روز':86400}[duration[2]]
    if not 1<=seconds<=366*86400:return None
    title=re.split(r'یادم\s*(?:بنداز|بینداز)|یادآوری\s*کن',text,maxsplit=1)
    if len(title)<2:return None
    title=re.sub(r'^\s*که\s*','',title[1]).strip(' .؟؛:')
    if not title or len(title)>300:return None
    return {'title':title,'due_at':datetime.now(timezone.utc)+timedelta(seconds=seconds)}
