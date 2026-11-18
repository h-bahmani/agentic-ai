"""Deterministic report arithmetic over verified tool evidence, never model guesses."""
import json
from decimal import Decimal

def channel_percentage(text: str, events: list[dict]) -> str | None:
    normalized = text.lower()
    if not ('درصد' in normalized or '%' in normalized) or not ('آنلاین' in normalized or 'online' in normalized):
        return None
    for event in reversed(events):
        if event.get('tool') != 'sales_breakdown' or event.get('arguments', {}).get('group') != 'channel':
            continue
        data = event['result']
        if 'text' in data:
            try: data = json.loads(data['text'])
            except (ValueError, TypeError): return None
        rows = data.get('data') if isinstance(data, dict) else None
        limit = event['arguments'].get('limit', 30)
        if not isinstance(rows, list) or not rows or len(rows) >= limit:
            return 'ابزار اجرا شد، اما کامل‌بودن فهرست کانال‌ها تأیید نشده است؛ درصد کل را نمی‌توان محاسبه کرد.'
        if any(not isinstance(r, dict) or type(r.get('orders')) is not int or r['orders'] < 0 for r in rows):
            return 'ابزار اجرا شد، اما تعداد سفارش‌ها برای محاسبه معتبر نیست؛ درصدی محاسبه نشد.'
        groups = [r.get('group') for r in rows]
        if len(set(groups)) != len(groups): return 'کانال تکراری در پاسخ ابزار وجود دارد؛ درصد تأیید نشد.'
        online = next((r['orders'] for r in rows if r.get('group') == 'online'), None)
        if online is None: return 'کانال online در دادهٔ برگشتی وجود ندارد؛ درصدی حدس زده نشد.'
        total = sum(r['orders'] for r in rows)
        if not total: return 'تعداد کل فروش‌های موفق صفر است؛ درصد قابل محاسبه نیست.'
        percentage = Decimal(online) * 100 / Decimal(total)
        return (f'سهم تعداد فروش‌های موفق آنلاین از کل فروش‌های موفق: {percentage:.3f}٪\n'
                f'{online:,} سفارش آنلاین از مجموع {total:,} سفارش موفق.\n'
                f'محاسبه: {online} ÷ {total} × ۱۰۰.\n'
                f'منبع: sales_breakdown، گروه channel، limit={limit}.\n'
                'این شاخص سهم آنلاین از فروش‌های موفق است؛ نرخ موفقیت در میان تمام سفارش‌های آنلاین از این داده مشخص نمی‌شود.')
    return None
