# MobinAid — دستیار سازمانی محلی

رابط فارسی برای ابزارهای شرکت، Outlook کلاسیک و تبدیل صوت با Whisper موجود.
مدل فعلی از Ollama اجرا می‌شود. برنامه در زمان اجرا بسته یا مدل دانلود نمی‌کند.

## از کجا شروع کنم؟

- **اجرای برنامه و عیب‌یابی پورت‌ها:** [STARTUP.md](STARTUP.md)
- **کد اتصال Outlook:** [backend/app/outlook.py](backend/app/outlook.py)
- **معماری و محدودیت‌های Outlook:** [docs/outlook.md](docs/outlook.md)
- **معماری کل سامانه:** [docs/architecture.md](docs/architecture.md)
- **برنامهٔ توسعهٔ اتصال‌های شرکت:** [docs/enterprise-upgrade-plan.md](docs/enterprise-upgrade-plan.md)
- **گزارش‌های تاریخی:** [MOBINAID_HANDOFF.md](MOBINAID_HANDOFF.md) و [docs/verification.md](docs/verification.md)؛ مرجع وضعیت فعلی نیستند.

## اجرا

در PowerShell:

```powershell
Set-Location 'C:\Users\ext.h.bahmanighajar\Documents\ChatGPT\agent\enterprise_agent\backend'
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' -m uvicorn app.main:app --host 127.0.0.1 --port 8570 --no-access-log
```

رابط: http://127.0.0.1:8570/ . برای Outlook از «اتصال‌ها» متصل شوید.
سرویس‌های وابسته و روش اجرای پس‌زمینه در STARTUP.md هستند.

## ساختار

```text
enterprise_agent/
  backend/
    app/           کد فعال API، مدل، MCP، Outlook و فایل/صوت
    tests/         تست‌های واحد و اتصال درون‌پردازه‌ای
    requirements.txt
  frontend/        رابط؛ فقط style.css فایل استایل فعال است
  scripts/
    start.ps1      اجرای backend با Python موجود
    test.ps1       اجرای تست‌ها
    verify.py      بررسی سلامت، MCP Outlook، برنامه‌ریز یا یک درخواست زنده
  docs/            معماری، Outlook و برنامهٔ توسعه
  data/            داده‌های محلی؛ حذف نکنید
  logs/            لاگ اجرای پس‌زمینه
  STARTUP.md
  MOBINAID_HANDOFF.md
```

کد Outlook در یک ماژول با بخش‌های جدا برای COM، سرور MCP، کلاینت، برنامه‌ریز و اجراست.
تست‌های Outlook در test_outlook.py تجمیع شده‌اند. کد قدیمیِ
تشخیص سناریوها از مسیر اجرایی و ماژول حذف شده است.

## بررسی و تست

از پوشهٔ enterprise_agent:

```powershell
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' scripts/verify.py health
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' scripts/verify.py outlook
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' scripts/verify.py plan --text 'سؤال ایمیلی موردنظر'
```

health فقط وضعیت سرویس‌ها را می‌خواند. outlook ابزارهای MCP و دسترسی واقعی را بررسی
می‌کند، بدون خواندن بدنهٔ ایمیل. plan مدل واقعی را صدا می‌زند و فقط برنامهٔ درخواست را
نمایش می‌دهد؛ هیچ ایمیلی جست‌وجو نمی‌شود. chat --text درخواست واقعی را اجرا می‌کند و
نشست مستقل دارد؛ اتصال خصوصی مرورگر دیگر را خودکار نمی‌گیرد. verify.py --help راهنماست.

تست‌های واحد:

```powershell
Set-Location backend
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' -m unittest discover -s tests -q
```

تست‌های عادی از fixture/double استفاده می‌کنند. موفقیتشان به معنی صحت عمومی مدل نیست.
ابزارهای بررسی زنده جدا و اختیاری‌اند؛ تستی به مدل آموزش نمی‌دهد.

## وضعیت واقعی

- API شرکت و MCP فقط‌خواندنی، Ollama محلی، Whisper موجود و Outlook کلاسیک adapter واقعی دارند.
- Outlook برای همین کاربر ویندوز است؛ هویت چندکاربرهٔ مرکزی پیاده‌سازی نشده.
- برنامه‌ریز عمومی هنوز در برخی عبارت‌ها شکست دارد؛ برنامهٔ نامعتبر باید قبل از اجرا رد شود.
- مرور ایمیل گزیدهٔ مستند است، نه تضمین خلاصه‌سازی معنایی. استخراج کار، احتمالی است.
- GPU شرکت و Telegram متصل نیستند. TTS و تحلیل تصویر/PDF کامل نیستند.
- تاریخچهٔ چت حافظه‌ای است؛ یادآوری و فایل‌ها در SQLite محلی‌اند. restart چت را پاک می‌کند.
- هیچ ارسال/حذف/ویرایش ایمیلی فعال نیست. محتوای ایمیل دستور اجرایی محسوب نمی‌شود.

## تنظیمات


فایل `backend/.env.example` خودکار خوانده نمی‌شود. در صورت نیاز یک کپی مستقل بسازید و با `-EnvFile` بدهید. متغیرهای موجود process نیز پشتیبانی می‌شوند.

| متغیر | پیش‌فرض / کاربرد |
|---|---|
| LLM_PROVIDER | ollama؛ adapter فعال فعلی |
| OLLAMA_BASE_URL | http://127.0.0.1:11434؛ فقط loopback |
| OLLAMA_MODEL | qwen3:4b؛ باید نصب کامل باشد |
| MCP_URL | http://127.0.0.1:8560/mcp؛ فقط loopback |
| MCP_ALLOWLIST | all_readonly؛ هر ۴۶ ابزار بازبینی‌شدهٔ فعلی؛ ابزار ناشناخته یا نوشتنی خودکار فعال نمی‌شود |
| LLM_CONTEXT | 4096 توکن در Ollama؛ بودجه ورودی محافظه‌کارانه برحسب نویسه است، tokenizer دقیق نیست |
| LLM_TEMPERATURE | 0 |
| LLM_TIMEOUT | 120 ثانیه برای درخواست مدل |
| RUN_TIMEOUT | 240 ثانیه برای کل اجرا |
| MAX_TOOL_STEPS | 6 فراخوانی؛ بین 1 و 12 |
| AGENT_API_TOKEN | اختیاری؛ shared secret توسعه، در فیلد توکن رابط وارد می‌شود |

