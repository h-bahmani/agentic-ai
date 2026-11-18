import tempfile
import unittest
from pathlib import Path
from datetime import datetime,timedelta,timezone
from app.local_workspace import LocalWorkspace
from app.outlook import OutlookLocal
from app.llm import ServiceError
from app.outlook import recent_mail_range,plain_date_listing,OutlookTools,render_mail_listing,model_view,simple_latest_request

class OutlookSafetyTests(unittest.TestCase):
    def test_unread_numeric_date_request_routes_without_generation(self):
        from app.outlook import unread_filter
        for digits in ['۷','7','٧']:
            text=f'ایمیل‌های خوانده‌نشده من در {digits} روز اخیر را نمایش بده.'
            self.assertTrue(plain_date_listing(text))
            self.assertTrue(unread_filter(text))
            now=datetime(2026,11,18,tzinfo=timezone.utc)
            bounds=recent_mail_range(text,now)
            self.assertEqual(datetime.fromisoformat(bounds['start_utc']),now-timedelta(days=7))
        self.assertIsNone(unread_filter('ایمیل امروز'))
        self.assertFalse(unread_filter('ایمیل خوانده شده'))
    def test_today_and_yesterday_use_tehran_midnight(self):
        now=datetime(2026,11,18,1,tzinfo=timezone.utc)
        today=recent_mail_range('من ایمیلی امروز داشتم؟',now)
        self.assertEqual(today['start_utc'],'2026-11-17T20:30:00+00:00')
        self.assertEqual(today['end_utc'],'2026-11-18T20:30:00+00:00')
        yesterday=recent_mail_range('دیروز ایمیل داشتم',now)
        self.assertEqual(yesterday['end_utc'],today['start_utc'])

    def test_simple_date_requests_are_not_model_planned(self):
        self.assertTrue(plain_date_listing('من ایمیلی امروز داشتم؟'))
        self.assertTrue(plain_date_listing('خوب ایمیلی که در یه هفته پیش به من زده شده چی بوده'))
        self.assertFalse(plain_date_listing('ایمیل قرارداد امروز'))
        self.assertFalse(plain_date_listing('ایمیل علی امروز'))
        self.assertFalse(plain_date_listing('متن ایمیل دیروز را خلاصه کن'))

    def test_short_references_do_not_expose_native_identifiers(self):
        tools=OutlookTools(None,frozenset())
        native='AB'*1000
        ref=tools.reference(native)
        self.assertEqual(ref,'mail1')
        self.assertEqual(tools.reference(native),ref)
        self.assertEqual(tools.ids[ref],native)
        self.assertNotIn(native,ref)

    def test_latest_read_workflow_does_not_discard_sender_filters(self):
        self.assertTrue(simple_latest_request('آخرین ایمیل دریافتی را پیدا کن، متنش را بخوان و موضوعش را بگو.'))
        self.assertTrue(simple_latest_request('آخرین ایمیل را خلاصه کن'))
        self.assertFalse(simple_latest_request('آخرین ایمیل علی را خلاصه کن'))

    def test_preview_is_explicit_and_source_preserved(self):
        source={'body':'الف'*3000,'body_truncated':True,'next_offset':1600}
        view=model_view(source)
        self.assertIn('پیش‌نمایش',view['body'])
        self.assertEqual(len(source['body']),9000)
        self.assertEqual(view['next_offset'],1600)

    def test_empty_inbox_report_does_not_claim_no_email_everywhere(self):
        bounds=recent_mail_range('امروز')
        text=render_mail_listing({'items':[],'scan_complete':True},bounds)
        self.assertIn('صندوق ورودی',text)
        self.assertIn('آرشیوهای دیگر بررسی نشده',text)

    def test_relative_date_is_seven_days_with_aware_endpoints(self):
        now=datetime(2026,11,18,9,tzinfo=timezone.utc)
        for text in ['یه هفته پیش','یک هفته قبل','هفته گذشته']:
            result=recent_mail_range(text,now)
            self.assertEqual(datetime.fromisoformat(result['start_utc']),now-timedelta(days=7))
            self.assertEqual(datetime.fromisoformat(result['end_utc']),now)
        self.assertIsNone(recent_mail_range('قرارداد'))
    def test_no_write_operation(self):
        with self.assertRaises(ServiceError):OutlookLocal().call('send')

    def test_calendar_rejects_naive_or_unbounded_range(self):
        for start,end in [('2026-11-18','2026-11-19'),('2026-11-18T00:00:00Z','2027-01-01T00:00:00Z')]:
            with self.assertRaises(ServiceError):OutlookLocal().call('calendar',start_utc=start,end_utc=end)

    def test_sync_deduplicates_reschedules_cancels_and_isolates(self):
        with tempfile.TemporaryDirectory() as directory:
            store=LocalWorkspace(Path(directory)/'test.sqlite3')
            now=datetime.now(timezone.utc)
            m={'id':'meeting1','title':'test','start':now+timedelta(hours=1),'due':now+timedelta(minutes=45)}
            store.sync_outlook_reminders('user1',[m],now)
            first=store.reminders('user1')[0]['id']
            store.sync_outlook_reminders('user1',[m],now)
            self.assertEqual(store.reminders('user1')[0]['id'],first)
            self.assertEqual(store.reminders('user2'),[])
            m={**m,'start':now+timedelta(hours=2),'due':now+timedelta(minutes=105)}
            store.sync_outlook_reminders('user1',[m],now)
            self.assertEqual(len(store.reminders('user1')),1)
            self.assertNotEqual(store.reminders('user1')[0]['id'],first)
            store.sync_outlook_reminders('user1',[],now)
            self.assertEqual(store.reminders('user1'),[])

    def test_sync_keeps_acknowledged_occurrence_seen(self):
        with tempfile.TemporaryDirectory() as directory:
            store=LocalWorkspace(Path(directory)/'test.sqlite3');now=datetime.now(timezone.utc)
            m={'id':'1','title':'meeting','start':now+timedelta(minutes=5),'due':now-timedelta(minutes=10)}
            store.sync_outlook_reminders('owner',[m],now);store.tick(now+timedelta(seconds=1))
            store.acknowledge('owner',store.reminders('owner')[0]['id'])
            store.sync_outlook_reminders('owner',[m],now)
            self.assertEqual(store.reminders('owner')[0]['status'],'seen')
