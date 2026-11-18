import unittest
from app.grounding import validate_tool_scope,validate_final,scoped_sales_request,verified_request_scope
from app.llm import ServiceError

class GuardTests(unittest.IsolatedAsyncioTestCase):
    def test_global_revenue_cannot_answer_city_request(self):
        with self.assertRaises(ServiceError):validate_tool_scope('مجموع فروش‌های موفق شهر شیراز؟','sales_revenue',{})

    def test_missing_date_and_channel_are_not_silently_ignored(self):
        for text,name,args in [('درآمد از 2026-01-01 تا 2026-12-31','sales_revenue',{}),
                               ('درآمد فروش آنلاین','sales_revenue',{}),
                               ('تعداد سفارش‌های موفق','sales_order_count',{'kind':'all'})]:
            with self.subTest(text=text),self.assertRaises(ServiceError):validate_tool_scope(text,name,args)

    def test_currency_and_numbers_require_evidence(self):
        events=[{'tool':'sales_revenue','result':{'data':{'revenue':12345}}}]
        validate_final('درآمد؟','مجموع مبلغ 12,345 است.',events)
        for text in ['مبلغ 21,131,010,000 است.','مبلغ 12,345 تومان است.','مبلغ 12,345 ریال است.']:
            with self.subTest(text=text),self.assertRaises(ServiceError):validate_final('درآمد؟',text,events)

    def test_no_results_cannot_produce_a_numeric_answer(self):
        with self.assertRaises(ServiceError):validate_final('درآمد؟','درآمد 100 است.',[])

    def test_different_city_metric_date_channel_not_a_memorized_answer(self):
        plan=scoped_sales_request('میانگین مبلغ فروش‌های موفق آنلاین در شهر قم از 2026-01-01 تا 2026-12-31 چقدر است؟')
        self.assertEqual(plan,{'kind':'sales_summary','city':'قم','metric':'average','dates':['2026-01-01','2026-12-31'],'channel':'online'})
        with self.assertRaises(ServiceError):scoped_sales_request('مجموع مبلغ فروش موفق شهر قم در ماه گذشته چقدر است؟')
        for text in ['مجموع مبلغ فروش موفق اقساطی شهر قم چقدر است؟',
                     'مجموع مبلغ فروش موفق شهر قم برای محصول خاص چقدر است؟']:
            with self.subTest(text=text),self.assertRaises(ServiceError):scoped_sales_request(text)

    async def test_city_catalog_discovers_names_without_city_keyword(self):
        class Tools:
            def validate(self,name,args):pass
            async def execute(self,name,args):
                if name=='business_semantics':return {'data':{'entities':{'regions':{'table':'places','columns':{'city':'label'}}},'default_schema':'test'}},1
                return {'data':[{'value':'بندر عباس'},{'value':'شهر تازه'}]},1
        catalog=[{'function':{'name':n}} for n in ['business_semantics','database_distinct']]
        async def emit(e):pass
        text,events=await verified_request_scope('مجموع مبلغ فروش‌های موفق در بندر عباس چقدر است؟',Tools(),catalog,emit)
        self.assertEqual(scoped_sales_request(text)['city'],'بندر عباس')
        self.assertEqual(len(events),2)

if __name__=='__main__':unittest.main()
