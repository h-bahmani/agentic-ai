import unittest
from app.grounding import city_constraint
from app.set_reports import set_report_request,build_set_report
from app.llm import ServiceError

class SetReportTests(unittest.IsolatedAsyncioTestCase):
    def test_question_dimension_is_not_a_literal_city(self):
        for text in ['کدام شهر بیشترین مبلغ فروش پرداخت‌شده را دارد؟','چه شهر بیشترین مبلغ فروش موفق دارد؟','کدوم شهر کمترین مبلغ فروش موفق را دارد؟']:
            self.assertIsNone(city_constraint(text))
            self.assertEqual(set_report_request(text)['kind'],'city_ranking')
        self.assertEqual(city_constraint('مبلغ فروش شهر شیراز چقدر است؟'),'شیراز')

    async def test_city_ranking_compares_all_rows_and_preserves_ties(self):
        class Tools:
            def validate(self,n,a):pass
            async def execute(self,n,a):
                self.args=a
                return {'data':[{'group':'شهر الف','orders':2,'revenue':40.5},{'group':'شهر ب','orders':1,'revenue':40.5},{'group':'شهر پ','orders':3,'revenue':20}]},1
        async def emit(e):pass
        tools=Tools();plan=set_report_request('کدام شهر بیشترین مبلغ فروش پرداخت‌شده را دارد؟')
        answer,_=await build_set_report(plan,tools,{'revenue':{'currency':None}},emit)
        self.assertIn('شهر الف',answer);self.assertIn('شهر ب',answer);self.assertNotIn('شهر پ',answer)
        self.assertIn('40.5',answer)
        self.assertEqual(tools.args['group'],'city')

    async def test_no_paid_uses_observed_minimum_not_an_arbitrary_date(self):
        class Tools:
            def validate(self,n,a):pass
            async def execute(self,n,a):
                if n=='database_column_stats':return {'data':{'total':100,'non_null':100,'min':'2018-05-10'}},1
                self.before=a['before_date'];return {'data':[]},1
        async def emit(e):pass
        semantics={'default_schema':'fixture','entities':{'sales':{'table':'orders','columns':{'date':'on_date'}},'customers':{'table':'people','columns':{'id':'pk','name':'label'}}}}
        tools=Tools();plan=set_report_request('کدام مشتریان سفارش پرداخت‌شده ندارند؟')
        answer,_=await build_set_report(plan,tools,semantics,emit)
        self.assertEqual(tools.before,'2018-05-10');self.assertIn('هیچ مشتری',answer)

    async def test_missing_dates_cannot_prove_absence_of_paid_orders(self):
        class Tools:
            def validate(self,n,a):pass
            async def execute(self,n,a):return {'data':{'total':100,'non_null':90,'min':'2018-05-10'}},1
        async def emit(e):pass
        semantics={'default_schema':'fixture','entities':{'sales':{'table':'orders','columns':{'date':'on_date'}},'customers':{}}}
        with self.assertRaises(ServiceError):await build_set_report({'kind':'customers_without_paid'},Tools(),semantics,emit)

    def test_extra_customer_filter_is_not_dropped(self):
        with self.assertRaises(ServiceError):set_report_request('کدام مشتریان ویژه سفارش پرداخت‌شده ندارند؟')

if __name__=='__main__':unittest.main()
