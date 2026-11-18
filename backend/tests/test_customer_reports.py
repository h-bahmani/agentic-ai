import unittest
from app.customer_reports import ranking_request, ranking_report, build_customer_ranking
from app.config import Settings
from app.llm import ServiceError

class RankingTools:
    def __init__(self,duplicate=False):self.specs=[];self.duplicate=duplicate
    def validate(self,name,args):pass
    async def execute(self,name,args):
        if name=='business_semantics':return {'data':{'entities':{
            'sales':{'table':'orders','columns':{'id':'pk','customer_id':'buyer','amount':'value','status':'state'}},
            'customers':{'table':'people','columns':{'id':'pk','name':'label','region_id':'place'}},
            'regions':{'table':'locations','columns':{'id':'pk','city':'label'}}},
            'default_schema':'test_schema','successful_sale':{'status':'confirmed'},'revenue':{'currency':None}}},1
        if name=='database_relationship_graph':return {'data':{'edges':[
            {'id':'e1','from_table':'orders','to_table':'people','from_columns':['buyer'],'to_columns':['pk']},
            {'id':'e2','from_table':'people','to_table':'locations','from_columns':['place'],'to_columns':['pk']}]}},1
        spec=args['spec'];self.specs.append(spec)
        sid=1 if spec['offset']==0 or self.duplicate else 2
        row={'orders.pk':sid,'orders.buyer':sid,'orders.value':70 if sid==1 else 90,'people.label':'نام یکسان','locations.label':'تهران'}
        return {'data':{'rows':[row],'truncated':spec['offset']==0,'next_offset':200 if spec['offset']==0 else None}},1

class ReportTests(unittest.IsolatedAsyncioTestCase):
    async def test_complete_pages_are_aggregated_by_ids_not_names(self):
        tools=RankingTools()
        async def emit(e):pass
        answer,events=await build_customer_ranking({'city':'تهران','limit':2},tools,Settings(),emit)
        self.assertIn('| 2 | نام یکسان | تهران | 1 | 90',answer)
        self.assertIn('| 1 | نام یکسان | تهران | 1 | 70',answer)
        self.assertEqual(len(events),4)
        self.assertEqual([s['offset'] for s in tools.specs],[0,200])
        self.assertIn({'column':'locations.label','op':'eq','value':'تهران'},tools.specs[0]['filters'])
        self.assertIn({'column':'orders.state','op':'eq','value':'confirmed'},tools.specs[0]['filters'])

    async def test_incomplete_pagination_cannot_produce_final_ranking(self):
        async def emit(e):pass
        with self.assertRaises(ServiceError) as cm:
            await build_customer_ranking({'city':'تهران','limit':2},RankingTools(),Settings(report_max_pages=1),emit)
        self.assertEqual(cm.exception.code,'report_incomplete')

    async def test_duplicate_sale_cannot_inflate_customer_total(self):
        async def emit(e):pass
        with self.assertRaises(ServiceError) as cm:
            await build_customer_ranking({'city':'تهران','limit':2},RankingTools(duplicate=True),Settings(),emit)
        self.assertEqual(cm.exception.code,'report_invalid')

    async def test_current_request_extracts_city_and_persian_count(self):
        self.assertEqual(ranking_request('۴ مشتری شهر تهران با بیشترین مبلغ خرید موفق'),{'city':'تهران','limit':4})
        self.assertIsNone(ranking_request('تو دیگه چی کارا می‌کنی'))

if __name__=='__main__':unittest.main()
