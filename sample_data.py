"""실제 기업과 무관한 가상 공시. 수집 후와 같은 분석 경로를 거친다."""
from account_mapper import ACCOUNTS


def demo(end=2025):
    snapshot={'schema_version':1,'synthetic':True,'company':{'code':'00000001','name':'가상제조','stock_code':'000001'},
              'profile':{'corp_cls':'Y','acc_mt':'12','induty_code':'261','stock_code':'000001'},
              'end_year':end,'fs_div':'CFS','years':{}}
    notes={'schema_version':1,'records':[]}
    for i,y in enumerate(range(end-3,end+1)):
        receipt=f'{y+1}0301000001'
        vals={'revenue':1000+i*100,'cost':600+i*60,'gross_profit':400+i*40,'operating_profit':150+i*20,
              'profit':100+i*10,'parent_profit':80+i*10,'parent_equity':500+i*50,'equity':600+i*60,
              'assets':1000+i*100,'liabilities':400+i*40,'interest':20+i*2,'receivables':100+i*10,'inventory':80+i*8}
        rows=[]
        for key,v in vals.items():
            ids,aliases,statements=ACCOUNTS[key]
            rows.append({'corp_code':'00000001','bsns_year':str(y),'reprt_code':'11011','rcept_no':receipt,
                         'sj_div':statements[0],'account_id':ids[0],'account_nm':aliases[0],
                         'thstrm_amount':str(v*1_000_000),'currency':'KRW'})
        periods=[{'bsns_year':str(y),'reprt_code':'11011','stock_code':'000001','rcept_no':receipt,
                  'fs_div':'CFS','sj_div':sj,'thstrm_dt':dt} for sj,dt in [('BS',f'{y}.12.31'),('IS',f'{y}.01.01 ~ {y}.12.31')]]
        snapshot['years'][str(y)]={'full':rows,'period':periods,'errors':[]}
        notes['records'].append({'corp_code':'00000001','year':y,'fs_div':'CFS','currency':'KRW','unit':'KRW','origin':'manual',
           'scope':'total_trade','reported_basis':'net','balance_date':f'{y}-12-31','rcept_no':receipt,
           'gross':str((105+i*11)*1_000_000),'allowance':str((5+i)*1_000_000),'net':str((100+i*10)*1_000_000),
           'ecl_expense':'2000000','writeoffs':'1000000','recoveries':None,'other_movements':'0',
           'source_url':'https://example.com/synthetic','source_ref':'가상 TEST 주석: 실제 기업 공시 아님'})
    return snapshot,notes
