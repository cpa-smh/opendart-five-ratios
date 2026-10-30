"""기업별 4개년 수집, 재실행 가능한 스냅샷, 계정·주석 연결."""
from datetime import datetime, timezone
from pathlib import Path
from dataclasses import asdict
from decimal import Decimal
import json
import re
from dart_client import DartError, NoDataError
from account_mapper import map_accounts, accounting_checks, amount
from validation import validate_period, profile_reasons
from credit_risk import load_credit_notes, assess_receivables, credit_diagnostics
from ratios import calculate, fraction
from documents import fetch_notes


def save_json(data, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')


def select_company(companies, name, identifier=None, ask=input):
    matches = [c for c in companies if c.name == name and c.stock_code]
    if identifier:
        matches = [c for c in matches if identifier in (c.code, c.stock_code)]
    if len(matches)==1:
        return matches[0]
    if not matches:
        raise ValueError('일치하는 상장기업이 없습니다. 정확한 기업명·종목코드를 확인하세요.')
    print('\n'.join(f'{c.name}: 종목 {c.stock_code}, DART {c.code}' for c in matches))
    choice = ask('위 후보의 종목코드 또는 DART 고유번호: ').strip()
    selected = [c for c in matches if choice in (c.code,c.stock_code)]
    if len(selected)!=1:
        raise ValueError('후보 중 하나를 정확히 지정하세요.')
    return selected[0]


def collect(client, company, end_year, out, allow_ofs=False, ask=input, download_notes=True):
    profile = client.profile(company.code)
    issues = profile_reasons(profile)
    if profile.get('stock_code') != company.stock_code:
        issues.append('기업개황 종목코드 불일치')
    if issues:
        raise ValueError('; '.join(issues))
    snapshot = {'schema_version':1, 'synthetic':False, 'company':asdict(company), 'profile':profile,
                'end_year':end_year, 'fs_div':'CFS', 'retrieved_at':datetime.now(timezone.utc).isoformat(), 'years':{}}
    out = Path(out)
    def get_years(basis):
        years = {}
        for year in range(end_year-3,end_year+1):
            print(f'{company.name} {year} {basis} 조회')
            item = {'full':[], 'period':[], 'errors':[], 'no_data':False}
            years[str(year)] = item
            try:
                item['full'] = client.financials(company.code,year,basis)
                item['period'] = client.period_accounts(company.code,year,company.stock_code)
            except NoDataError:
                item['no_data'] = not item['full']
                item['errors'].append('재무제표 또는 기간 확인 자료 없음')
            except DartError as exc:
                item['errors'].append(str(exc))
            snapshot['years'] = years
            snapshot['fs_div'] = basis
            save_json(snapshot,out/'raw_snapshot.json')
        return years
    years = get_years('CFS')
    if any(x['no_data'] for x in years.values()):
        if allow_ofs or ask('연결 자료가 없는 연도가 있습니다. 전체 4개년을 별도로 조회할까요? [y/N]: ').strip().lower()=='y':
            years = get_years('OFS')
    snapshot['years'] = years
    if download_notes:
        for year,item in years.items():
            receipts = {r.get('rcept_no') for r in item['full']}
            if len(receipts)==1:
                try:
                    item['document_status'] = fetch_notes(client,receipts.pop(),snapshot['fs_div'],out/'documents')
                except DartError as exc:
                    item['document_status'] = str(exc)
    save_json(snapshot,out/'raw_snapshot.json')
    return snapshot


def load_snapshot(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if data.get('schema_version')!=1 or type(data.get('synthetic')) is not bool or data.get('fs_div') not in ('CFS','OFS'):
        raise ValueError('지원하지 않는 스냅샷 형식')
    end = data.get('end_year')
    if type(end) is not int or end<2018 or end>datetime.now().year:
        raise ValueError('스냅샷 연도 오류')
    company = data.get('company',{})
    if not re.fullmatch(r'\d{8}',str(company.get('code',''))) or not re.fullmatch(r'\d{6}',str(company.get('stock_code',''))):
        raise ValueError('스냅샷 기업 식별자 오류')
    if profile_reasons(data.get('profile',{})):
        raise ValueError('스냅샷 기업 기준 오류')
    if set(data.get('years',{})) != {str(y) for y in range(end-3,end+1)}:
        raise ValueError('스냅샷 4개년 자료 필요')
    for year,item in data['years'].items():
        for group in ('full','period'):
            if not isinstance(item.get(group),list):
                raise ValueError('스냅샷 행 형식 오류')
            for row in item[group]:
                if row.get('bsns_year')!=year or row.get('reprt_code')!='11011':
                    raise ValueError('스냅샷 연도/보고서 불일치')
                if group=='full' and row.get('corp_code')!=company['code']:
                    raise ValueError('스냅샷 기업 불일치')
                if group=='period' and row.get('stock_code')!=company['stock_code']:
                    raise ValueError('스냅샷 종목 불일치')
    return data


def load_supplement(path):
    notes = load_credit_notes(path)
    for note in notes.values():
        value = note.get('interest_expense')
        note['interest_expense'] = amount(value)
        if value is not None and (note['interest_expense'] is None or note['interest_expense']<0 or note.get('interest_scope')!='profit_loss'):
            raise ValueError('이자비용은 0 이상 숫자와 interest_scope=profit_loss가 필요합니다.')
    return notes


def analyze(snapshot, notes):
    observations = []
    basis = snapshot['fs_div']; code = snapshot['company']['code']
    for year in range(snapshot['end_year']-3,snapshot['end_year']+1):
        raw = snapshot['years'][str(year)]
        period = validate_period(raw['full'],raw['period'],year,basis,int(snapshot['profile']['acc_mt']))
        mapped = map_accounts(raw['full'])
        note = notes.get((code,year))
        credit = assess_receivables(mapped['receivables'],note,period['receipts'],fs_div=basis)
        issues = list(raw.get('errors',[])) + period['reasons']
        if raw.get('document_status'):
            issues.append(raw['document_status'])
        if note and note['balance_date'] != period['balance_date']:
            credit.update(net=None,gross=None,allowance=None,ecl_expense=None,writeoffs=None,other_movements=None,reconciliation=None,status='REVIEW',reason='주석 기준일과 실제 결산일 불일치')
        # 주석 전체 순액과 API 유동 순액은 다를 수 있다. 주석의 유동 순액이 있으면 별도로 대사한다.
        if note and note.get('current_net') is not None and note['fs_div']==basis and note['rcept_no'] in period['receipts']:
            current_net = amount(note['current_net'])
            if current_net is None or (mapped['receivables'].amount is not None and current_net != mapped['receivables'].amount):
                credit.update(net=None,gross=None,allowance=None,ecl_expense=None,writeoffs=None,other_movements=None,reconciliation=None,status='REVIEW',reason='주석 유동 순액과 API 매출채권 불일치')
        values = {k:v.amount for k,v in mapped.items()}
        interest_source = 'API'
        if note and note['fs_div']==basis and note['balance_date']==period['balance_date'] and note['rcept_no'] in period['receipts'] and note.get('interest_expense') is not None:
            if values['interest'] is not None and values['interest'] != note['interest_expense']:
                values['interest'] = None
                issues.append('API와 보완 이자비용 충돌: 이자보상배율 제외')
                interest_source = '충돌'
            else:
                values['interest'] = note['interest_expense']; interest_source = note['origin']
        o = {'year':year,'fs_div':basis,'eligible':period['valid'] and not profile_reasons(snapshot['profile']),
             'values':values,'mapping':mapped,'credit':credit,'note':note,'interest_source':interest_source,
             'period':period,'issues':issues,'checks':accounting_checks(mapped),'raw':raw}
        if observations:
            previous = observations[-1]
            o['ratios'] = calculate(o,previous)
            aux = {'allowance_ratio':fraction(credit['allowance'],credit['gross'])[0] if o['eligible'] else None,
                   'ecl_to_sales':fraction(credit['ecl_expense'],values['revenue'])[0] if o['eligible'] else None}
            o['aux'] = aux
            o['findings'],o['rollforward'] = credit_diagnostics(o,previous,
                {**aux,'ar_turnover':o['ratios']['ar_turnover']['value']}, {})
            if not (o['eligible'] and previous['eligible']):
                o['rollforward'] = None
        observations.append(o)
    return observations
