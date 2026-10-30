"""Decimal 계산. 비율은 0.15 형태로 저장하며 표시할 때만 %로 바꾼다."""
from decimal import Decimal

METRICS = {
    'operating_margin': ('영업이익률', '%', '영업이익 / 매출액'),
    'roe': ('ROE', '%', '당기순이익 / 평균 자본'),
    'ar_turnover': ('매출채권회전율', '배', '매출액 / 평균 매출채권 순액'),
    'debt_ratio': ('부채비율', '%', '부채총계 / 자본총계'),
    'interest_coverage': ('이자보상배율', '배', '영업이익 / 손익 이자비용'),
}


def average(a, b):
    return (a+b)/2 if a is not None and b is not None else None


def fraction(n, d):
    if n is None or d is None:
        return None, '필수 금액 누락'
    if d <= 0:
        return None, '분모가 0 또는 음수'
    return n/d, ''


def calculate(current, previous):
    v, p = current['values'], previous.get('values', {})
    c, pc = current['credit'], previous.get('credit', {})
    # 기초와 기말 지배지분이 모두 있어야 지배주주 ROE를 계산한다.
    parent = current['fs_div'] == 'CFS' and all(x is not None for x in
        (v.get('parent_profit'), v.get('parent_equity'), p.get('parent_equity')))
    basis = '지배기업 소유주' if parent else ('연결 전체(지배지분 자료 부족)' if current['fs_div']=='CFS' else '별도 전체')
    profit = v.get('parent_profit' if parent else 'profit')
    equity_key = 'parent_equity' if parent else 'equity'
    pairs = {
        'operating_margin': (v.get('operating_profit'), v.get('revenue')),
        'roe': (profit, average(p.get(equity_key), v.get(equity_key))),
        'ar_turnover': (v.get('revenue'), average(pc.get('net'), c.get('net'))),
        'debt_ratio': (v.get('liabilities'), v.get('equity')),
        'interest_coverage': (v.get('operating_profit'), v.get('interest')),
    }
    results = {}
    for key, (n, d) in pairs.items():
        value, reason = fraction(n, d)
        if not current['eligible']:
            value, reason = None, '당기 기간/통화/기업 검증 실패'
        elif key in ('roe','ar_turnover') and (not previous.get('eligible') or previous.get('fs_div') != current['fs_div']):
            value, reason = None, '기초 기간 또는 재무제표 범위 미검증'
        if key == 'ar_turnover' and value is None and not reason.startswith(('당기','기초')):
            reason += '; 매출채권 전체 순액과 기초 잔액 주석 확인 필요'
        if key == 'interest_coverage' and v.get('interest') is None:
            reason += '; 손익 이자비용 누락(금융비용/이자지급액으로 대체하지 않음)'
        results[key] = {'value':value, 'reason':reason, 'numerator':n, 'denominator':d,
                        'basis':basis if key=='roe' else '전체 매출액 대용 지표' if key=='ar_turnover' else current['fs_div']}
    return results
