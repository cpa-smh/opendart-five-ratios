"""명령행 또는 대화형 실행. 비밀정보가 포함될 수 있는 예외 원문을 출력하지 않는다."""
import argparse
from datetime import datetime
import os
from pathlib import Path
import sys
from dotenv import load_dotenv
from dart_client import DartClient,DartError
from pipeline import collect,analyze,select_company,save_json,load_snapshot,load_supplement

ROOT=Path(__file__).resolve().parent


def main(argv=None):
    p=argparse.ArgumentParser(description='OpenDART 3개년 재무비율 5개 및 매출채권·ECL 분석')
    source=p.add_mutually_exclusive_group(); source.add_argument('--demo',action='store_true'); source.add_argument('--snapshot',type=Path)
    p.add_argument('--company'); p.add_argument('--identifier',help='6자리 종목코드 또는 8자리 DART 고유번호')
    p.add_argument('--end-year',type=int); p.add_argument('--supplement',type=Path)
    p.add_argument('--allow-ofs',action='store_true',help='CFS 누락 시 전체 4개년 OFS 전환을 사전 승인')
    p.add_argument('--skip-documents',action='store_true',help='공시 원문 다운로드 생략')
    p.add_argument('--output-dir',type=Path); p.add_argument('--font',type=Path)
    a=p.parse_args(argv)
    try:
        if a.demo and (a.company or a.identifier or a.supplement):
            raise ValueError('가상 데모와 실제 기업/주석을 혼합할 수 없습니다.')
        if a.snapshot and (a.company or a.identifier or a.end_year):
            raise ValueError('스냅샷 실행의 기업과 연도는 저장된 자료를 사용합니다.')
        out=a.output_dir or ROOT/'reports'/datetime.now().strftime('run_%Y%m%d_%H%M%S_%f')
        if out.exists() and any(out.iterdir()): raise ValueError('기존 결과 보호: 비어 있는 새 출력 폴더를 지정하세요.')
        out.mkdir(parents=True,exist_ok=True)
        if a.snapshot:
            snapshot=load_snapshot(a.snapshot)
        elif a.demo:
            from sample_data import demo
            snapshot,supplement=demo(a.end_year or 2025)
            save_json(supplement,out/'demo_supplement.json'); a.supplement=out/'demo_supplement.json'
        else:
            name=a.company or input('기업명: ').strip()
            year=a.end_year or int(input('마지막 사업연도 (예: 2025): '))
            if not 2018<=year<=datetime.now().year:
                raise ValueError('마지막 사업연도는 2018년부터 현재 연도까지 입력하세요.')
            load_dotenv(ROOT/'.env',override=False)
            client=DartClient(os.getenv('OPENDART_API_KEY',''))
            try:
                company=select_company(client.companies(),name,a.identifier)
                snapshot=collect(client,company,year,out,a.allow_ofs,download_notes=not a.skip_documents)
            finally:
                client.close()
        save_json(snapshot,out/'raw_snapshot.json')
        notes=load_supplement(a.supplement)
        observations=analyze(snapshot,notes)
        save_json({'company':snapshot['company'],'synthetic':snapshot['synthetic'],'records':[
            {k:v for k,v in o.items() if k not in ('mapping','raw')} for o in observations]},out/'analysis.json')
        from excel_report import write_report
        path=write_report(snapshot,observations,out/'financial_analysis.xlsx')
        from charts import generate
        try:
            generate(observations[1:],out/'charts',snapshot['synthetic'],a.font)
        except ValueError:
            print('Excel 생성 완료. PNG용 한글 글꼴이 없습니다. --font 옵션으로 다시 실행하세요.')
        print('Excel 저장: '+str(path.resolve()))
        missing=sum(r['value'] is None for o in observations[1:] for r in o['ratios'].values())
        print(f'3개년 15개 결과 중 {15-missing}개 계산 완료, {missing}개 자료 확인 필요')
        return 0
    except DartError as exc:
        print(str(exc)); return 1
    except (ValueError,TypeError,KeyError,EOFError):
        print('입력/자료 형식을 확인하세요. 기업·연도·보완 JSON·스냅샷 또는 출력 폴더가 올바르지 않습니다. README의 예제를 참고하세요.'); return 1
    except OSError:
        print('파일을 읽거나 저장하지 못했습니다. 경로·권한과 Excel 파일이 열려 있는지 확인하세요.'); return 1


if __name__=='__main__':
    sys.exit(main())
