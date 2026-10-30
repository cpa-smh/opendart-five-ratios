"""공시 ZIP의 해당 감사보고서만 저장한다. 불확실한 표는 보완 입력 대상으로 남긴다."""
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, BadZipFile
from defusedxml import ElementTree as ET
from defusedxml.common import DefusedXmlException
from dart_client import DartError, check_status
from html.parser import HTMLParser


class TextReader(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts=[]

    def handle_data(self, data):
        if data.strip(): self.parts.append(data.strip())


def fetch_notes(client, receipt, fs_div, folder):
    payload = client._get('document.xml', rcept_no=receipt).content
    try:
        if not payload.startswith(b'PK'):
            check_status(ET.fromstring(payload).findtext('status'))
            raise DartError('공시 원문이 ZIP 형식이 아닙니다.')
        with ZipFile(BytesIO(payload)) as archive:
            suffix = '_00761.xml' if fs_div=='CFS' else '_00760.xml'
            names = [n for n in archive.namelist() if n.lower().endswith(suffix)]
            if len(names) != 1:
                raise DartError('해당 연결/별도 감사보고서 원문을 특정할 수 없습니다. 공시 링크에서 주석을 확인하세요.')
            if archive.getinfo(names[0]).file_size > 30_000_000:
                raise DartError('공시 원문 크기 제한을 초과했습니다.')
            data = archive.read(names[0])
        # 원문 XML은 파일로 보존. 실행하거나 HTML로 열지 않는다.
        folder = Path(folder); folder.mkdir(parents=True, exist_ok=True)
        (folder/f'{receipt}_{fs_div}.xml').write_bytes(data)
        # DART 문서는 비정형 XML/HTML 및 EUC-KR 선언을 포함할 수 있다.
        # 외부 엔터티를 해석하지 않는 HTMLParser로 열람용 텍스트만 추출한다.
        try:
            decoded=data.decode('utf-8-sig')
        except UnicodeDecodeError:
            decoded=data.decode('cp949')
        reader=TextReader(); reader.feed(decoded)
        text='\n'.join(reader.parts)
        (folder/f'{receipt}_{fs_div}.txt').write_text(text, encoding='utf-8')
        return '원문 저장 완료. 이자비용·매출채권 주석은 범위 확인 후 --supplement로 입력'
    except (BadZipFile, ET.ParseError, DefusedXmlException, KeyError, UnicodeError):
        raise DartError('공시 원문 ZIP/XML을 읽을 수 없습니다.') from None
