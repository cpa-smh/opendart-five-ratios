"""결측 지표는 NaN으로 남겨 선을 끊는다."""
import math
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from ratios import METRICS


def generate(records, folder, synthetic=False, font=None):
    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = next((n for n in ('Malgun Gothic','Noto Sans CJK KR','AppleGothic') if n in available), None)
    if font:
        font_manager.fontManager.addfont(str(font))
        chosen = font_manager.FontProperties(fname=str(font)).get_name()
    if not chosen:
        raise ValueError('한글 글꼴이 없습니다. --font에 한글 TTF/OTF 경로를 지정하세요.')
    plt.rcParams.update({'font.family':chosen,'axes.unicode_minus':False})
    folder=Path(folder); folder.mkdir(parents=True,exist_ok=True)
    files=[]
    for key,(label,unit,_) in METRICS.items():
        values=[r['ratios'][key]['value'] for r in records]
        values=[float(v)*(100 if unit=='%' else 1) if v is not None else math.nan for v in values]
        fig,ax=plt.subplots(figsize=(7,4),layout='constrained')
        ax.plot([str(r['year']) for r in records],values,marker='o',color='#205493')
        ax.set_title(('[가상 데모] ' if synthetic else '')+label)
        ax.set_ylabel(unit); ax.spines[['right','top']].set_visible(False)
        if all(math.isnan(v) for v in values):
            ax.text(.5,.5,'필수 자료 미확인 — 계산 불가',ha='center',transform=ax.transAxes)
        path=folder/f'{key}.png'; fig.savefig(path,dpi=150); plt.close(fig); files.append(path)
    return files
