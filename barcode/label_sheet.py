# ╔══════════════════════════════════════════════════════╗
# ║   폼텍 라벨지 PDF 생성 모듈 (label_sheet.py)   ║
# ╚══════════════════════════════════════════════════════╝
#
# 무엇을 하는 파일인가?
#   배대지에서 바코드를 붙일 때 "어떤 상품에 붙이는 라벨인지" 한눈에 보이도록
#   폼텍 라벨지(3102 소형 / 3218 대형) 칸마다
#       [상품 이미지] + [옵션1 / 옵션2] + [Code128 바코드]
#   를 넣고, 수량 열에 적힌 개수만큼 같은 라벨을 반복해 A4 PDF로 만든다.
#
# 왜 PDF인가?
#   업로드된 .doc 양식은 결국 "A4 위에 칸을 어디에 그릴지"를 정한 표일 뿐이다.
#   같은 좌표로 PDF를 만들면 어느 PC에서 열어도 칸이 밀리지 않고,
#   워드처럼 폰트/버전에 따라 줄이 바뀌는 일이 없다. (인쇄 시 배율 100% 필수)
#
# 구조 (위에서 아래로 읽으면 흐름이 이어진다)
#   1. FORMS            : 양식별 물리 치수 (.doc 파일에서 직접 읽어낸 값)
#   2. 엑셀 이미지 추출  : B열 그림을 행 번호 → 이미지 바이트로 매핑
#   3. 라벨 항목 수집    : 엑셀 한 행 → LabelItem (이미지·옵션·바코드·수량)
#   4. PDF 렌더링        : 칸 좌표 계산 + 그림/글자/바코드 그리기

import io
import re
import posixpath
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from functools import lru_cache
from typing import Optional

from PIL import Image
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.graphics.barcode.code128 import Code128


# 배포 버전 — 앱 화면에 그대로 찍는다.
# Streamlit Cloud 는 재배포가 늦거나 건너뛰는 일이 있어서, 지금 도는 코드가
# 최신인지 눈으로 확인할 방법이 필요하다. 이 기능을 고칠 때마다 함께 올린다.
VERSION = '2026-09-17d · 줄 단위 배치'


# ══════════════════════════════════════════════════════
# 1. 양식 사양
# ══════════════════════════════════════════════════════
# 값의 출처: 업로드된 3218.doc / 3102.doc 의 표 속성(sprmTDefTable, 행 높이,
# 페이지 여백)을 그대로 mm 로 환산했다. 칸 사이 간격은 0 (표 셀이 맞붙어 있음).
#   3102 : 4열 × 10행 = 40칸, 칸 49.5 × 26.85mm, 왼쪽 6mm / 위 14.2mm 여백
#   3218 : 3열 ×  6행 = 18칸, 칸 66.0 × 45.0mm,  왼쪽 6mm / 위 10.0mm 여백
FORMS = {
    '3102': dict(
        title='폼텍 3102 (소형 · 49.5×27mm · 4×10칸)',
        cols=4, rows=10, label_w=49.5, label_h=26.85,
        left=6.0, top=14.2,
        # 아래는 칸 안쪽 배치 값 (mm / pt). 소형이라 여백을 최소로 잡았다.
        pad=1.6,          # 칸 테두리에서 내용까지 안쪽 여백
        img_size=13.0,    # 상품 이미지 정사각형 한 변
        img_gap=1.3,      # 이미지와 글자 사이 간격
        bc_h=7.2,         # 바코드 막대 높이
        num_size=5.4,     # 바코드 숫자 글자 크기(pt)
        name_size=5.6, name_lines=1,   # 소형은 1줄 (2줄이면 제조국·연령 문구가 밀려난다)
        opt_size=6.2,
        fixed_size=4.3,   # 제조국/연령 등 고정 문구
        max_bar_width=0.42,   # 바코드 모듈 최대 폭(mm) — 너무 굵으면 칸을 넘친다
    ),
    '3218': dict(
        title='폼텍 3218 (대형 · 66×45mm · 3×6칸)',
        cols=3, rows=6, label_w=66.0, label_h=45.0,
        left=6.0, top=10.0,
        pad=2.2,
        img_size=17.0,
        img_gap=1.8,
        bc_h=9.5,
        num_size=6.8,
        name_size=7.4, name_lines=2,
        opt_size=7.2,
        fixed_size=4.8,
        max_bar_width=0.5,
    ),
}


# ══════════════════════════════════════════════════════
# 2. 엑셀에서 상품 이미지 꺼내기
# ══════════════════════════════════════════════════════
# 엑셀에 그림이 들어가는 방식은 세 가지나 되고, 각각 파일 안 저장 위치가 다르다.
#   (a) 떠 있는 그림   : 드로잉(xl/drawings) — openpyxl 이 ws._images 로 읽어준다
#   (b) 셀 안 그림     : 엑셀 365 "셀에 배치" — xl/richData 에 숨어 있어 직접 파싱
#   (c) WPS DISPIMG    : =DISPIMG("ID_...") 수식 — xl/cellimages.xml 에서 매핑
#   (d) 이미지 URL 문자열 — 셀에 http 주소만 있는 경우, 내려받아 사용
# 넷 다 시도하고, 하나라도 성공하면 그 행의 이미지로 쓴다.

_NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'rv': 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata',
    'rvrel': 'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def _zip_read(zf, path):
    """zip 안 파일을 읽되 없으면 None (경로 앞 '/' 유무 차이도 흡수)."""
    for cand in (path, path.lstrip('/')):
        try:
            return zf.read(cand)
        except KeyError:
            continue
    return None


def _rels_map(zf, rels_path):
    """관계 파일(.rels)을 {rId: 절대경로} 로 바꿔준다."""
    data = _zip_read(zf, rels_path)
    out = {}
    if not data:
        return out
    base = posixpath.dirname(posixpath.dirname(rels_path))  # _rels 의 상위 폴더
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return out
    for rel in root.findall('rel:Relationship', _NS):
        rid, target = rel.get('Id'), rel.get('Target', '')
        if not rid or not target:
            continue
        if target.startswith('/'):
            out[rid] = target.lstrip('/')
        else:
            out[rid] = posixpath.normpath(posixpath.join(base, target))
    return out


def _sheet_xml_path(zf, sheet_title):
    """시트 이름 → 그 시트의 xml 경로 (xl/worksheets/sheetN.xml)."""
    wb_xml = _zip_read(zf, 'xl/workbook.xml')
    if not wb_xml:
        return None
    rels = _rels_map(zf, 'xl/_rels/workbook.xml.rels')
    try:
        root = ET.fromstring(wb_xml)
    except ET.ParseError:
        return None
    for sh in root.iter('{%s}sheet' % _NS['main']):
        if sh.get('name') == sheet_title:
            rid = sh.get('{%s}id' % _NS['r'])
            return rels.get(rid)
    return None


def _images_from_drawings(ws, col_image):
    """(a) 떠 있는 그림 — openpyxl 이 앵커(어느 셀 위에 놓였는지)까지 준다."""
    exact, any_col = {}, {}
    for im in getattr(ws, '_images', []) or []:
        frm = getattr(getattr(im, 'anchor', None), '_from', None)
        if frm is None:
            continue
        row, col = frm.row + 1, frm.col + 1       # openpyxl 앵커는 0부터 센다
        try:
            data = im._data()
        except Exception:
            continue
        if not data:
            continue
        if col == col_image:
            exact.setdefault(row, data)
        any_col.setdefault(row, data)
    return exact, any_col


def _images_from_richdata(zf, ws, col_image):
    """(b) 엑셀 365 '셀에 배치' 그림.

    저장 구조가 4단계나 이어진다:
      셀 <c vm="N">  →  metadata.xml (valueMetadata N번째 → rvb i)
                     →  rdrichvalue.xml (i번째 rv 의 첫 <v> = 관계 인덱스)
                     →  richValueRel.xml (인덱스 → rId)
                     →  richValueRel.xml.rels (rId → xl/media/imageX.png)
    중간 단계가 없거나 깨진 파일은 'vm 번호 = rv 번호' 로 가정하고 진행한다.
    """
    sheet_path = _sheet_xml_path(zf, ws.title)
    sheet_xml = _zip_read(zf, sheet_path) if sheet_path else None
    rv_xml = _zip_read(zf, 'xl/richData/rdrichvalue.xml')
    rel_xml = _zip_read(zf, 'xl/richData/richValueRel.xml')
    if not (sheet_xml and rv_xml and rel_xml):
        return {}

    # rId → 실제 이미지 경로
    rel_targets = _rels_map(zf, 'xl/richData/_rels/richValueRel.xml.rels')
    try:
        rel_root = ET.fromstring(rel_xml)
    except ET.ParseError:
        return {}
    rel_ids = [el.get('{%s}id' % _NS['r']) for el in rel_root
               if el.tag.endswith('}rel')]

    # rv 인덱스 → 관계 인덱스 (첫 번째 <v> 값)
    try:
        rv_root = ET.fromstring(rv_xml)
    except ET.ParseError:
        return {}
    rv_rel_idx = []
    for rv in rv_root:
        if not rv.tag.endswith('}rv'):
            continue
        vs = [v.text for v in rv if v.tag.endswith('}v')]
        try:
            rv_rel_idx.append(int(vs[0]))
        except (IndexError, TypeError, ValueError):
            rv_rel_idx.append(None)

    # vm 번호 → rv 인덱스 (metadata.xml). 없으면 1:1 로 본다.
    vm_to_rv = {}
    meta_xml = _zip_read(zf, 'xl/metadata.xml')
    if meta_xml:
        try:
            meta = ET.fromstring(meta_xml)
            fut_rvb = []
            for fm in meta.iter('{%s}futureMetadata' % _NS['main']):
                if fm.get('name') != 'XLRICHVALUE':
                    continue
                for bk in fm.iter('{%s}bk' % _NS['main']):
                    rvb = next(bk.iter('{%s}rvb' % _NS['rv']), None)
                    fut_rvb.append(int(rvb.get('i')) if rvb is not None else None)
            vm = meta.find('main:valueMetadata', _NS)
            if vm is not None:
                for n, bk in enumerate(vm.findall('main:bk', _NS), start=1):
                    rc = bk.find('main:rc', _NS)
                    if rc is None:
                        continue
                    v = int(rc.get('v', '0'))
                    if 0 <= v < len(fut_rvb) and fut_rvb[v] is not None:
                        vm_to_rv[n] = fut_rvb[v]
        except (ET.ParseError, ValueError):
            vm_to_rv = {}

    # 시트에서 vm 속성이 붙은 셀 찾기
    out = {}
    col_letter = get_column_letter(col_image)
    try:
        sroot = ET.fromstring(sheet_xml)
    except ET.ParseError:
        return {}
    for c in sroot.iter('{%s}c' % _NS['main']):
        vm = c.get('vm')
        ref = c.get('r', '')
        if not vm or not ref:
            continue
        m = re.match(r'([A-Z]+)(\d+)$', ref)
        if not m or m.group(1) != col_letter:
            continue
        row = int(m.group(2))
        try:
            vm_n = int(vm)
            rv_i = vm_to_rv.get(vm_n, vm_n - 1)
            rel_i = rv_rel_idx[rv_i]
            target = rel_targets.get(rel_ids[rel_i])
            data = _zip_read(zf, target) if target else None
        except (IndexError, TypeError, ValueError):
            data = None
        if data:
            out.setdefault(row, data)
    return out


def _dispimg_map(zf):
    """(c) WPS 의 DISPIMG ID → 이미지 바이트 매핑 (xl/cellimages.xml)."""
    xml = _zip_read(zf, 'xl/cellimages.xml')
    if not xml:
        return {}
    rels = _rels_map(zf, 'xl/_rels/cellimages.xml.rels')
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return {}
    out = {}
    for pic in root.iter('{%s}pic' % _NS['xdr']):
        cnv = next(pic.iter('{%s}cNvPr' % _NS['xdr']), None)
        blip = next(pic.iter('{%s}blip' % _NS['a']), None)
        if cnv is None or blip is None:
            continue
        name = cnv.get('name')
        rid = blip.get('{%s}embed' % _NS['r'])
        target = rels.get(rid)
        data = _zip_read(zf, target) if target else None
        if name and data:
            out[name] = data
    return out


def _fetch_url_image(url, cache, timeout=10):
    """(d) 셀에 적힌 이미지 URL 내려받기 — 같은 주소는 한 번만."""
    if url in cache:
        return cache[url]
    data = None
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 barcode-app'})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
    except Exception:
        data = None
    cache[url] = data
    return data


def collect_row_images(xlsx_bytes, ws, col_image, start_row):
    """행 번호 → 이미지 바이트. 위 (a)~(d) 를 순서대로 시도한다."""
    exact, any_col = _images_from_drawings(ws, col_image)
    rich, disp = {}, {}
    try:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            rich = _images_from_richdata(zf, ws, col_image)
            disp = _dispimg_map(zf)
    except (zipfile.BadZipFile, OSError):
        pass

    url_cache = {}
    out = {}
    for r in range(start_row, ws.max_row + 1):
        data = exact.get(r) or rich.get(r)
        if not data:
            val = ws.cell(r, col_image).value
            sval = str(val).strip() if val is not None else ''
            m = re.search(r'DISPIMG\(\s*"([^"]+)"', sval, flags=re.I)
            if m:
                data = disp.get(m.group(1))
            elif sval.lower().startswith(('http://', 'https://')):
                data = _fetch_url_image(sval, url_cache)
        if not data:
            data = any_col.get(r)      # 열이 살짝 어긋나게 놓인 그림도 구제
        if data:
            out[r] = data
    return out


# ══════════════════════════════════════════════════════
# 3. 라벨 항목 수집
# ══════════════════════════════════════════════════════
@dataclass
class LabelItem:
    row: int
    name: str
    barcode: str
    opt1: str = ''
    opt2: str = ''
    material: str = ''
    qty: int = 1
    image: Optional[Image.Image] = None
    notes: list = field(default_factory=list)   # 이 행에서 생긴 경고들


def _to_pil(data, max_px=420):
    """이미지 바이트 → RGB PIL (투명 배경은 흰색으로, 크기는 인쇄용으로 축소)."""
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return None
    if im.mode in ('RGBA', 'LA', 'P'):
        im = im.convert('RGBA')
        bg = Image.new('RGB', im.size, 'white')
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != 'RGB':
        im = im.convert('RGB')
    # 라벨 위 이미지는 기껏해야 2cm — 큰 원본을 그대로 넣으면 PDF 만 무거워진다
    if max(im.size) > max_px:
        im.thumbnail((max_px, max_px), Image.LANCZOS)
    return im


def _cell_text(ws, r, col):
    if not col:
        return ''
    v = ws.cell(r, col).value
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _parse_qty(raw):
    """'12', '12개', '12.0', ' 12 ' 모두 12 로. 못 읽으면 None."""
    if raw is None:
        return None
    s = str(raw).strip().replace(',', '')
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    if not m:
        return None
    try:
        return int(float(m.group(0)))
    except ValueError:
        return None


def collect_label_items(xlsx_bytes, settings):
    """엑셀 → LabelItem 목록.

    settings 키:
      start_row, col_name, col_barcode, col_material,
      col_image, col_opt1, col_opt2, col_qty
    반환: (items, warnings)
    """
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    start_row = int(settings['start_row'])
    col_image = int(settings['col_image'])

    row_images = collect_row_images(xlsx_bytes, ws, col_image, start_row)

    items, warnings = [], []
    pil_cache = {}          # 같은 이미지 바이트는 한 번만 디코딩
    for r in range(start_row, ws.max_row + 1):
        bc = _cell_text(ws, r, settings['col_barcode'])
        if not bc:
            continue
        item = LabelItem(
            row=r,
            name=_cell_text(ws, r, settings['col_name']),
            barcode=bc,
            opt1=_cell_text(ws, r, settings.get('col_opt1')),
            opt2=_cell_text(ws, r, settings.get('col_opt2')),
            material=_cell_text(ws, r, settings.get('col_material')),
        )
        qty_raw = ws.cell(r, settings['col_qty']).value if settings.get('col_qty') else None
        qty = _parse_qty(qty_raw)
        if qty is None:
            item.qty = 1
            item.notes.append('수량 없음 → 1장')
        elif qty <= 0:
            item.qty = 0
            item.notes.append('수량 0 → 건너뜀')
        else:
            item.qty = qty

        data = row_images.get(r)
        if data:
            key = hash(data)
            if key not in pil_cache:
                pil_cache[key] = _to_pil(data)
            item.image = pil_cache[key]
            if item.image is None:
                item.notes.append('이미지 파일을 열 수 없음')
        else:
            item.notes.append('이미지 없음')

        if item.notes:
            warnings.append(f'{r}행 ({item.name[:20] or bc}): ' + ', '.join(item.notes))
        if item.qty > 0:
            items.append(item)
    return items, warnings


# ══════════════════════════════════════════════════════
# 4. PDF 렌더링
# ══════════════════════════════════════════════════════
def _font_names(bold='NanumBold', reg='NanumReg'):
    """(굵은 한글, 보통 한글, 굵은 중국어, 보통 중국어) 이름을 돌려준다.

    한글 폰트가 없으면 기본 폰트로라도 그리고, 중국어 폰트가 없으면 None.
    """
    have = set(pdfmetrics.getRegisteredFontNames())
    return (bold if bold in have else 'Helvetica-Bold',
            reg if reg in have else 'Helvetica',
            'NotoSCBold' if 'NotoSCBold' in have else None,
            'NotoSCReg' if 'NotoSCReg' in have else None)


# ── 한글·한자 섞인 글자 그리기 ─────────────────────────
# 1688 옵션값은 '均码-黑色-左' 처럼 한자가 섞인다. 나눔고딕에는 한자가 없어서
# 그냥 그리면 글자가 빠진 채 부호만 남는다. 그래서 글자 하나하나 확인해서
# 나눔고딕에 있으면 나눔고딕으로, 없으면 중국어 폰트로 나눠 그린다.
@lru_cache(maxsize=8192)
def _font_has_char(font_name, cp):
    """그 폰트에 이 글자가 실제로 들어 있는가."""
    if not font_name:
        return False
    try:
        face = pdfmetrics.getFont(font_name).face
        table = getattr(face, 'charToGlyph', None)
        if table is None:          # 내장 기본 폰트는 확인할 방법이 없다 — 있다고 본다
            return True
        return cp in table
    except Exception:
        return True


def _font_runs(text, primary, cn_font):
    """글자열을 [(폰트, 이어지는 글자들), ...] 로 쪼갠다."""
    runs = []
    for ch in text:
        f = primary
        if cn_font and not _font_has_char(primary, ord(ch)) and _font_has_char(cn_font, ord(ch)):
            f = cn_font
        if runs and runs[-1][0] == f:
            runs[-1][1] += ch
        else:
            runs.append([f, ch])
    return runs


def _mixed_width(text, primary, cn_font, size):
    return sum(pdfmetrics.stringWidth(t, f, size) for f, t in _font_runs(text, primary, cn_font))


def _draw_mixed_centred(c, text, cx, y, primary, cn_font, size):
    """가운데 정렬로 그린다. 폰트가 바뀌는 구간마다 나눠 이어 붙인다."""
    runs = _font_runs(text, primary, cn_font)
    total = sum(pdfmetrics.stringWidth(t, f, size) for f, t in runs)
    x = cx - total / 2
    for f, t in runs:
        c.setFont(f, size)
        c.drawString(x, y, t)
        x += pdfmetrics.stringWidth(t, f, size)


def _ellipsis_mixed(text, primary, cn_font, size, max_w):
    if _mixed_width(text, primary, cn_font, size) <= max_w:
        return text
    out = text
    while out and _mixed_width(out + '…', primary, cn_font, size) > max_w:
        out = out[:-1]
    return out + '…'


def _wrap(text, font, size, max_w):
    """글자 단위 줄바꿈. 한글은 띄어쓰기가 드물어 단어 단위로는 못 자른다.
    공백이 있으면 단어 경계를 우선 쓰고, 한 단어가 너무 길면 글자로 쪼갠다."""
    text = ' '.join(str(text).split())      # 연속 공백/개행 정리
    if not text:
        return []
    lines, cur = [], ''
    for word in text.split(' '):
        cand = (cur + ' ' + word) if cur else word
        if pdfmetrics.stringWidth(cand, font, size) <= max_w:
            cur = cand
            continue
        if cur:
            lines.append(cur)
            cur = ''
        # 단어 자체가 한 줄보다 길면 글자 단위로
        piece = ''
        for ch in word:
            if pdfmetrics.stringWidth(piece + ch, font, size) <= max_w:
                piece += ch
            else:
                if piece:
                    lines.append(piece)
                piece = ch
        cur = piece
    if cur:
        lines.append(cur)
    return lines


def _ellipsis(text, font, size, max_w):
    """한 줄에 안 들어가면 뒤를 잘라 '…' 로."""
    if pdfmetrics.stringWidth(text, font, size) <= max_w:
        return text
    out = text
    while out and pdfmetrics.stringWidth(out + '…', font, size) > max_w:
        out = out[:-1]
    return out + '…'


def _draw_lines(c, lines, x, y_top, font, size, leading, max_lines=None):
    """y_top(윗선)부터 아래로 줄을 그린다. 그린 뒤의 y(다음 윗선) 를 돌려준다."""
    if max_lines is not None:
        lines = lines[:max_lines]
    c.setFont(font, size)
    y = y_top
    for ln in lines:
        c.drawString(x, y - size * 0.82, ln)   # 베이스라인 보정 (한글 폰트 기준)
        y -= leading
    return y


def _draw_barcode(c, value, x, y, avail_w, bar_h, max_bar_width_mm):
    """벡터 Code128 을 avail_w 안에 가운데 정렬로 그리고 실제 폭(pt)을 돌려준다."""
    probe = Code128(value, barWidth=1, barHeight=bar_h, quiet=0, humanReadable=False)
    modules = probe.width            # barWidth=1 일 때 폭 = 모듈 수
    bw = min(max_bar_width_mm * mm, avail_w / modules)
    bc = Code128(value, barWidth=bw, barHeight=bar_h, quiet=0, humanReadable=False)
    bc.drawOn(c, x + (avail_w - bc.width) / 2, y)
    return bc.width, bw


def _draw_label(c, form, item, x, y, opts, fonts):
    """칸 하나 그리기. (x, y) 는 칸의 왼쪽 아래 모서리 (PDF 좌표계)."""
    F = form
    f_bold, f_reg, f_cn_bold, f_cn_reg = fonts
    W, H = F['label_w'] * mm, F['label_h'] * mm
    pad = F['pad'] * mm
    img_sz = F['img_size'] * mm

    if opts.get('border'):
        c.setLineWidth(0.3)
        c.setStrokeColorRGB(0.75, 0.75, 0.75)
        c.rect(x, y, W, H)

    # ── 아래쪽: 바코드 + 숫자 ─────────────────────────
    num_size = F['num_size']
    num_h = num_size * 1.15
    bc_y = y + pad + num_h
    bc_avail_w = W - pad * 2
    bar_h = F['bc_h'] * mm
    try:
        _draw_barcode(c, item.barcode, x + pad, bc_y, bc_avail_w, bar_h, F['max_bar_width'])
    except Exception:
        c.setFont(f_reg, num_size)
        c.drawCentredString(x + W / 2, bc_y + bar_h / 2, '(바코드 생성 실패)')
    c.setFont(f_reg, num_size)
    c.drawCentredString(x + W / 2, y + pad + num_size * 0.2, item.barcode)

    # ── 위쪽: 이미지(왼쪽) + 글자(오른쪽) ──────────────
    top = y + H - pad
    zone_bottom = bc_y + bar_h + 0.6 * mm          # 바코드 위 여유
    zone_h = top - zone_bottom
    img_box = min(img_sz, zone_h)
    text_x = x + pad
    if item.image is not None:
        try:
            c.drawImage(ImageReader(item.image), x + pad, top - img_box,
                        width=img_box, height=img_box,
                        preserveAspectRatio=True, anchor='c', mask=None)
        except Exception:
            pass
        text_x = x + pad + img_box + F['img_gap'] * mm
    else:
        # 이미지가 없으면 빈 네모 대신 그 자리를 글자에 내준다
        pass
    text_w = x + W - pad - text_x

    lines_spec = []     # (텍스트 줄 목록, 폰트, 크기, 최대 줄 수)
    if item.name and opts.get('show_name', True):
        lines_spec.append((_wrap(item.name, f_bold, F['name_size'], text_w),
                           f_bold, F['name_size'], F['name_lines']))
    for label, val in (('옵션1', item.opt1), ('옵션2', item.opt2)):
        if val:
            txt = _ellipsis(f'{label}: {val}', f_reg, F['opt_size'], text_w)
            lines_spec.append(([txt], f_reg, F['opt_size'], 1))
    if item.material and opts.get('show_material', True):
        txt = _ellipsis(f'재질: {item.material}', f_reg, F['opt_size'] - 0.6, text_w)
        lines_spec.append(([txt], f_reg, F['opt_size'] - 0.6, 1))

    yy = top
    for lines, font, size, max_lines in lines_spec:
        leading = size * 1.18
        n = min(len(lines), max_lines)
        if yy - leading * n < zone_bottom:          # 자리가 모자라면 남은 줄은 포기
            n = max(0, int((yy - zone_bottom) // leading))
            if n == 0:
                break
        yy = _draw_lines(c, lines[:n], text_x, yy, font, size, leading)

    # ── 고정 문구(제조국·연령·취급주의 …) ──────────────
    # 글자 영역 아래 남는 자리에 넣되, 이미지보다 아래로 내려가면
    # 칸 전체 폭을 쓸 수 있으므로 그때부터는 왼쪽 끝부터 채운다.
    fixed = [t for t in (opts.get('fixed_lines') or []) if t and str(t).strip()]
    if fixed:
        size = F['fixed_size']
        joined = '  ·  '.join(str(t).strip() for t in fixed)
        img_bottom = top - img_box if item.image is not None else top
        # 이미지 옆 좁은 영역과 이미지 아래 넓은 영역, 두 구간으로 나눠 채운다
        for _ in range(4):                       # 안 들어가면 글자를 조금씩 줄인다
            leading = size * 1.15
            remaining = joined
            yy2 = yy
            drawn_all = False
            # 구간 1: 이미지 옆 (좁음)
            narrow_lines = []
            if yy2 > img_bottom + leading:
                cnt = int((yy2 - max(img_bottom, zone_bottom)) // leading)
                narrow_lines = _wrap(remaining, f_reg, size, text_w)
                if cnt < len(narrow_lines):
                    used, rest = narrow_lines[:cnt], ' '.join(narrow_lines[cnt:])
                else:
                    used, rest = narrow_lines, ''
                remaining = rest
                narrow_lines = used
            # 구간 2: 이미지 아래 (칸 전체 폭)
            wide_start = min(yy2 - leading * len(narrow_lines), img_bottom)
            wide_lines = _wrap(remaining, f_reg, size, W - pad * 2) if remaining else []
            wide_cnt = int((wide_start - zone_bottom) // leading)
            if len(wide_lines) <= max(wide_cnt, 0):
                drawn_all = True
            if drawn_all or size <= 3.6:
                if narrow_lines:
                    yy2 = _draw_lines(c, narrow_lines, text_x, yy2, f_reg, size, leading)
                if wide_lines:
                    _draw_lines(c, wide_lines[:max(wide_cnt, 0)], x + pad, wide_start,
                                f_reg, size, leading)
                break
            size -= 0.4


# ── 칸 하나 = 상품 사진 (구분용) ──────────────────────
def _draw_photo_cell(c, form, item, x, y, opts, fonts):
    """구분용 칸 — 상품 사진과 옵션1/옵션2만 넣는다.

    이 라벨지는 중국 현지(배대지) 작업자가 보고 어떤 상품인지 가려내는 용도라
    한글 상품명은 넣지 않는다. 사진과 옵션값(색상·사이즈 등)만 있으면 충분하다.
    옵션 글자는 칸 폭에 맞춰 최대한 크게 키운다.
    """
    F = form
    f_bold, f_reg, f_cn_bold, f_cn_reg = fonts
    W, H = F['label_w'] * mm, F['label_h'] * mm
    pad = F['pad'] * mm

    if opts.get('border'):
        c.setLineWidth(0.3)
        c.setStrokeColorRGB(0.75, 0.75, 0.75)
        c.rect(x, y, W, H)

    texts = [str(t) for t in (item.opt1, item.opt2) if t]

    # 옵션 글자 크기: 되도록 크게 하되 두 가지 한도를 지킨다.
    #   (1) 글자 영역이 칸 높이의 38% 를 넘지 않게 — 사진이 너무 눌리지 않도록
    #   (2) 칸 폭 안에 들어가게 — 긴 옵션값은 크기를 줄여가며 맞춘다
    # 폭을 잴 때도 한자 구간은 중국어 폰트로 재야 실제 그려지는 길이와 맞는다.
    max_w = W - pad * 2
    size = min(F['opt_size'] * 2.2,
               (H * 0.38) / (1.25 * max(len(texts), 1)))
    while size > F['opt_size'] * 0.55:
        if all(_mixed_width(t, f_bold, f_cn_bold, size) <= max_w for t in texts):
            break
        size -= 0.3
    leading = size * 1.25
    band_h = leading * len(texts) + (0.8 * mm if texts else 0)

    # 사진: 남는 자리를 비율 유지로 꽉 채운다
    img_h = H - pad * 2 - band_h
    img_w = W - pad * 2
    if item.image is not None and img_h > 2 * mm:
        iw, ih = item.image.size
        scale = min(img_w / iw, img_h / ih)
        dw, dh = iw * scale, ih * scale
        try:
            c.drawImage(_reader(item.image),
                        x + (W - dw) / 2, y + band_h + pad + (img_h - dh) / 2,
                        width=dw, height=dh, mask=None)
        except Exception:
            pass
    elif item.image is None and img_h > 2 * mm:
        c.setFont(f_reg, F['opt_size'])
        c.drawCentredString(x + W / 2, y + band_h + pad + img_h / 2, '(이미지 없음)')

    # 옵션 글자 (아래에서 위로 쌓음) — 한글/한자 섞여도 글자별로 폰트를 골라 그린다
    yy = y + pad + leading * (len(texts) - 1)
    for t in texts:
        _draw_mixed_centred(c, _ellipsis_mixed(t, f_bold, f_cn_bold, size, max_w),
                            x + W / 2, yy, f_bold, f_cn_bold, size)
        yy -= leading


# ── 칸 하나 = 기존 바코드 라벨 이미지 ──────────────────
def _draw_label_image_cell(c, form, img, x, y, opts, fonts, stretch=False):
    """소형/대형 라벨 생성기가 만든 이미지를 칸에 그대로 넣는다.

    stretch=False : 원본 디자인 비율 그대로, 칸 안에 가운데 맞춤 (기본)
    stretch=True  : 칸 비율에 맞춰 다시 그려진 이미지라 칸을 꽉 채운다
    """
    F = form
    W, H = F['label_w'] * mm, F['label_h'] * mm
    pad = 0.6 * mm                 # 칸 경계에 물리지 않도록 아주 얇은 여백만

    if opts.get('border'):
        c.setLineWidth(0.3)
        c.setStrokeColorRGB(0.75, 0.75, 0.75)
        c.rect(x, y, W, H)

    if img is None:
        c.setFont(fonts[1], F['opt_size'])
        c.drawCentredString(x + W / 2, y + H / 2, '(라벨 생성 실패)')
        return

    avail_w, avail_h = W - pad * 2, H - pad * 2
    if stretch:
        dw, dh = avail_w, avail_h
    else:
        iw, ih = img.size
        scale = min(avail_w / iw, avail_h / ih)
        dw, dh = iw * scale, ih * scale
    try:
        c.drawImage(_reader(img), x + (W - dw) / 2, y + (H - dh) / 2,
                    width=dw, height=dh, mask=None)
    except Exception:
        pass


@lru_cache(maxsize=512)
def _reader_cached(key, img):
    """ImageReader 재사용 — 같은 그림을 여러 칸에 그려도 PDF 에 한 번만 담긴다.

    매번 새 ImageReader 를 만들면 reportlab 이 같은 그림인 줄 모르고
    칸 수만큼 따로 저장해 파일이 몇 배로 부푼다.
    """
    return ImageReader(img)


def _reader(img):
    if img is None:
        return None
    try:
        return _reader_cached(id(img), img)
    except Exception:
        return ImageReader(img)


def _cell_pixels(form, dpi=300):
    """칸 크기(mm) → 인쇄용 픽셀 크기. 라벨 이미지를 이 크기로 그린다."""
    return (max(1, round(form['label_w'] / 25.4 * dpi)),
            max(1, round(form['label_h'] / 25.4 * dpi)))


def create_label_sheet_pdf(items, form_key, opts=None, start_slot=1, label_image_fn=None):
    """LabelItem 목록 → 폼텍 라벨지 PDF (BytesIO).

    form_key       : '3102' 또는 '3218'
    opts           : dict(
                       layout='separate' | 'combined',
                       fixed_lines=[...], show_name=True, show_material=True,
                       border=False, stretch=False, photo_cell=True)
    start_slot     : 첫 장에서 몇 번째 칸부터 채울지 (쓰다 남은 라벨지용, 1부터)
    label_image_fn : (item, px_w, px_h) → PIL 이미지.
                     기존 소형/대형 라벨 생성기를 그대로 불러 쓰라고 밖에서 넘겨준다.
                     (label_sheet 가 barcode_app 을 import 하면 순환 참조가 된다)

    layout='separate' (기본, 배대지 부착용) — 줄 단위 배치:
        상품마다 새 줄에서 시작해 맨 왼쪽 칸에 상품사진, 나머지 칸에 바코드 라벨.
        수량이 남으면 다음 줄을 통째로 바코드로 이어 채운다.

            [사진] [바코드] [바코드] [바코드]     ← 이 상품 시작
            [바코드][바코드][바코드][바코드]     ← 수량이 남으면 다음 줄로
            [사진] [바코드] [바코드]  (빈칸)     ← 다음 상품은 항상 새 줄부터

        줄 맨 왼쪽만 보면 어떤 상품 구간인지 바로 알 수 있어, 수량이 많아도
        작업자가 헷갈리지 않는다. (현장에서 실제로 이렇게 출력해 쓰고 있다)
    layout='combined':
        한 칸에 사진·옵션·바코드를 모두 넣는 예전 방식.

    반환 : (pdf_buf, total_labels, page_count)
    """
    F = FORMS[form_key]
    opts = opts or {}
    layout = opts.get('layout', 'separate')
    stretch = bool(opts.get('stretch'))
    photo_cell = opts.get('photo_cell', True)
    fonts = _font_names()
    cols, rows = F['cols'], F['rows']
    per_page = cols * rows
    page_w, page_h = A4
    lw, lh = F['label_w'] * mm, F['label_h'] * mm
    left, top = F['left'] * mm, F['top'] * mm

    # 칸에 채울 순서를 먼저 만든다 — ('photo', item) / ('label', item) / None(빈칸)
    # 시작 칸 앞은 빈칸으로 메워, 줄 맞춤 계산이 시작 위치까지 고려하게 한다.
    slot0 = max(1, min(int(start_slot), per_page)) - 1
    seq = [None] * slot0
    if layout == 'combined':
        for it in items:
            seq.extend([('combined', it)] * int(it.qty))
    elif not photo_cell:
        for it in items:
            seq.extend([('label', it)] * int(it.qty))
    else:
        for it in items:
            # 상품이 바뀌면 항상 새 줄부터 — 줄 맨 왼쪽이 곧 그 상품의 시작점이 된다
            while len(seq) % cols:
                seq.append(None)
            seq.append(('photo', it))
            seq.extend([('label', it)] * int(it.qty))

    # 바코드 라벨 이미지는 상품당 한 번만 만들어 재사용한다 (수량이 많아도 빠르다)
    px_w, px_h = _cell_pixels(F)
    label_imgs = {}
    if layout != 'combined' and label_image_fn is not None:
        for it in items:
            try:
                label_imgs[id(it)] = label_image_fn(it, px_w if stretch else None,
                                                    px_h if stretch else None)
            except Exception:
                label_imgs[id(it)] = None

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f'바코드 라벨지 {F["title"]}')
    pages = 1 if seq else 0
    used = 0
    for idx, cell in enumerate(seq):
        if idx and idx % per_page == 0:
            c.showPage()
            pages += 1
        if cell is None:          # 줄 맞춤으로 비워 둔 칸
            continue
        kind, it = cell
        used += 1
        pos = idx % per_page
        col, row = pos % cols, pos // cols
        x = left + col * lw
        y = page_h - top - (row + 1) * lh
        if kind == 'photo':
            _draw_photo_cell(c, F, it, x, y, opts, fonts)
        elif kind == 'label':
            _draw_label_image_cell(c, F, label_imgs.get(id(it)), x, y, opts, fonts, stretch)
        else:
            _draw_label(c, F, it, x, y, opts, fonts)
    if not any(cell is not None for cell in seq):
        c.setFont(fonts[1], 12)
        c.drawCentredString(page_w / 2, page_h / 2, '만들 라벨이 없습니다')
    c.save()
    buf.seek(0)
    return buf, used, pages
