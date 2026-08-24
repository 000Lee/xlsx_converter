# -*- coding: utf-8 -*-
"""
한글 문서(.hwp / .hwpx) 표 -> XLSX 일괄 변환

한/글 · 엑셀 설치 불필요. 파일 내부를 직접 파싱해서 병합 / 테두리 /
배경색 / 글꼴 / 글자색 / 굵기 / 취소선 / 열너비 / 행높이를 그대로 옮긴다.

형광펜은 옮기지 않는다. 엑셀에는 글자 단위 형광펜이 없어서 칸 전체를
칠할 수밖에 없는데, 실제 문서는 칸 안의 일부 글자(취소된 이름)에만
형광펜을 치므로 칸을 칠하면 범위를 틀리게 표시하게 된다. 형광펜이 걸린
글자에는 취소선도 함께 걸려 있어 정보 손실도 사실상 없다.

.hwp 는 OLE 복합문서(바이너리), .hwpx 는 ZIP 안의 XML(OWPML)로 구조가
전혀 다르다. 확장자가 아니라 파일 앞 4바이트로 판별해서 각각 읽은 뒤,
같은 형태로 정규화해서 엑셀 쓰기 부분을 공유한다.

쓰기는 xlsxwriter 를 쓴다. openpyxl 은 문자열을 무조건 t="inlineStr" 로만
저장하는데(cell/_writer.py), 엑셀은 인라인 문자열 안의 리치텍스트 런을
무시하고 마지막 런만 남긴다. 한 셀에 색이 섞인 칸이 통째로 깨지는 원인이라
sharedStrings 에 제대로 기록하는 xlsxwriter 로 써야 한다.

준비:  pip install olefile xlsxwriter
실행:  python hwp2xlsx.py
"""

import os
import re
import sys
import glob
import struct
import zlib
import zipfile
import traceback

import olefile
import xlsxwriter

# 영문 윈도우(cp437 등)에서 한글을 콘솔에 찍다가 UnicodeEncodeError 로
# 죽는 것을 막는다. 화면은 '?' 로 깨져도 로그파일(UTF-8)에는 온전히 남는다.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors='replace')
    except Exception:
        pass

# ============================ 설정 ============================
# 비워두면 이 스크립트(또는 exe)가 있는 폴더에서 .hwp / .hwpx 를 찾는다.
IN_DIR = r""
# 비워두면 "바탕화면\시간표_엑셀변환\<원본폴더이름>" 에 저장한다.
# 공유폴더에서 돌려도 결과물은 항상 실행한 사람 PC 에 쌓인다.
OUT_DIR = r""

# 한/글 전용 글꼴을 PC에 있는 글꼴로 바꾸고 싶을 때만 사용 (비워두면 원본 글꼴명 유지)
FONT_MAP = {
    # "함초롬돋움": "맑은 고딕",
    # "함초롬바탕": "맑은 고딕",
}
# ==============================================================


# ---------------------- HWP 레코드 저수준 ----------------------
def _records(buf):
    """HWP 레코드 스트림 -> (tag, level, payload) 제너레이터"""
    i = 0
    while i + 4 <= len(buf):
        h = struct.unpack('<I', buf[i:i + 4])[0]
        tag, lvl, sz = h & 0x3FF, (h >> 10) & 0x3FF, (h >> 20) & 0xFFF
        i += 4
        if sz == 0xFFF:                       # 확장 길이
            sz = struct.unpack('<I', buf[i:i + 4])[0]
            i += 4
        yield tag, lvl, buf[i:i + sz]
        i += sz


def _rgb(c):
    """HWP COLORREF(0x00BBGGRR) -> 'RRGGBB'. 0xFFFFFFFF 는 '없음'."""
    if c is None or c == 0xFFFFFFFF:
        return None
    return '%02X%02X%02X' % (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF)


# HWP 테두리 굵기 코드 -> mm
_BW = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5,
       0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]


def _side(t, w, c):
    """HWP 테두리(종류, 굵기코드, 색) -> (xlsxwriter 테두리코드, 색 '#RRGGBB')

    xlsxwriter 테두리코드: 0 없음 / 1 Thin / 2 Medium / 3 Dashed / 4 Dotted
    5 Thick / 6 Double / 7 Hair / 8 Medium dashed / 9 Dash dot
    10 Medium dash dot / 11 Dash dot dot / 12 Medium dash dot dot
    """
    if t == 0:                                # 없음
        return 0, None
    mm = _BW[w] if w < len(_BW) else 0.5
    if t in (2, 6):
        style = 8 if mm >= 0.4 else 3         # dashed
    elif t == 3:
        style = 4                             # dotted
    elif t == 4:
        style = 10 if mm >= 0.4 else 9        # dash dot
    elif t == 5:
        style = 12 if mm >= 0.4 else 11       # dash dot dot
    elif t in (8, 9, 10, 11):
        style = 6                             # double
    else:                                     # 1 = 실선
        style = 5 if mm >= 0.7 else (2 if mm >= 0.4 else 1)
    return style, '#' + (_rgb(c) or '000000')


def _font_name(n):
    return FONT_MAP.get(n, n)


# --------------------------- DocInfo ---------------------------
def parse_docinfo(buf):
    """글꼴 / 테두리채우기 / 글자모양 / 문단모양 -> 참조 ID 로 찾는 dict

    HWPX 파서(parse_hwpx)와 같은 형태를 내놓아서 이후 처리를 공유한다.
    테두리는 여기서 이미 (엑셀 테두리코드, 색) 로 변환해 둔다.
    """
    faces, borderfills, charshapes, parashapes = [], {}, {}, {}
    for tag, lvl, d in _records(buf):
        if tag == 19 and len(d) >= 3:                       # FACE_NAME
            ln = struct.unpack('<H', d[1:3])[0]
            faces.append(d[3:3 + ln * 2].decode('utf-16-le', 'replace'))

        elif tag == 20 and len(d) >= 32:                    # BORDER_FILL
            b = {}
            for k, off in (('l', 2), ('r', 8), ('t', 14), ('b', 20)):
                b[k] = _side(d[off], d[off + 1],
                             struct.unpack('<I', d[off + 2:off + 6])[0])
            fill = None
            if len(d) >= 36:
                ft = struct.unpack('<I', d[32:36])[0]
                if (ft & 1) and len(d) >= 48:               # 단색 채우기
                    back = struct.unpack('<I', d[36:40])[0]
                    ptype = struct.unpack('<i', d[44:48])[0]
                    if ptype < 0:                           # 무늬 없음 = 단색
                        fill = _rgb(back)
            b['fill'] = fill
            borderfills[len(borderfills) + 1] = b           # 셀은 1부터 참조

        elif tag == 21 and len(d) >= 56:                    # CHAR_SHAPE
            prop = struct.unpack('<I', d[46:50])[0]
            charshapes[len(charshapes)] = dict(
                face=struct.unpack('<H', d[0:2])[0],        # 아래에서 이름으로 치환
                size=struct.unpack('<i', d[42:46])[0] / 100.0,
                italic=bool(prop & 1),
                bold=bool((prop >> 1) & 1),
                # 밑줄 종류 유효값은 1(아래쪽)/3(위쪽). 2 는 밑줄이 아니며
                # 취소선 글자모양에서 함께 세팅되므로 아래쪽만 인정한다.
                underline=(((prop >> 2) & 3) == 1),
                strike=bool((prop >> 18) & 7),
                color=_rgb(struct.unpack('<I', d[52:56])[0]) or '000000')

        elif tag == 25 and len(d) >= 4:                     # PARA_SHAPE
            parashapes[len(parashapes)] = (struct.unpack('<I', d[0:4])[0] >> 2) & 7

    for cs in charshapes.values():                          # 글꼴 번호 -> 이름
        i = cs['face']
        cs['font'] = faces[i] if i < len(faces) else None

    return borderfills, charshapes, parashapes


# ---------------------------- HWPX ----------------------------
# .hwpx 는 ZIP 안에 XML 이 든 개방형 포맷(OWPML)이다. 바이트 위치가 아니라
# 태그 이름으로 읽으므로 .hwp 바이너리보다 해석이 명확하다.
_HWPX_ALIGN = {'JUSTIFY': 0, 'LEFT': 1, 'RIGHT': 2, 'CENTER': 3,
               'DISTRIBUTE': 4, 'DISTRIBUTE_SPACE': 5}
_HWPX_VALIGN = {'TOP': 0, 'CENTER': 1, 'BOTTOM': 2}


def _mm(s):
    """'0.5 mm' -> 0.5"""
    try:
        return float(str(s).replace('mm', '').strip())
    except ValueError:
        return 0.5


def _side_hwpx(el):
    """<hh:leftBorder type=".." width="0.5 mm" color="#000000"/> -> Side"""
    if el is None:
        return 0, None
    t = (el.get('type') or 'NONE').upper()
    if t in ('NONE', ''):
        return 0, None
    mm = _mm(el.get('width'))
    if 'DASH_DOT_DOT' in t:
        style = 12 if mm >= 0.4 else 11
    elif 'DASH_DOT' in t:
        style = 10 if mm >= 0.4 else 9
    elif 'DOT' in t:
        style = 4
    elif 'DASH' in t:
        style = 8 if mm >= 0.4 else 3
    elif 'DOUBLE' in t or 'THICK' in t and 'SLIM' in t:
        style = 6
    else:                                       # SOLID 등
        style = 5 if mm >= 0.7 else (2 if mm >= 0.4 else 1)
    c = el.get('color') or '#000000'
    return style, (c if c.startswith('#') else '#' + c)


def _q(tag):
    """네임스페이스를 무시하고 지역 이름만 비교하기 위한 헬퍼"""
    return tag.rsplit('}', 1)[-1]


def _find(el, name):
    for c in el.iter():
        if _q(c.tag) == name:
            return c
    return None


def parse_hwpx_header(root):
    """header.xml -> (테두리채우기, 글자모양, 문단모양)  ID 로 찾는 dict"""
    fonts = {}                                   # (언어, id) -> 글꼴 이름
    for ff in root.iter():
        if _q(ff.tag) != 'fontface':
            continue
        lang = ff.get('lang') or ''
        for fo in ff:
            if _q(fo.tag) == 'font':
                fonts[(lang, fo.get('id'))] = fo.get('face')

    borderfills, charshapes, parashapes = {}, {}, {}
    for el in root.iter():
        name = _q(el.tag)

        if name == 'borderFill':
            b = {}
            for key, tag in (('l', 'leftBorder'), ('r', 'rightBorder'),
                             ('t', 'topBorder'), ('b', 'bottomBorder')):
                b[key] = _side_hwpx(_find(el, tag))
            fill = None
            brush = _find(el, 'winBrush')
            if brush is not None:
                fc = brush.get('faceColor')
                if fc and fc.lower() not in ('none', ''):
                    fill = fc.lstrip('#').upper()
            b['fill'] = fill
            borderfills[el.get('id')] = b

        elif name == 'charPr':
            ul = _find(el, 'underline')
            so = _find(el, 'strikeout')
            ref = _find(el, 'fontRef')
            font = None
            if ref is not None:
                font = fonts.get(('HANGUL', ref.get('hangul'))) \
                    or fonts.get(('LATIN', ref.get('latin')))
            color = (el.get('textColor') or '#000000').lstrip('#').upper()
            charshapes[el.get('id')] = dict(
                font=font,
                size=float(el.get('height') or 1000) / 100.0,
                bold=_find(el, 'bold') is not None,
                italic=_find(el, 'italic') is not None,
                # underline 은 type 이 NONE 이 아닐 때만 밑줄
                underline=bool(ul is not None
                               and (ul.get('type') or 'NONE').upper() not in ('NONE', '')),
                # strikeout 은 shape 이 실제 선 모양일 때만 취소선.
                # NONE 은 물론이고 3D 도 '긋지 않음'으로 저장되어 있다.
                strike=bool(so is not None
                            and (so.get('shape') or 'NONE').upper() not in ('NONE', '', '3D')),
                color=color if len(color) == 6 else '000000')

        elif name == 'paraPr':
            al = _find(el, 'align')
            parashapes[el.get('id')] = _HWPX_ALIGN.get(
                (al.get('horizontal') or '').upper() if al is not None else '', 0)

    return borderfills, charshapes, parashapes


def _run_text(r):
    """<hp:run> 안의 글자를 뽑는다.

    <hp:t> 안에 형광펜(markpen) 같은 표식이 들어가면 실제 글자가
    <hp:t>.text 가 아니라 그 표식의 tail 에 놓인다. 그래서 하위 노드를
    전부 훑으면서 text 와 tail 을 순서대로 모아야 글자가 누락되지 않는다.
    """
    out = []

    def walk(el):
        if _q(el.tag) == 'lineBreak':
            out.append('\n')
        if el.text:
            out.append(el.text)
        for ch in el:
            walk(ch)
            if ch.tail:
                out.append(ch.tail)

    for t in r:
        if _q(t.tag) == 't':
            walk(t)
    return ''.join(out)


def parse_hwpx_section(root, warn):
    """section*.xml -> [(표정보, 셀목록), ...]"""
    tables = []
    for tbl in root.iter():
        if _q(tbl.tag) != 'tbl':
            continue
        cells = []
        for tc in tbl.iter():
            if _q(tc.tag) != 'tc':
                continue
            addr = _find(tc, 'cellAddr')
            span = _find(tc, 'cellSpan')
            size = _find(tc, 'cellSz')
            sub = _find(tc, 'subList')
            if addr is None:
                continue

            paras = []
            for p in tc.iter():
                if _q(p.tag) != 'p':
                    continue
                runs = []
                for r in p:
                    if _q(r.tag) != 'run':
                        continue
                    txt = _run_text(r)
                    if txt:
                        runs.append((r.get('charPrIDRef'), txt))
                paras.append(dict(shape=p.get('paraPrIDRef'), runs=runs))

            cells.append(dict(
                col=int(addr.get('colAddr') or 0),
                row=int(addr.get('rowAddr') or 0),
                cspan=int(span.get('colSpan') or 1) if span is not None else 1,
                rspan=int(span.get('rowSpan') or 1) if span is not None else 1,
                w=int(size.get('width') or 0) if size is not None else 0,
                h=int(size.get('height') or 0) if size is not None else 0,
                bf=tc.get('borderFillIDRef'),
                valign=_HWPX_VALIGN.get(
                    (sub.get('vertAlign') or '').upper() if sub is not None else '', 1),
                paras=paras))

        if cells:
            tables.append((dict(rows=int(tbl.get('rowCnt') or 0),
                                cols=int(tbl.get('colCnt') or 0)), cells))
    return tables


def read_hwpx(path, warn):
    """.hwpx -> (테두리채우기, 글자모양, 문단모양, 표목록)"""
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        hdr = next((n for n in names if n.lower().endswith('header.xml')), None)
        if hdr is None:
            raise ValueError('header.xml 이 없음 - HWPX 가 아닌 것 같음')
        bfs, css, pss = parse_hwpx_header(ET.fromstring(z.read(hdr)))

        secs = sorted(n for n in names
                      if re.match(r'Contents/section\d+\.xml$', n, re.I))
        tables = []
        for s in secs:
            root = ET.fromstring(z.read(s))
            n_shape = sum(1 for e in root.iter()
                          if _q(e.tag) in ('rect', 'ellipse', 'polygon', 'pic', 'container'))
            if n_shape:
                warn.append('그리기개체/글상자 %d개 감지 - 개체 안의 표는 누락될 수 있음'
                            % n_shape)
            tables.extend(parse_hwpx_section(root, warn))
    return bfs, css, pss, tables


# ------------------------- 문단 텍스트 -------------------------
# 1 WCHAR 를 차지하는 제어문자 (나머지 제어문자는 8 WCHAR)
_CH1 = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}


def para_text(data):
    """PARA_TEXT -> (평문, 각 글자의 WCHAR 위치)"""
    out, pos = [], []
    i, n = 0, len(data) // 2
    while i < n:
        c = struct.unpack('<H', data[i * 2:i * 2 + 2])[0]
        if c < 32:
            if c in _CH1:
                if c == 10:                    # 강제 줄나눔
                    out.append('\n')
                    pos.append(i)
                i += 1
            else:
                if c == 9:                     # 탭
                    out.append(' ')
                    pos.append(i)
                i += 8
        else:
            out.append(chr(c))
            pos.append(i)
            i += 1
    return ''.join(out), pos


# --------------------------- Section ---------------------------
def parse_tables(buf, warn):
    """섹션 안의 모든 표를 파싱해서 [(정보, 셀목록), ...] 반환"""
    recs = list(_records(buf))
    tables = []

    n_shape = sum(1 for t, l, d in recs if t == 76)
    if n_shape:
        warn.append('그리기개체/글상자 %d개 감지 - 개체 안의 표는 누락될 수 있음' % n_shape)

    for ti, (tag, tlvl, d) in enumerate(recs):
        if tag != 77 or len(d) < 8:
            continue
        nrows, ncols = struct.unpack('<HH', d[4:8])
        cells, cur = [], None
        deep_cell = False

        for tag2, lvl, dd in recs[ti + 1:]:
            if tag2 == 71 and lvl < tlvl:      # 다음 컨트롤 = 이 표의 끝
                break
            if tag2 == 77 and lvl > tlvl:      # 중첩 표
                deep_cell = True
                continue

            if tag2 == 72 and lvl == tlvl and len(dd) >= 34:       # 셀
                attr = struct.unpack('<I', dd[4:8])[0]
                ca, ra, cs, rs = struct.unpack('<HHHH', dd[8:16])
                w, h = struct.unpack('<II', dd[16:24])
                cur = dict(col=ca, row=ra, cspan=cs, rspan=rs, w=w, h=h,
                           bf=struct.unpack('<H', dd[32:34])[0],
                           valign=(attr >> 5) & 3, paras=[])
                cells.append(cur)
            elif cur is None:
                continue
            # 셀 문단은 LIST_HEADER 와 같은 레벨, 문단 내용은 한 단계 아래
            elif tag2 == 66 and lvl == tlvl and len(dd) >= 10:
                cur['paras'].append(dict(
                    shape=struct.unpack('<H', dd[8:10])[0],
                    text='', pos=[], runs=[]))
            elif tag2 == 67 and lvl == tlvl + 1 and cur['paras']:
                cur['paras'][-1]['text'], cur['paras'][-1]['pos'] = para_text(dd)
            elif tag2 == 68 and lvl == tlvl + 1 and cur['paras']:
                cur['paras'][-1]['runs'] = [
                    struct.unpack('<II', dd[k:k + 8])
                    for k in range(0, len(dd) - 7, 8)]

        if deep_cell:
            warn.append('중첩된 표 발견 - 안쪽 표 내용은 옮기지 않음')
        if cells:
            for c in cells:                       # HWPX 와 같은 형태로 정규화
                c['paras'] = [dict(shape=p['shape'], runs=_resolve_runs(p))
                              for p in c['paras']]
            tables.append((dict(rows=nrows, cols=ncols), cells))
    return tables


def _resolve_runs(p):
    """PARA_CHAR_SHAPE 의 (시작위치, 글자모양ID) 목록 -> [(글자모양ID, 글자)]"""
    t, pos = p['text'], p['pos']
    runs = p['runs'] or [(0, 0)]
    out = []
    for k, (start, csid) in enumerate(runs):
        end = runs[k + 1][0] if k + 1 < len(runs) else 1 << 30
        seg = ''.join(ch for ch, wp in zip(t, pos) if start <= wp < end)
        if seg:
            out.append((csid, seg))
    return out


# ------------------------- 시트에 쓰기 -------------------------
# HWP 정렬: 0 양쪽 / 1 왼쪽 / 2 오른쪽 / 3 가운데 / 4 배분 / 5 나눔
_HALIGN = ('justify', 'left', 'right', 'center', 'distributed', 'distributed')
# xlsxwriter 는 세로 가운데를 'vcenter' 로 받는다. 'center' 를 주면
# 가로 정렬로만 해석되고 세로는 엑셀 기본값(아래)이 되어 글자가 바닥에 붙는다.
_VALIGN = ('top', 'vcenter', 'bottom')


def _fkey(cs):
    """글자모양 비교용 키 (같은 서식이면 같은 값)"""
    if cs is None:
        return None
    return (cs['font'], cs['size'], cs['bold'], cs['italic'],
            cs['underline'], cs['strike'], cs['color'])


def _font_props(cs):
    """글자모양 -> xlsxwriter 글꼴 속성 (리치텍스트 런에 쓰는 부분)"""
    p = {'font_size': cs['size'], 'font_color': '#' + cs['color']}
    if cs['font']:
        p['font_name'] = _font_name(cs['font'])
    if cs['bold']:
        p['bold'] = True
    if cs['italic']:
        p['italic'] = True
    if cs['underline']:
        p['underline'] = 1
    if cs['strike']:
        p['font_strikeout'] = True
    return p


def _fmt(wb, cache, props):
    """같은 서식은 Format 객체를 재사용 (엑셀 스타일 개수 폭증 방지)"""
    key = tuple(sorted(props.items()))
    f = cache.get(key)
    if f is None:
        f = cache[key] = wb.add_format(props)
    return f


def _cell_props(bf, align, valign):
    """셀 단위 서식: 테두리 / 배경 / 정렬"""
    p = {'text_wrap': True,
         'align': _HALIGN[align] if align is not None and align < 6 else 'center',
         'valign': _VALIGN[valign]}
    if bf:
        for key, side in (('left', 'l'), ('right', 'r'), ('top', 't'), ('bottom', 'b')):
            style, color = bf[side]            # 파서에서 이미 변환해 둔 값
            if style:
                p[key] = style
                p[key + '_color'] = color
        if bf['fill'] and bf['fill'] != 'FFFFFF':
            p['bg_color'] = '#' + bf['fill']
    return p


def write_table(wb, ws, cells, bfs, css, pss, cache):
    colw, rowh, spans = {}, {}, []

    for c in cells:
        r0, c0 = c['row'], c['col']            # xlsxwriter 는 0부터
        bf = bfs.get(c['bf'])

        if c['cspan'] == 1:
            colw[c0] = max(colw.get(c0, 0), c['w'])
        else:
            spans.append((c0, c['cspan'], c['w']))
        if c['rspan'] == 1:
            rowh[r0] = max(rowh.get(r0, 0), c['h'])

        # --- 텍스트 조립 ---
        # 문단 사이 줄바꿈은 "\n 하나만 있는 런"으로 만들지 않는다.
        # OOXML 은 앞뒤 공백이 있는 <t> 에 xml:space="preserve" 를 요구하므로
        # 앞 조각 끝에 붙여서 공백뿐인 런이 생기지 않게 한다.
        segs, align, pending = [], None, ''
        for pi, p in enumerate(c['paras']):
            if pi:
                if segs:
                    segs[-1][1] += '\n'
                else:
                    pending += '\n'
            if align is None:
                align = pss.get(p['shape'])
            for csid, seg in p['runs']:
                if not seg:
                    continue
                segs.append([css.get(csid), pending + seg])
                pending = ''
        if pending and segs:
            segs[-1][1] += pending

        text = ''.join(s for _, s in segs)
        base = _cell_props(bf, align, c['valign'])
        mixed = len({_fkey(cs) for cs, _ in segs}) > 1

        # 서식이 하나뿐이면 그 글꼴을 셀 서식에 합친다.
        # 섞였으면 셀 서식의 글꼴은 건드리지 않고(중립) 런별로 지정한다.
        if segs and not mixed and segs[0][0]:
            base.update(_font_props(segs[0][0]))
        cell_fmt = _fmt(wb, cache, base)

        r1, c1 = r0 + c['rspan'] - 1, c0 + c['cspan'] - 1
        merged = c['rspan'] > 1 or c['cspan'] > 1
        if merged:
            ws.merge_range(r0, c0, r1, c1, '', cell_fmt)

        if not text.strip():
            if not merged:
                ws.write_blank(r0, c0, None, cell_fmt)
        elif mixed:
            frags = []
            for cs, s in segs:
                frags.append(_fmt(wb, cache, _font_props(cs)) if cs else None)
                frags.append(s)
            frags = [x for x in frags if x is not None]
            ws.write_rich_string(r0, c0, *frags, cell_fmt)
        else:
            ws.write_string(r0, c0, text, cell_fmt)

    # 어떤 열은 그 열을 홑겹으로 쓰는 셀이 하나도 없어서(전부 병합 셀에만
    # 걸쳐 있어서) 위에서 너비를 못 정한다. 그대로 두면 엑셀 기본값(8.43)이
    # 되어 그 열만 두 배로 벌어진다. 병합 셀 너비에서 이미 아는 열을 빼고
    # 남은 폭을 모르는 열끼리 나눠 갖게 한다. 좁은 span 부터 처리해야
    # 미지 열이 하나만 남는 경우가 많아 정확해진다.
    for c0, cspan, w in sorted(spans, key=lambda x: x[1]):
        cols = range(c0, c0 + cspan)
        unknown = [i for i in cols if i not in colw]
        if not unknown:
            continue
        rest = w - sum(colw.get(i, 0) for i in cols if i in colw)
        each = max(1, rest // len(unknown))
        for i in unknown:
            colw[i] = each

    # HWPUNIT = 1/7200 inch
    for ci, w in colw.items():
        px = w / 7200.0 * 96.0                              # 96dpi 픽셀
        # 엑셀이 파일에 저장하는 열너비 = (픽셀 - 5) / 7.
        # xlsxwriter 는 넘긴 값에 5/7 을 더해서 저장하므로 미리 빼둔다.
        stored = max(1.0, (px - 5) / 7.0)
        ws.set_column(ci, ci, stored - 5.0 / 7.0)
    for ri, h in rowh.items():
        ws.set_row(ri, max(9.0, h / 100.0))                 # HWPUNIT/100 = pt

    ws.hide_gridlines(2)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


# --------------------------- 변환 1건 ---------------------------
def read_hwp5(src, warn):
    """.hwp (OLE 바이너리) -> (테두리채우기, 글자모양, 문단모양, 표목록)"""
    ole = olefile.OleFileIO(src)
    try:
        hdr = ole.openstream('FileHeader').read()
        if hdr[:17] != b'HWP Document File':
            raise ValueError('HWP 5.0 형식이 아님 (구버전 .hwp 3.0 등)')
        flags = struct.unpack('<I', hdr[36:40])[0]
        if flags & 0x02:
            raise ValueError('암호가 걸린 문서')
        if flags & 0x04:
            raise ValueError('배포용(복사방지) 문서 - 변환 불가')
        compressed = flags & 0x01

        def raw(name):
            b = ole.openstream(name).read()
            return zlib.decompress(b, -15) if compressed else b

        bfs, css, pss = parse_docinfo(raw('DocInfo'))
        sections = sorted(s for s in ('/'.join(x) for x in ole.listdir())
                          if s.startswith('BodyText/Section'))
        tables = []
        for s in sections:
            tables.extend(parse_tables(raw(s), warn))
    finally:
        ole.close()
    return bfs, css, pss, tables


def convert(src, dst):
    warn = []
    with open(src, 'rb') as fp:
        magic = fp.read(8)

    # 확장자가 아니라 실제 내용으로 판별한다 (확장자만 바꿔둔 파일 대비).
    if magic[:4] == b'PK\x03\x04':                        # ZIP -> HWPX
        bfs, css, pss, tables = read_hwpx(src, warn)
    elif magic == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':    # OLE -> HWP 5.0
        bfs, css, pss, tables = read_hwp5(src, warn)
    else:
        raise ValueError('한글 문서가 아님 (.hwp/.hwpx 둘 다 아님)')

    if not tables:
        raise ValueError('표를 찾지 못함')
    if len(tables) > 1:
        warn.append('표가 %d개 - 시트를 나눠서 저장함' % len(tables))

    wb = xlsxwriter.Workbook(dst, {'strings_to_numbers': False,
                                   'strings_to_formulas': False,
                                   'strings_to_urls': False})
    try:
        cache = {}
        for i, (info, cells) in enumerate(tables, 1):
            ws = wb.add_worksheet('표%d' % i if len(tables) > 1 else '표')
            write_table(wb, ws, cells, bfs, css, pss, cache)
    finally:
        wb.close()

    info = tables[0][0]
    return dict(rows=info['rows'], cols=info['cols'],
                cells=len(tables[0][1]), tables=len(tables), warn=warn)


# ----------------------------- 배치 -----------------------------
FROZEN = getattr(sys, 'frozen', False)          # exe 로 빌드된 상태인지


def _pause():
    """exe 로 실행했을 때 창이 바로 닫히지 않게"""
    if FROZEN:
        try:
            input('\n계속하려면 Enter 를 누르세요...')
        except EOFError:
            pass


_LOG = []


def log(msg=''):
    """화면 + 로그파일 양쪽에 남긴다.

    exe 를 더블클릭/드래그드롭으로 실행하면 창이 순식간에 닫히거나
    아무 반응이 없어 보일 때가 있어서, 무슨 일이 있었는지 파일로 남긴다.
    """
    print(msg)
    _LOG.append(msg)


def _flush_log():
    if not FROZEN:
        return
    import datetime
    for d in (os.path.dirname(sys.executable), os.environ.get('TEMP', '')):
        if not d:
            continue
        try:
            with open(os.path.join(d, '변환로그.txt'), 'a', encoding='utf-8') as fp:
                fp.write('\n===== %s =====\n' % datetime.datetime.now()
                         .strftime('%Y-%m-%d %H:%M:%S'))
                fp.write('exe      : %s\n' % sys.executable)
                fp.write('작업폴더 : %s\n' % os.getcwd())
                fp.write('인자     : %r\n' % (sys.argv[1:],))
                fp.write('\n'.join(_LOG) + '\n')
            return
        except Exception:
            continue


def _base_dir():
    """IN_DIR 이 비어 있으면 실행 파일(또는 스크립트)이 놓인 폴더"""
    if IN_DIR:
        return IN_DIR
    return os.path.dirname(os.path.abspath(sys.executable if FROZEN else __file__))


OUT_FOLDER_NAME = '시간표_엑셀변환'


def _my_dir():
    """내 PC 의 바탕화면(없으면 문서) 경로.

    공유폴더에 결과를 쓰면 권한이 없어 실패하거나 남의 폴더를 어지럽히므로
    결과물은 항상 실행한 사람 PC 에 저장한다. OneDrive 로 옮겨진 바탕화면도
    잡히도록 레지스트리에 등록된 실제 경로를 먼저 본다.
    """
    try:
        import winreg
        key = r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            for name in ('Desktop', 'Personal'):     # Personal = 내 문서
                try:
                    p = os.path.expandvars(winreg.QueryValueEx(k, name)[0])
                    if os.path.isdir(p):
                        return p
                except OSError:
                    pass
    except Exception:
        pass
    home = os.path.expanduser('~')
    for p in (os.path.join(home, 'Desktop'), os.path.join(home, 'Documents'), home):
        if os.path.isdir(p):
            return p
    return None


def _out_dir_for(src_dir):
    """결과물을 저장할 폴더. 바탕화면\\시간표_엑셀변환\\<원본폴더이름>"""
    root = _my_dir()
    if not root:
        return os.path.join(src_dir, 'xlsx')         # 최후 수단
    name = os.path.basename(os.path.normpath(src_dir)) or '변환결과'
    return os.path.join(root, OUT_FOLDER_NAME, name)


def _collect(paths):
    """파일/폴더 경로 목록 -> .hwp 파일 목록"""
    files = []
    for pth in paths:
        pth = pth.rstrip('"')
        if os.path.isdir(pth):
            for ext in ('*.hwp', '*.hwpx'):
                files += glob.glob(os.path.join(pth, ext))
        elif pth.lower().endswith(('.hwp', '.hwpx')):
            files.append(pth)
    return sorted(f for f in files if not os.path.basename(f).startswith('~'))


def main():
    args = [a for a in sys.argv[1:] if a.strip()]

    # exe: 인자는 전부 변환 대상(드래그&드롭). 인자가 없으면 exe 옆 폴더.
    # 스크립트: "입력폴더 출력폴더" 두 개를 받는 형태도 허용.
    if not FROZEN and len(args) == 2 and os.path.isdir(args[0]):
        src_dir, out_dir = args[0], args[1]
        files = _collect([args[0]])
    elif args:
        files = _collect(args)
        src_dir = os.path.dirname(os.path.abspath(files[0])) if files else ''
        out_dir = (OUT_DIR or _out_dir_for(src_dir)) if files else ''
    else:
        src_dir = _base_dir()
        files = _collect([src_dir])
        out_dir = OUT_DIR or _out_dir_for(src_dir)

    if not files:
        log('[!] 변환할 .hwp 파일이 없습니다.')
        log('    찾아본 곳 : %s' % (src_dir or '(없음)'))
        log('    받은 인자 : %r' % (args,))
        log('')
        log('    - 이 프로그램과 같은 폴더에 .hwp 파일을 두고 다시 실행하거나')
        log('    - .hwp 파일(또는 폴더)을 이 프로그램 아이콘 위로 끌어다 놓으세요.')
        _flush_log()
        _pause()
        return 1

    os.makedirs(out_dir, exist_ok=True)
    log('입력 : %s' % src_dir)
    log('출력 : %s' % out_dir)
    log('대상 : %d개\n' % len(files))

    ok, failed, warned = 0, [], []
    for n, f in enumerate(files, 1):
        name = os.path.basename(f)
        dst = os.path.join(out_dir, os.path.splitext(name)[0] + '.xlsx')
        try:
            r = convert(f, dst)
            ok += 1
            log('[%3d/%d] OK   %s  (%d행 x %d열, %d셀)'
                % (n, len(files), name, r['rows'], r['cols'], r['cells']))
            for w in r['warn']:
                log('              ! %s' % w)
                warned.append((name, w))
        except Exception as e:
            failed.append((name, '%s: %s' % (type(e).__name__, e)))
            log('[%3d/%d] FAIL %s  -> %s' % (n, len(files), name, e))

    log('\n' + '=' * 60)
    log('성공 %d / 실패 %d' % (ok, len(failed)))
    if warned:
        log('\n[확인 필요]')
        for name, w in warned:
            log('  - %s : %s' % (name, w))
    if failed:
        log('\n[실패 목록]')
        for name, err in failed:
            log('  - %s : %s' % (name, err))
    _flush_log()
    _pause()
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        _LOG.append(traceback.format_exc())
        _flush_log()
        _pause()
        sys.exit(1)
