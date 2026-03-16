import math
import re
from typing import Any, Dict, Optional, Tuple
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title='목재 재단 프로그램 Fast v8', layout='wide')
BOARD_PRESETS = {
    '기본 4x8 (1220 x 2440)': (2440.0, 1220.0),
    '4x6 (1220 x 1830)': (1830.0, 1220.0),
    '맞춤 입력': None,
}

def parse_spec(spec_raw: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    if spec_raw is None or (isinstance(spec_raw, float) and math.isnan(spec_raw)):
        return None, None, None
    text = str(spec_raw).strip()
    m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*$', text)
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))

def to_int(v: Any, default: int = 0) -> int:
    try:
        if v is None or v == '' or (isinstance(v, float) and math.isnan(v)):
            return default
        return int(float(v))
    except Exception:
        return default

def is_cutting_target(row: Dict[str, Any], width, height, thickness) -> bool:
    material = str(row.get('재질') or '').strip().upper()
    image_flag = str(row.get('대표이미지') or '').strip().upper()
    qty = to_int(row.get('정소요량'), 0) or to_int(row.get('실소요량'), 0)
    if width is None or height is None or thickness is None:
        return False
    if qty <= 0 or image_flag == 'Y':
        return False
    return not any(word in material for word in ['BOX', '포장', '철물', '경첩'])

def load_bom_from_dataframe(df: pd.DataFrame):
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna('')
    items, errors = [], []
    for idx, row in df.iterrows():
        raw = row.to_dict()
        w, h, t = parse_spec(raw.get('규격'))
        item = {
            'selected': True,
            'row_no': int(idx) + 2,
            'product_code': str(raw.get('품목코드') or '').strip(),
            'part_code': str(raw.get('부품코드') or '').strip(),
            'part_name': str(raw.get('품목명') or '').strip(),
            'color': str(raw.get('색상') or '').strip(),
            'bom_qty': max(1, to_int(raw.get('정소요량'), 0) or to_int(raw.get('실소요량'), 0)),
            'actual_cut_qty': 1,
            'qty': 1,
            'spec_raw': str(raw.get('규격') or '').strip(),
            'width_mm': w,
            'height_mm': h,
            'thickness_mm': t,
            'material_name': str(raw.get('재질') or '').strip(),
            'process_name': str(raw.get('소요공정') or '').strip(),
            'is_cutting_target': is_cutting_target(raw, w, h, t),
        }
        item['qty'] = item['bom_qty'] * item['actual_cut_qty']
        if not item['product_code']:
            errors.append({'row': int(idx)+2, 'field': '품목코드', 'message': '제품코드 누락'})
        if not item['part_code']:
            errors.append({'row': int(idx)+2, 'field': '부품코드', 'message': '부품코드 누락'})
        items.append(item)
    return items, errors

def expand_parts(parts):
    expanded = []
    for p in parts:
        for _ in range(max(1, to_int(p.get('qty'), 1))):
            expanded.append({
                'product_code': p['product_code'], 'part_code': p['part_code'], 'part_name': p['part_name'],
                'color': p['color'], 'width_mm': float(p['width_mm']), 'height_mm': float(p['height_mm']),
                'thickness_mm': float(p['thickness_mm'])
            })
    expanded.sort(key=lambda x: x['width_mm'] * x['height_mm'], reverse=True)
    return expanded

def prune_rects(rects):
    out = []
    for i, r in enumerate(rects):
        if r['w'] <= 0 or r['h'] <= 0:
            continue
        contained = False
        for j, o in enumerate(rects):
            if i != j and r['x'] >= o['x'] and r['y'] >= o['y'] and r['x']+r['w'] <= o['x']+o['w'] and r['y']+r['h'] <= o['y']+o['h']:
                contained = True
                break
        if not contained:
            out.append(r)
    return out

def split_rect(rect, pw, ph, kerf):
    rects = []
    rw = rect['w'] - pw - kerf
    bh = rect['h'] - ph - kerf
    if rw > 0:
        rects.append({'x': rect['x'] + pw + kerf, 'y': rect['y'], 'w': rw, 'h': ph})
    if bh > 0:
        rects.append({'x': rect['x'], 'y': rect['y'] + ph + kerf, 'w': rect['w'], 'h': bh})
    return prune_rects(rects)

def try_place(free_rects, part, rotate_allowed):
    variants = [(part['width_mm'], part['height_mm'], False)]
    if rotate_allowed and abs(part['width_mm'] - part['height_mm']) > 1e-9:
        variants.append((part['height_mm'], part['width_mm'], True))
    best = None
    for idx, rect in enumerate(free_rects):
        for w, h, rotated in variants:
            if w <= rect['w'] and h <= rect['h']:
                score = (rect['w'] * rect['h'] - w * h, min(rect['w'] - w, rect['h'] - h))
                if best is None or score < best['score']:
                    best = {'score': score, 'idx': idx, 'w': w, 'h': h, 'rotated': rotated}
    return best

def build_groups(parts, mix):
    groups = {}
    for p in parts:
        key = (p['color'], float(p['thickness_mm'])) if mix else (p['product_code'], p['color'], float(p['thickness_mm']))
        groups.setdefault(key, []).append(p)
    return groups

def optimize_parts(parts, board_width, board_height, kerf, margin, rotate_allowed, mix):
    all_sheets = []
    for key, group_parts in build_groups(parts, mix).items():
        group_name = f'색상:{key[0]} / 두께:{key[1]}' if mix else f'제품:{key[0]} / 색상:{key[1]} / 두께:{key[2]}'
        usable_w = board_width - margin * 2
        usable_h = board_height - margin * 2
        sheets = []
        for part in expand_parts(group_parts):
            placed = False
            for sheet in sheets:
                best = try_place(sheet['free_rects'], part, rotate_allowed)
                if best:
                    rect = sheet['free_rects'].pop(best['idx'])
                    sheet['placements'].append({'x_mm': round(rect['x']+margin,1), 'y_mm': round(rect['y']+margin,1), 'width_mm': round(best['w'],1), 'height_mm': round(best['h'],1), 'part_code': part['part_code']})
                    sheet['free_rects'].extend(split_rect(rect, best['w'], best['h'], kerf))
                    sheet['free_rects'] = prune_rects(sheet['free_rects'])
                    placed = True
                    break
            if not placed:
                sheet = {'sheet_no': len(sheets)+1, 'group_name': group_name, 'placements': [], 'free_rects': [{'x':0.0,'y':0.0,'w':usable_w,'h':usable_h}]}
                best = try_place(sheet['free_rects'], part, rotate_allowed)
                if best:
                    rect = sheet['free_rects'].pop(best['idx'])
                    sheet['placements'].append({'x_mm': round(rect['x']+margin,1), 'y_mm': round(rect['y']+margin,1), 'width_mm': round(best['w'],1), 'height_mm': round(best['h'],1), 'part_code': part['part_code']})
                    sheet['free_rects'].extend(split_rect(rect, best['w'], best['h'], kerf))
                    sheet['free_rects'] = prune_rects(sheet['free_rects'])
                    sheets.append(sheet)
        start = len(all_sheets)
        for i, s in enumerate(sheets, start=1):
            s['sheet_no'] = start + i
            all_sheets.append(s)
    total_part_area = sum(p['width_mm'] * p['height_mm'] for s in all_sheets for p in s['placements'])
    board_area = board_width * board_height
    total_board_area = len(all_sheets) * board_area if all_sheets else 0
    waste_area = max(0.0, total_board_area - total_part_area)
    yield_rate = round((total_part_area / total_board_area) * 100, 2) if total_board_area else 0.0
    return {'board_width_mm': round(board_width,1), 'board_height_mm': round(board_height,1), 'used_boards': len(all_sheets), 'waste_area': round(waste_area,1), 'yield_rate': yield_rate, 'sheets': all_sheets}

def analyze_alternatives(parts, bw, bh, kerf, margin, rotate_allowed, mix):
    rows = []
    for name, xw, xh, mg in [('현재 조건', bw, bh, margin), ('여유치 5.0', bw, bh, 5.0), ('여유치 8.0', bw, bh, 8.0), ('여유치 10.0', bw, bh, 10.0), ('원장 4x6 / 여유치 10.0', 1830.0, 1220.0, 10.0)]:
        r = optimize_parts(parts, xw, xh, kerf, mg, rotate_allowed, mix)
        rows.append({'시나리오': name, '원장 가로': xw, '원장 세로': xh, '여유치': mg, '사용 원장 수': r['used_boards'], '수율(%)': r['yield_rate'], '자투리 면적': r['waste_area']})
    df = pd.DataFrame(rows)
    best = df.sort_values(['수율(%)', '사용 원장 수'], ascending=[False, True]).iloc[0]
    summary = f"분석 결과, '{best['시나리오']}' 조건이 가장 높은 수율을 보였습니다. 예상 수율은 {best['수율(%)']}%이며, 사용 원장 수는 {int(best['사용 원장 수'])}장입니다."
    return df, summary

def make_svg(sheet, board_width_mm, board_height_mm, kerf=0.0):
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale)
    svg_height = int(board_height_mm * scale)
    parts = []
    for p in sheet['placements']:
        x, y = p['x_mm'] * scale, p['y_mm'] * scale
        w, h = p['width_mm'] * scale, p['height_mm'] * scale
        label = f"{p['part_code']} ({p['width_mm']}x{p['height_mm']})"
        parts.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect><text x="{x+4}" y="{y+16}" font-size="12" fill="#111">{label}</text></g>')
    return f'<div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;"><svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>{"".join(parts)}</svg></div>'

st.title('목재 재단 프로그램 Fast v8')
st.caption('수율 개선 분석 복구 + BOM 전체 보기')

uploaded_file = st.file_uploader('BOM 엑셀 업로드', type=['xlsx', 'xls'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    items, errors = load_bom_from_dataframe(df)
    st.session_state['bom_items'] = items
    st.session_state['bom_errors'] = errors
    st.success(f'업로드 완료: {len(items)}행, 오류 {len(errors)}건')

bom_items = st.session_state.get('bom_items', [])
if bom_items:
    query = st.text_input('제품코드 입력', value=st.session_state.get('product_query', ''))
    st.session_state['product_query'] = query
    products = sorted({x['product_code'] for x in bom_items if x['product_code']})
    matched_products = [p for p in products if query.strip().upper() in p.upper()] if query.strip() else products[:100]
    selected_products = st.multiselect('조회 결과 / 재단 대상 제품 선택', matched_products, default=matched_products[:1] if matched_products else [])
    if selected_products:
        target_items = [x for x in bom_items if x['product_code'] in selected_products and x['is_cutting_target'] and x.get('width_mm') and x.get('height_mm') and x.get('thickness_mm')]
        left, right = st.columns([1.35, 1.15])
        with left:
            mix = st.checkbox('같은 색상 + 같은 두께 혼합 재단', value=False)
            bulk_actual_cut_qty = st.number_input('실제 재단 수량 일괄 입력', min_value=1, value=1, step=1)
            rows = []
            for item in target_items:
                row = dict(item)
                row['selected'] = True
                row['actual_cut_qty'] = int(bulk_actual_cut_qty)
                row['qty'] = int(row['bom_qty']) * int(row['actual_cut_qty'])
                rows.append(row)
            edited_df = st.data_editor(pd.DataFrame(rows), use_container_width=True, height=540, disabled=['row_no','product_code','part_code','part_name','color','spec_raw','width_mm','height_mm','thickness_mm','material_name','process_name','is_cutting_target','qty'])
            if not edited_df.empty:
                edited_df['bom_qty'] = edited_df['bom_qty'].fillna(1).astype(int).clip(lower=1)
                edited_df['actual_cut_qty'] = edited_df['actual_cut_qty'].fillna(1).astype(int).clip(lower=1)
                edited_df['qty'] = edited_df['bom_qty'] * edited_df['actual_cut_qty']
                a, b = st.columns(2)
                a.metric('선택 품목 수', int((edited_df['selected'] == True).sum()))
                b.metric('총 재단 수량 합계', int(edited_df.loc[edited_df['selected'] == True, 'qty'].sum()))
        with right:
            preset = st.selectbox('원장 규격 선택', list(BOARD_PRESETS.keys()), index=0)
            preset_w, preset_h = (2440.0, 1220.0) if BOARD_PRESETS[preset] is None else BOARD_PRESETS[preset]
            board_width = st.number_input('원장 가로(mm)', min_value=100.0, value=float(preset_w), step=1.0, format='%.1f')
            board_height = st.number_input('원장 세로(mm)', min_value=100.0, value=float(preset_h), step=1.0, format='%.1f')
            kerf = st.number_input('톱날폭(mm)', min_value=0.0, value=4.8, step=0.1, format='%.1f')
            margin = st.number_input('여유치(mm)', min_value=0.0, value=10.0, step=0.1, format='%.1f')
            board_batch = st.number_input('한 번에 재단할 원장 매수', min_value=1, value=1, step=1)
            cut_count = st.number_input('재단 매수', min_value=1, value=1, step=1)
            rotate_allowed = st.checkbox('회전 허용', value=True)
            if st.button('최적화 실행', type='primary', use_container_width=True):
                optimized_items = edited_df.to_dict('records')
                optimized_items = [x for x in optimized_items if bool(x.get('selected'))]
                for item in optimized_items:
                    item['qty'] = int(item['bom_qty']) * int(item['actual_cut_qty'])
                optimized_items = apply_cut_count(optimized_items, int(cut_count))
                result = optimize_parts(optimized_items, float(board_width), float(board_height), float(kerf), float(margin), rotate_allowed, mix)
                result['kerf_mm'] = float(kerf)
                result['total_selected_qty'] = int(sum(item['qty'] for item in optimized_items))
                st.session_state['opt_result'] = result
                analysis_df, analysis_summary = analyze_alternatives(optimized_items, float(board_width), float(board_height), float(kerf), float(margin), rotate_allowed, mix)
                st.session_state['analysis_df'] = analysis_df
                st.session_state['analysis_summary'] = analysis_summary
            opt_result = st.session_state.get('opt_result')
            if opt_result:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric('사용 원장 수', opt_result['used_boards'])
                c2.metric('수율', f"{opt_result['yield_rate']}%")
                c3.metric('자투리 면적', f"{opt_result['waste_area']:,}")
                c4.metric('총 재단 수량 합계', opt_result.get('total_selected_qty', 0))
                if opt_result['sheets']:
                    labels = [f"Sheet {s['sheet_no']} | {s.get('group_name', '')}" for s in opt_result['sheets']]
                    selected_sheet_label = st.selectbox('시트 선택', labels)
                    selected_sheet_no = int(selected_sheet_label.split('|')[0].replace('Sheet', '').strip())
                    selected_sheet = next(s for s in opt_result['sheets'] if s['sheet_no'] == selected_sheet_no)
                    components.html(make_svg(selected_sheet, opt_result['board_width_mm'], opt_result['board_height_mm'], opt_result.get('kerf_mm', 0.0)), height=700, scrolling=True)
                analysis_df = st.session_state.get('analysis_df')
                analysis_summary = st.session_state.get('analysis_summary', '')
                if analysis_df is not None:
                    st.subheader('수율 개선 분석')
                    if analysis_summary:
                        st.info(analysis_summary)
                    st.dataframe(analysis_df, use_container_width=True, height=220)
else:
    st.info('BOM 엑셀 파일을 업로드하면 제품 조회와 최적화를 사용할 수 있습니다.')
