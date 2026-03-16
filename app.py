import math
import re
import traceback
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="목재 재단 프로그램 Fast v7", layout="wide")

BOARD_PRESETS = {
    "기본 4x8 (1220 x 2440)": (2440.0, 1220.0),
    "4x6 (1220 x 1830)": (1830.0, 1220.0),
    "맞춤 입력": None,
}

def parse_spec(spec_raw: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    if spec_raw is None or (isinstance(spec_raw, float) and math.isnan(spec_raw)):
        return None, None, None
    text = str(spec_raw).strip()
    match = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*$", text)
    if not match:
        return None, None, None
    return float(match.group(1)), float(match.group(2)), float(match.group(3))

def normalize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    return value

def to_int(value: Any, default: int = 0) -> int:
    value = normalize_value(value)
    if value in (None, ""):
        return default
    try:
        return int(float(value))
    except Exception:
        return default

def to_float(value: Any, default: float = 0.0) -> float:
    value = normalize_value(value)
    if value in (None, ""):
        return default
    try:
        return float(value)
    except Exception:
        return default

def is_cutting_target(row: Dict[str, Any], width: Optional[float], height: Optional[float], thickness: Optional[float]) -> bool:
    material = str(normalize_value(row.get("재질")) or "").strip().upper()
    image_flag = str(normalize_value(row.get("대표이미지")) or "").strip().upper()
    qty = to_int(row.get("정소요량"), 0) or to_int(row.get("실소요량"), 0)
    if width is None or height is None or thickness is None:
        return False
    if qty <= 0 or image_flag == "Y":
        return False
    if any(word in material for word in ["BOX", "포장", "철물", "경첩"]):
        return False
    return True

def load_bom_from_dataframe(df: pd.DataFrame):
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    items = []
    errors = []
    for idx, row in df.iterrows():
        raw = row.to_dict()
        width, height, thickness = parse_spec(raw.get("규격"))
        product_code = str(raw.get("품목코드") or "").strip()
        part_code = str(raw.get("부품코드") or "").strip()
        part_name = str(raw.get("품목명") or "").strip()
        bom_qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        item = {
            "selected": True,
            "row_no": int(idx) + 2,
            "product_code": product_code,
            "part_code": part_code,
            "part_name": part_name,
            "color": str(raw.get("색상") or "").strip(),
            "bom_qty": max(1, bom_qty),
            "actual_cut_qty": 1,
            "qty": max(1, bom_qty),
            "spec_raw": str(raw.get("규격") or "").strip(),
            "width_mm": width,
            "height_mm": height,
            "thickness_mm": thickness,
            "material_name": str(raw.get("재질") or "").strip(),
            "process_name": str(raw.get("소요공정") or "").strip(),
            "is_cutting_target": is_cutting_target(raw, width, height, thickness),
        }
        if not product_code:
            errors.append({"row": int(idx) + 2, "field": "품목코드", "message": "제품코드 누락"})
        if not part_code:
            errors.append({"row": int(idx) + 2, "field": "부품코드", "message": "부품코드 누락"})
        if item["spec_raw"] and width is None:
            errors.append({"row": int(idx) + 2, "field": "규격", "message": f"규격 파싱 실패: {item['spec_raw']}"})
        items.append(item)
    return items, errors

def apply_cut_count(parts: List[Dict[str, Any]], cut_count: int):
    adjusted = []
    for p in parts:
        row = dict(p)
        row["qty"] = max(1, to_int(row.get("qty"), 1)) * cut_count
        adjusted.append(row)
    return adjusted

def expand_parts(parts):
    expanded = []
    for p in parts:
        repeat = max(1, to_int(p.get("qty"), 1))
        for _ in range(repeat):
            expanded.append({
                "product_code": p["product_code"],
                "part_code": p["part_code"],
                "part_name": p["part_name"],
                "color": p["color"],
                "width_mm": float(p["width_mm"]),
                "height_mm": float(p["height_mm"]),
                "thickness_mm": float(p["thickness_mm"]),
                "material_name": p["material_name"],
            })
    expanded.sort(key=lambda x: x["width_mm"] * x["height_mm"], reverse=True)
    return expanded

def prune_rects(rects):
    kept = []
    for i, r in enumerate(rects):
        if r["w"] <= 0 or r["h"] <= 0:
            continue
        contained = False
        for j, other in enumerate(rects):
            if i == j:
                continue
            if r["x"] >= other["x"] - 1e-9 and r["y"] >= other["y"] - 1e-9 and r["x"] + r["w"] <= other["x"] + other["w"] + 1e-9 and r["y"] + r["h"] <= other["y"] + other["h"] + 1e-9:
                contained = True
                break
        if not contained:
            kept.append(r)
    unique = []
    seen = set()
    for r in kept:
        key = (round(r["x"], 4), round(r["y"], 4), round(r["w"], 4), round(r["h"], 4))
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique

def split_rect_guillotine(rect, placed_w, placed_h, kerf):
    right_w = rect["w"] - placed_w - kerf
    bottom_h = rect["h"] - placed_h - kerf
    new_rects = []
    if right_w > 0:
        new_rects.append({"x": rect["x"] + placed_w + kerf, "y": rect["y"], "w": right_w, "h": placed_h})
    if bottom_h > 0:
        new_rects.append({"x": rect["x"], "y": rect["y"] + placed_h + kerf, "w": rect["w"], "h": bottom_h})
    return prune_rects(new_rects)

def try_place_part(free_rects, part, rotate_allowed):
    variants = [(part["width_mm"], part["height_mm"], False)]
    if rotate_allowed and abs(part["width_mm"] - part["height_mm"]) > 1e-9:
        variants.append((part["height_mm"], part["width_mm"], True))
    best = None
    for idx, rect in enumerate(free_rects):
        for w, h, rotated in variants:
            if w <= rect["w"] + 1e-9 and h <= rect["h"] + 1e-9:
                waste = rect["w"] * rect["h"] - w * h
                short_side = min(rect["w"] - w, rect["h"] - h)
                score = (waste, short_side)
                if best is None or score < best["score"]:
                    best = {"score": score, "rect_index": idx, "w": w, "h": h, "rotated": rotated}
    return best

def optimize_group(parts, board_width, board_height, kerf, margin, rotate_allowed, group_name):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    if usable_w <= 0 or usable_h <= 0:
        raise ValueError("원장 크기보다 여유치가 큽니다.")
    sheets = []
    unplaced = []
    for part in expand_parts(parts):
        placed = False
        for sheet in sheets:
            best = try_place_part(sheet["free_rects"], part, rotate_allowed)
            if best is None:
                continue
            rect = sheet["free_rects"].pop(best["rect_index"])
            placement = {"x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1), "rotated": best["rotated"], "part_code": part["part_code"], "part_name": part["part_name"], "product_code": part["product_code"], "color": part["color"], "thickness_mm": round(part["thickness_mm"], 1), "material_name": part["material_name"]}
            sheet["placements"].append(placement)
            sheet["free_rects"].extend(split_rect_guillotine(rect, best["w"], best["h"], kerf))
            sheet["free_rects"] = prune_rects(sheet["free_rects"])
            placed = True
            break
        if not placed:
            new_sheet = {"sheet_no": len(sheets) + 1, "group_name": group_name, "placements": [], "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}]}
            best = try_place_part(new_sheet["free_rects"], part, rotate_allowed)
            if best is None:
                unplaced.append(part)
            else:
                rect = new_sheet["free_rects"].pop(best["rect_index"])
                placement = {"x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1), "rotated": best["rotated"], "part_code": part["part_code"], "part_name": part["part_name"], "product_code": part["product_code"], "color": part["color"], "thickness_mm": round(part["thickness_mm"], 1), "material_name": part["material_name"]}
                new_sheet["placements"].append(placement)
                new_sheet["free_rects"].extend(split_rect_guillotine(rect, best["w"], best["h"], kerf))
                new_sheet["free_rects"] = prune_rects(new_sheet["free_rects"])
                sheets.append(new_sheet)
    return sheets, unplaced

def build_groups(parts, mix_same_color_thickness):
    groups = {}
    if mix_same_color_thickness:
        for p in parts:
            key = (p["color"], float(p["thickness_mm"]))
            groups.setdefault(key, []).append(p)
    else:
        for p in parts:
            key = (p["product_code"], p["color"], float(p["thickness_mm"]))
            groups.setdefault(key, []).append(p)
    return groups

def optimize_parts(parts, board_width, board_height, kerf, margin, rotate_allowed, mix_same_color_thickness):
    groups = build_groups(parts, mix_same_color_thickness)
    all_sheets = []
    all_unplaced = []
    for key, group_parts in groups.items():
        group_name = f"색상:{key[0]} / 두께:{key[1]}" if mix_same_color_thickness else f"제품:{key[0]} / 색상:{key[1]} / 두께:{key[2]}"
        sheets, unplaced = optimize_group(group_parts, board_width, board_height, kerf, margin, rotate_allowed, group_name)
        start_no = len(all_sheets)
        for i, s in enumerate(sheets, start=1):
            s["sheet_no"] = start_no + i
            all_sheets.append(s)
        all_unplaced.extend(unplaced)
    total_part_area = sum(p["width_mm"] * p["height_mm"] for s in all_sheets for p in s["placements"])
    board_area = board_width * board_height
    used_boards = len(all_sheets)
    total_board_area = used_boards * board_area if used_boards else 0.0
    waste_area = max(0.0, total_board_area - total_part_area)
    yield_rate = round((total_part_area / total_board_area) * 100, 2) if total_board_area else 0.0
    return {"board_width_mm": round(board_width, 1), "board_height_mm": round(board_height, 1), "used_boards": used_boards, "total_part_area": round(total_part_area, 1), "waste_area": round(waste_area, 1), "yield_rate": yield_rate, "unplaced_count": len(all_unplaced), "unplaced_parts": all_unplaced, "sheets": all_sheets}

def make_svg(sheet, board_width_mm, board_height_mm, kerf=0.0):
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale)
    svg_height = int(board_height_mm * scale)
    parts_svg = []
    kerf_svg = []
    for p in sheet["placements"]:
        x = p["x_mm"] * scale
        y = p["y_mm"] * scale
        w = p["width_mm"] * scale
        h = p["height_mm"] * scale
        label = f'{p["part_code"]} ({p["width_mm"]}x{p["height_mm"]})'
        parts_svg.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect><text x="{x + 4}" y="{y + 16}" font-size="12" fill="#111">{label}</text></g>')
        if kerf > 0:
            k = kerf * scale
            kerf_svg.append(f'<rect x="{x + w}" y="{y}" width="{k}" height="{h}" fill="#fca5a5" fill-opacity="0.35"></rect>')
            kerf_svg.append(f'<rect x="{x}" y="{y + h}" width="{w}" height="{k}" fill="#fca5a5" fill-opacity="0.35"></rect>')
    return f'<div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;"><svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>{"".join(kerf_svg)}{"".join(parts_svg)}</svg></div>'

st.title("목재 재단 프로그램 Fast v7")
st.caption("총 재단 수량 자동 계산 + 일괄 실제 재단 수량 입력")

uploaded_file = st.file_uploader("BOM 엑셀 업로드", type=["xlsx", "xls"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        items, errors = load_bom_from_dataframe(df)
        st.session_state["bom_items"] = items
        st.session_state["bom_errors"] = errors
        st.success(f"업로드 완료: {len(items)}행, 오류 {len(errors)}건")
    except Exception as exc:
        st.error(f"엑셀 파일 읽기 오류: {exc}")

bom_items = st.session_state.get("bom_items", [])
if bom_items:
    products = sorted({x["product_code"] for x in bom_items if x["product_code"]})
    query = st.text_input("제품코드 입력", value=st.session_state.get("product_query", ""))
    st.session_state["product_query"] = query
    matched_products = [p for p in products if query.strip().upper() in p.upper()] if query.strip() else products[:50]
    selected_products = st.multiselect("조회 결과 / 재단 대상 제품 선택", matched_products, default=matched_products[:1] if matched_products else [])
    if selected_products:
        target_items = [x for x in bom_items if x["product_code"] in selected_products and x["is_cutting_target"] and x.get("width_mm") and x.get("height_mm") and x.get("thickness_mm")]
        left, right = st.columns([1.25, 1.2])
        with left:
            mix_same_color_thickness = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=False)
            bulk_actual_cut_qty = st.number_input("실제 재단 수량 일괄 입력", min_value=1, value=1, step=1)
            editable_rows = []
            for item in target_items:
                row = dict(item)
                row["selected"] = True
                row["bom_qty"] = max(1, to_int(row.get("bom_qty"), 1))
                row["actual_cut_qty"] = int(bulk_actual_cut_qty)
                row["qty"] = row["bom_qty"] * row["actual_cut_qty"]
                editable_rows.append(row)
            temp_df = pd.DataFrame(editable_rows)
            edited_df = st.data_editor(
                temp_df,
                use_container_width=True,
                height=520,
                disabled=["row_no","product_code","part_code","part_name","color","spec_raw","width_mm","height_mm","thickness_mm","material_name","process_name","is_cutting_target","qty"],
                column_config={
                    "selected": st.column_config.CheckboxColumn("선택", default=True),
                    "bom_qty": st.column_config.NumberColumn("BOM 재단 수량", min_value=1, step=1, format="%d"),
                    "actual_cut_qty": st.column_config.NumberColumn("실제 재단 수량", min_value=1, step=1, format="%d"),
                    "qty": st.column_config.NumberColumn("총 재단 수량", format="%d"),
                },
            )
            if not edited_df.empty:
                edited_df["bom_qty"] = edited_df["bom_qty"].fillna(1).astype(int).clip(lower=1)
                edited_df["actual_cut_qty"] = edited_df["actual_cut_qty"].fillna(1).astype(int).clip(lower=1)
                edited_df["qty"] = edited_df["bom_qty"] * edited_df["actual_cut_qty"]
                st.metric("총 재단 수량 합계", int(edited_df.loc[edited_df["selected"] == True, "qty"].sum()))
        with right:
            preset = st.selectbox("원장 규격 선택", list(BOARD_PRESETS.keys()), index=0)
            preset_w, preset_h = (2440.0, 1220.0) if BOARD_PRESETS[preset] is None else BOARD_PRESETS[preset]
            board_width = st.number_input("원장 가로(mm)", min_value=100.0, value=float(preset_w), step=1.0, format="%.1f")
            board_height = st.number_input("원장 세로(mm)", min_value=100.0, value=float(preset_h), step=1.0, format="%.1f")
            kerf = st.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
            margin = st.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
            cut_count = st.number_input("재단 매수", min_value=1, value=1, step=1)
            if st.button("최적화 실행", type="primary", use_container_width=True):
                optimized_items = edited_df.to_dict("records")
                optimized_items = [x for x in optimized_items if bool(x.get("selected"))]
                for item in optimized_items:
                    item["qty"] = int(item["bom_qty"]) * int(item["actual_cut_qty"])
                optimized_items = apply_cut_count(optimized_items, int(cut_count))
                result = optimize_parts(optimized_items, float(board_width), float(board_height), float(kerf), float(margin), True, mix_same_color_thickness)
                result["kerf_mm"] = float(kerf)
                st.session_state["opt_result"] = result
            opt_result = st.session_state.get("opt_result")
            if opt_result and opt_result["sheets"]:
                labels = [f"Sheet {s['sheet_no']} | {s.get('group_name', '')}" for s in opt_result["sheets"]]
                selected_sheet_label = st.selectbox("시트 선택", labels)
                selected_sheet_no = int(selected_sheet_label.split("|")[0].replace("Sheet", "").strip())
                selected_sheet = next(s for s in opt_result["sheets"] if s["sheet_no"] == selected_sheet_no)
                components.html(make_svg(selected_sheet, opt_result["board_width_mm"], opt_result["board_height_mm"], opt_result.get("kerf_mm", 0.0)), height=700, scrolling=True)
else:
    st.info("BOM 엑셀 파일을 업로드하면 제품 조회와 최적화를 사용할 수 있습니다.")
