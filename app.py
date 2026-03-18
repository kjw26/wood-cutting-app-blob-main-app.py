
import math
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="목재 재단 프로그램 Fast v9", layout="wide")

BOARD_PRESETS = {
    "기본 4x8 (1220 x 2440)": (2440.0, 1220.0),
    "4x6 (1220 x 1830)": (1830.0, 1220.0),
    "맞춤 입력": None,
}

def parse_spec(spec_raw: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    if spec_raw is None or (isinstance(spec_raw, float) and math.isnan(spec_raw)):
        return None, None, None
    text = str(spec_raw).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*$", text)
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))

def to_int(v: Any, default=0) -> int:
    try:
        if pd.isna(v):
            return default
        return int(float(v))
    except Exception:
        return default

def to_float(v: Any, default=0.0) -> float:
    try:
        if pd.isna(v):
            return default
        return float(v)
    except Exception:
        return default

def is_cutting_target(row: Dict[str, Any], w, h, t) -> bool:
    material = str(row.get("재질") or "").upper()
    image_flag = str(row.get("대표이미지") or "").upper()
    qty = to_int(row.get("정소요량"), 0) or to_int(row.get("실소요량"), 0)
    if w is None or h is None or t is None or qty <= 0 or image_flag == "Y":
        return False
    if any(x in material for x in ["BOX", "포장", "철물", "경첩"]):
        return False
    return True

def load_bom(df: pd.DataFrame):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    items, errors = [], []
    for idx, row in df.iterrows():
        raw = row.to_dict()
        w, h, t = parse_spec(raw.get("규격"))
        product_code = str(raw.get("품목코드") or "").strip()
        part_code = str(raw.get("부품코드") or "").strip()
        bom_qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        item = {
            "selected": True,
            "row_no": int(idx) + 2,
            "product_code": product_code,
            "part_code": part_code,
            "part_name": str(raw.get("품목명") or "").strip(),
            "color": str(raw.get("색상") or "").strip(),
            "bom_qty": max(1, bom_qty),
            "actual_cut_qty": 1,
            "qty": max(1, bom_qty),
            "spec_raw": str(raw.get("규격") or "").strip(),
            "width_mm": w,
            "height_mm": h,
            "thickness_mm": t,
            "material_name": str(raw.get("재질") or "").strip(),
            "is_cutting_target": is_cutting_target(raw, w, h, t),
        }
        if not product_code:
            errors.append({"row": int(idx)+2, "field": "품목코드", "message": "제품코드 누락"})
        if not part_code:
            errors.append({"row": int(idx)+2, "field": "부품코드", "message": "부품코드 누락"})
        items.append(item)
    return pd.DataFrame(items), pd.DataFrame(errors)

def parse_plan_workbook(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    rows = []
    for sheet in xls.sheet_names:
        if "주차" not in str(sheet):
            continue
        raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if raw.shape[0] < 4 or raw.shape[1] < 8:
            continue
        header_row = 2
        current_product = None
        for r in range(3, raw.shape[0]):
            product = raw.iat[r,0]
            color = raw.iat[r,2] if raw.shape[1] > 2 else None
            if pd.notna(product) and str(product).strip():
                current_product = str(product).strip()
            if not current_product:
                continue
            for c in range(7, min(raw.shape[1], 13)):
                date_val = raw.iat[header_row, c]
                qty = raw.iat[r, c]
                if pd.isna(date_val) or pd.isna(qty):
                    continue
                q = to_int(qty, 0)
                if q <= 0:
                    continue
                try:
                    d = pd.to_datetime(date_val).date()
                except Exception:
                    continue
                rows.append({
                    "sheet": str(sheet),
                    "date": str(d),
                    "product_code": current_product,
                    "color": "" if pd.isna(color) else str(color).strip(),
                    "plan_qty": q,
                })
    return pd.DataFrame(rows)

def expand_parts(parts):
    expanded = []
    for p in parts:
        for _ in range(max(1, to_int(p.get("qty"), 1))):
            expanded.append({
                "product_code": p["product_code"],
                "part_code": p["part_code"],
                "part_name": p["part_name"],
                "color": p["color"],
                "width_mm": float(p["width_mm"]),
                "height_mm": float(p["height_mm"]),
                "thickness_mm": float(p["thickness_mm"]),
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
            if i != j and r["x"] >= other["x"] and r["y"] >= other["y"] and r["x"]+r["w"] <= other["x"]+other["w"] and r["y"]+r["h"] <= other["y"]+other["h"]:
                contained = True
                break
        if not contained:
            kept.append(r)
    return kept

def split_rect(rect, placed_w, placed_h, kerf):
    right_w = rect["w"] - placed_w - kerf
    bottom_h = rect["h"] - placed_h - kerf
    out = []
    if right_w > 0:
        out.append({"x": rect["x"] + placed_w + kerf, "y": rect["y"], "w": right_w, "h": placed_h})
    if bottom_h > 0:
        out.append({"x": rect["x"], "y": rect["y"] + placed_h + kerf, "w": rect["w"], "h": bottom_h})
    return prune_rects(out)

def try_place_part(free_rects, part, rotate_allowed):
    variants = [(part["width_mm"], part["height_mm"], False)]
    if rotate_allowed and abs(part["width_mm"] - part["height_mm"]) > 1e-9:
        variants.append((part["height_mm"], part["width_mm"], True))
    best = None
    for idx, rect in enumerate(free_rects):
        for w, h, rotated in variants:
            if w <= rect["w"] + 1e-9 and h <= rect["h"] + 1e-9:
                score = (rect["w"] * rect["h"] - w * h, min(rect["w"] - w, rect["h"] - h))
                if best is None or score < best["score"]:
                    best = {"score": score, "rect_index": idx, "w": w, "h": h, "rotated": rotated}
    return best

def optimize_parts(parts, board_width, board_height, kerf, margin, rotate_allowed, mix_same_color_thickness):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    if usable_w <= 0 or usable_h <= 0:
        raise ValueError("원장 크기보다 여유치가 큽니다.")
    groups = {}
    for p in parts:
        key = (p["color"], float(p["thickness_mm"])) if mix_same_color_thickness else (p["product_code"], p["color"], float(p["thickness_mm"]))
        groups.setdefault(key, []).append(p)

    all_sheets = []
    for key, group_parts in groups.items():
        group_name = f"색상:{key[0]} / 두께:{key[1]}" if mix_same_color_thickness else f"제품:{key[0]} / 색상:{key[1]} / 두께:{key[2]}"
        sheets = []
        for part in expand_parts(group_parts):
            placed = False
            for sheet in sheets:
                best = try_place_part(sheet["free_rects"], part, rotate_allowed)
                if best is None:
                    continue
                rect = sheet["free_rects"].pop(best["rect_index"])
                sheet["placements"].append({
                    "x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1),
                    "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1),
                    "part_code": part["part_code"], "product_code": part["product_code"], "color": part["color"]
                })
                sheet["free_rects"].extend(split_rect(rect, best["w"], best["h"], kerf))
                sheet["free_rects"] = prune_rects(sheet["free_rects"])
                placed = True
                break
            if not placed:
                sheet = {"sheet_no": len(sheets)+1, "group_name": group_name, "placements": [], "free_rects": [{"x":0.0,"y":0.0,"w":usable_w,"h":usable_h}]}
                best = try_place_part(sheet["free_rects"], part, rotate_allowed)
                if best is None:
                    continue
                rect = sheet["free_rects"].pop(best["rect_index"])
                sheet["placements"].append({
                    "x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1),
                    "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1),
                    "part_code": part["part_code"], "product_code": part["product_code"], "color": part["color"]
                })
                sheet["free_rects"].extend(split_rect(rect, best["w"], best["h"], kerf))
                sheet["free_rects"] = prune_rects(sheet["free_rects"])
                sheets.append(sheet)
        all_sheets.extend(sheets)

    total_area = sum(p["width_mm"] * p["height_mm"] for s in all_sheets for p in s["placements"])
    board_area = board_width * board_height
    used = len(all_sheets)
    yield_rate = round((total_area / (used * board_area)) * 100, 2) if used else 0.0
    return {"board_width_mm": board_width, "board_height_mm": board_height, "used_boards": used, "yield_rate": yield_rate, "waste_area": max(0.0, used*board_area-total_area), "sheets": all_sheets}

def analyze_alternatives(parts, bw, bh, kerf, margin, rotate_allowed, mix_same_color_thickness):
    scenarios = [("현재 조건", bw, bh, margin), ("여유치 5.0", bw, bh, 5.0), ("여유치 8.0", bw, bh, 8.0), ("여유치 10.0", bw, bh, 10.0), ("원장 4x6 / 여유치 10.0", 1830.0, 1220.0, 10.0)]
    rows = []
    for name, sw, sh, sm in scenarios:
        r = optimize_parts(parts, sw, sh, kerf, sm, rotate_allowed, mix_same_color_thickness)
        rows.append({"시나리오": name, "원장 가로": sw, "원장 세로": sh, "여유치": sm, "사용 원장 수": r["used_boards"], "수율(%)": r["yield_rate"], "자투리 면적": r["waste_area"]})
    df = pd.DataFrame(rows)
    best = df.sort_values(["수율(%)", "사용 원장 수"], ascending=[False, True]).iloc[0]
    summary = f"분석 결과, '{best['시나리오']}' 조건이 가장 높은 수율을 보였습니다. 예상 수율은 {best['수율(%)']}%이며, 사용 원장 수는 {int(best['사용 원장 수'])}장입니다."
    return df, summary

def make_svg(sheet, board_width_mm, board_height_mm):
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale)
    svg_height = int(board_height_mm * scale)
    parts_svg = []
    for p in sheet["placements"]:
        x, y = p["x_mm"] * scale, p["y_mm"] * scale
        w, h = p["width_mm"] * scale, p["height_mm"] * scale
        label = f'{p["part_code"]} ({p["width_mm"]}x{p["height_mm"]})'
        parts_svg.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect><text x="{x+4}" y="{y+16}" font-size="12" fill="#111">{label}</text></g>')
    return f'<div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;"><svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>{"".join(parts_svg)}</svg></div>'

st.title("목재 재단 프로그램 Fast v9")
st.caption("BOM 전체 표시 + 월생산계획 자동 재단")

bom_file = st.file_uploader("BOM 엑셀 업로드", type=["xlsx", "xls"], key="bom")
plan_file = st.file_uploader("월생산계획 업로드", type=["xlsx", "xls"], key="plan")

bom_df = None
if bom_file is not None:
    raw_bom = pd.read_excel(bom_file)
    bom_df, bom_errors = load_bom(raw_bom)
    st.success(f"BOM 업로드 완료: {len(bom_df)}행")
    with st.expander("BOM 전체 데이터 보기", expanded=True):
        st.dataframe(bom_df, use_container_width=True, height=500)

if bom_df is not None and plan_file is not None:
    plan_df = parse_plan_workbook(plan_file)
    if plan_df.empty:
        st.warning("월생산계획 파일에서 날짜별 생산계획을 찾지 못했습니다. 주차 시트를 확인해 주세요.")
    else:
        st.success(f"월생산계획 해석 완료: {len(plan_df)}건")
        board_preset = st.selectbox("원장 규격 선택", list(BOARD_PRESETS.keys()), index=0)
        bw, bh = (2440.0, 1220.0) if BOARD_PRESETS[board_preset] is None else BOARD_PRESETS[board_preset]
        c1, c2, c3 = st.columns(3)
        kerf = c1.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
        margin = c2.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
        rotate_allowed = c3.checkbox("회전 허용", value=True)
        mix_same = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=True)

        daily_rows = []
        detail_results = {}
        for d, g in plan_df.groupby("date"):
            parts = []
            for _, row in g.iterrows():
                product_code = str(row["product_code"]).strip()
                plan_qty = to_int(row["plan_qty"], 0)
                matched = bom_df[(bom_df["product_code"] == product_code) & (bom_df["is_cutting_target"] == True)].copy()
                if matched.empty:
                    continue
                matched["actual_cut_qty"] = plan_qty
                matched["qty"] = matched["bom_qty"].astype(int) * plan_qty
                parts.extend(matched.to_dict("records"))
            if not parts:
                continue
            result = optimize_parts(parts, bw, bh, float(kerf), float(margin), rotate_allowed, mix_same)
            detail_results[d] = result
            for grp, gg in pd.DataFrame(parts).groupby(["color", "thickness_mm"]):
                sub = optimize_parts(gg.to_dict("records"), bw, bh, float(kerf), float(margin), rotate_allowed, True)
                daily_rows.append({
                    "date": d,
                    "color": grp[0],
                    "thickness_mm": grp[1],
                    "총 재단 수량": int(gg["qty"].sum()),
                    "사용 원장 수": sub["used_boards"],
                    "수율(%)": sub["yield_rate"],
                    "자투리 면적": sub["waste_area"],
                })

        daily_df = pd.DataFrame(daily_rows).sort_values(["date", "thickness_mm", "color"])
        st.subheader("날짜별 / 두께별 / 색상별 재단 계획")
        st.dataframe(daily_df, use_container_width=True, height=420)

        if not daily_df.empty:
            analysis_df, summary = analyze_alternatives(
                [r for _, r in bom_df[bom_df["is_cutting_target"] == True].head(200).iterrows()],
                bw, bh, float(kerf), float(margin), rotate_allowed, mix_same
            )
            st.subheader("재단 수율 개선 분석")
            st.info(summary)
            st.dataframe(analysis_df, use_container_width=True, height=220)

            selected_date = st.selectbox("날짜별 분할도 보기", sorted(detail_results.keys()))
            selected_result = detail_results[selected_date]
            if selected_result["sheets"]:
                labels = [f"Sheet {s['sheet_no']} | {s.get('group_name','')}" for s in selected_result["sheets"]]
                selected_sheet_label = st.selectbox("시트 선택", labels)
                selected_sheet_no = int(selected_sheet_label.split("|")[0].replace("Sheet", "").strip())
                selected_sheet = next(s for s in selected_result["sheets"] if s["sheet_no"] == selected_sheet_no)
                st.subheader(f"{selected_date} 분할도")
                components.html(make_svg(selected_sheet, selected_result["board_width_mm"], selected_result["board_height_mm"]), height=700, scrolling=True)
                st.metric("해당 일자 사용 원장 수", selected_result["used_boards"])
