
import io
import math
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="목재 재단 프로그램 Fast v13", layout="wide")

BOARD_PRESETS = {
    "기본 4x8 (1220 x 2440)": (2440.0, 1220.0),
    "4x6 (1220 x 1830)": (1830.0, 1220.0),
    "맞춤 입력": None,
}
RESULT_COLS = ["date", "color", "thickness_mm", "총 재단 수량", "사용 원장 수", "수율(%)", "자투리 면적"]

def normalize_text(v: Any) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()

def to_int(v: Any, default=0) -> int:
    try:
        if pd.isna(v):
            return default
        return int(float(v))
    except Exception:
        return default

def parse_spec(spec_raw: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    text = normalize_text(spec_raw)
    if not text:
        return None, None, None
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*$", text)
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))

def is_cutting_target(row: Dict[str, Any], w, h, t) -> bool:
    material = normalize_text(row.get("재질")).upper()
    image_flag = normalize_text(row.get("대표이미지")).upper()
    qty = to_int(row.get("정소요량"), 0) or to_int(row.get("실소요량"), 0)
    if w is None or h is None or t is None:
        return False
    if qty <= 0 or image_flag == "Y":
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
        product_code = normalize_text(raw.get("품목코드"))
        part_code = normalize_text(raw.get("부품코드"))
        bom_qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        item = {
            "selected": True, "row_no": int(idx) + 2, "product_code": product_code, "part_code": part_code,
            "part_name": normalize_text(raw.get("품목명")), "color": normalize_text(raw.get("색상")),
            "bom_qty": max(1, bom_qty), "actual_cut_qty": 1, "qty": max(1, bom_qty),
            "spec_raw": normalize_text(raw.get("규격")), "width_mm": w, "height_mm": h, "thickness_mm": t,
            "material_name": normalize_text(raw.get("재질")), "is_cutting_target": is_cutting_target(raw, w, h, t),
        }
        if not product_code:
            errors.append({"row": int(idx) + 2, "field": "품목코드", "message": "제품코드 누락"})
        if not part_code:
            errors.append({"row": int(idx) + 2, "field": "부품코드", "message": "부품코드 누락"})
        items.append(item)
    return pd.DataFrame(items), (pd.DataFrame(errors) if errors else pd.DataFrame(columns=["row","field","message"]))

def read_bom_from_source(uploaded_file, url_text):
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file)
    if normalize_text(url_text):
        resp = requests.get(url_text.strip(), timeout=30)
        resp.raise_for_status()
        return pd.read_excel(io.BytesIO(resp.content))
    return None

def find_header_layout(raw: pd.DataFrame):
    max_rows = min(len(raw), 20)
    max_cols = min(len(raw.columns), 30)
    for r in range(max_rows):
        for c in range(max_cols):
            val = normalize_text(raw.iat[r, c]).replace(" ", "")
            if "품목코드" in val:
                color_col = None
                for cc in range(c + 1, min(max_cols, c + 6)):
                    val2 = normalize_text(raw.iat[r, cc]).replace(" ", "")
                    if "색상" in val2:
                        color_col = cc
                        break
                if color_col is None:
                    continue
                date_cols = []
                for dc in range(color_col + 1, max_cols):
                    ts = pd.to_datetime(raw.iat[r, dc], errors="coerce")
                    if not pd.isna(ts) and ts.weekday() <= 5:
                        date_cols.append((dc, str(ts.date())))
                if len(date_cols) >= 2:
                    return {"header_row": r, "product_col": c, "color_col": color_col, "date_cols": date_cols}
    return None

def parse_plan_workbook_auto(file):
    xls = pd.ExcelFile(file)
    all_rows, parse_logs = [], []
    for sheet in xls.sheet_names:
        raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if raw.empty:
            parse_logs.append({"sheet": sheet, "status": "skip", "reason": "빈 시트"})
            continue
        layout = find_header_layout(raw)
        if layout is None:
            parse_logs.append({"sheet": sheet, "status": "skip", "reason": "품목코드/색상/날짜 헤더를 찾지 못함"})
            continue
        hdr = layout["header_row"]; product_col = layout["product_col"]; color_col = layout["color_col"]; date_cols = layout["date_cols"]
        current_product = ""; added = 0
        for r in range(hdr + 1, len(raw)):
            product_cell = normalize_text(raw.iat[r, product_col])
            if product_cell:
                current_product = product_cell
            if not current_product:
                continue
            color = normalize_text(raw.iat[r, color_col])
            for c, d in date_cols:
                qty = to_int(raw.iat[r, c], 0)
                if qty > 0:
                    all_rows.append({"sheet": str(sheet), "date": d, "product_code": current_product, "color": color, "plan_qty": qty})
                    added += 1
        parse_logs.append({"sheet": sheet, "status": "ok" if added > 0 else "empty", "header_row": hdr, "product_col": product_col, "color_col": color_col, "date_col_count": len(date_cols), "read_rows": added})
    plan_df = pd.DataFrame(all_rows)
    if plan_df.empty:
        plan_df = pd.DataFrame(columns=["sheet","date","product_code","color","plan_qty"])
    else:
        plan_df = plan_df.groupby(["date","product_code","color"], as_index=False)["plan_qty"].sum()
    log_df = pd.DataFrame(parse_logs) if parse_logs else pd.DataFrame(columns=["sheet","status","reason"])
    return plan_df, log_df

def expand_parts(parts):
    expanded = []
    for p in parts:
        for _ in range(max(1, to_int(p.get("qty"), 1))):
            expanded.append({
                "product_code": p["product_code"], "part_code": p["part_code"], "part_name": p["part_name"],
                "color": p["color"], "width_mm": float(p["width_mm"]), "height_mm": float(p["height_mm"]), "thickness_mm": float(p["thickness_mm"]),
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
            if i != j and r["x"] >= other["x"] and r["y"] >= other["y"] and r["x"] + r["w"] <= other["x"] + other["w"] and r["y"] + r["h"] <= other["y"] + other["h"]:
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
                sheet["placements"].append({"x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1), "part_code": part["part_code"], "product_code": part["product_code"], "color": part["color"], "thickness_mm": part["thickness_mm"]})
                sheet["free_rects"].extend(split_rect(rect, best["w"], best["h"], kerf)); sheet["free_rects"] = prune_rects(sheet["free_rects"]); placed = True; break
            if not placed:
                sheet = {"sheet_no": len(sheets) + 1, "group_name": group_name, "placements": [], "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}]}
                best = try_place_part(sheet["free_rects"], part, rotate_allowed)
                if best is None:
                    continue
                rect = sheet["free_rects"].pop(best["rect_index"])
                sheet["placements"].append({"x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1), "part_code": part["part_code"], "product_code": part["product_code"], "color": part["color"], "thickness_mm": part["thickness_mm"]})
                sheet["free_rects"].extend(split_rect(rect, best["w"], best["h"], kerf)); sheet["free_rects"] = prune_rects(sheet["free_rects"]); sheets.append(sheet)
        all_sheets.extend(sheets)
    total_area = sum(p["width_mm"] * p["height_mm"] for s in all_sheets for p in s["placements"])
    board_area = board_width * board_height
    used = len(all_sheets)
    yield_rate = round((total_area / (used * board_area)) * 100, 2) if used else 0.0
    return {"board_width_mm": board_width, "board_height_mm": board_height, "used_boards": used, "yield_rate": yield_rate, "waste_area": max(0.0, used * board_area - total_area), "sheets": all_sheets}

def analyze_alternatives(parts, bw, bh, kerf, margin, rotate_allowed, mix_same_color_thickness):
    scenarios = [("현재 조건", bw, bh, margin), ("여유치 5.0", bw, bh, 5.0), ("여유치 8.0", bw, bh, 8.0), ("여유치 10.0", bw, bh, 10.0), ("원장 4x6 / 여유치 10.0", 1830.0, 1220.0, 10.0)]
    rows = []
    for name, sw, sh, sm in scenarios:
        r = optimize_parts(parts, sw, sh, kerf, sm, rotate_allowed, mix_same_color_thickness)
        rows.append({"시나리오": name, "원장 가로": sw, "원장 세로": sh, "여유치": sm, "사용 원장 수": r["used_boards"], "수율(%)": r["yield_rate"], "자투리 면적": r["waste_area"]})
    df = pd.DataFrame(rows)
    best = df.sort_values(["수율(%)","사용 원장 수"], ascending=[False, True]).iloc[0]
    summary = f"분석 결과, '{best['시나리오']}' 조건이 가장 높은 수율을 보였습니다. 예상 수율은 {best['수율(%)']}%이며, 사용 원장 수는 {int(best['사용 원장 수'])}장입니다."
    return df, summary

def make_svg(sheet, board_width_mm, board_height_mm):
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale); svg_height = int(board_height_mm * scale)
    parts_svg = []
    for p in sheet["placements"]:
        x, y = p["x_mm"] * scale, p["y_mm"] * scale; w, h = p["width_mm"] * scale, p["height_mm"] * scale
        label = f'{p["part_code"]} ({p["width_mm"]}x{p["height_mm"]})'
        parts_svg.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect><text x="{x+4}" y="{y+16}" font-size="12" fill="#111">{label}</text></g>')
    return f'<div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;"><svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>{"".join(parts_svg)}</svg></div>'

def build_workorder_excel(daily_df, analysis_df, workorder_summaries, placement_tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        daily_df.to_excel(writer, sheet_name="일정요약", index=False)
        if analysis_df is not None and not analysis_df.empty:
            analysis_df.to_excel(writer, sheet_name="분석요약", index=False)
        for date_key, summary_df in workorder_summaries.items():
            summary_df.to_excel(writer, sheet_name=f"{date_key}_요약"[:31], index=False)
            place_df = placement_tables.get(date_key)
            if place_df is not None and not place_df.empty:
                place_df.to_excel(writer, sheet_name=f"{date_key}_분할"[:31], index=False)
    output.seek(0)
    return output.getvalue()

st.title("목재 재단 프로그램 Fast v13")
st.caption("생산계획 자동 인식 + 날짜별 작업지시서")

st.subheader("BOM 데이터 소스")
bom_url = st.text_input("BOM URL (GitHub raw / 직접 다운로드 링크)")
bom_file = st.file_uploader("또는 BOM 엑셀 업로드", type=["xlsx", "xls"], key="bom")
plan_file = st.file_uploader("주차별 생산계획 업로드", type=["xlsx", "xls"], key="plan")

bom_df = None
try:
    raw_bom = read_bom_from_source(bom_file, bom_url)
    if raw_bom is not None:
        bom_df, bom_errors = load_bom(raw_bom)
        st.success(f"BOM 로드 완료: {len(bom_df)}행")
        with st.expander("BOM 전체 데이터 보기", expanded=False):
            st.dataframe(bom_df, use_container_width=True, height=420)
        if not bom_errors.empty:
            with st.expander("BOM 검증 오류", expanded=False):
                st.dataframe(bom_errors, use_container_width=True, height=180)
except Exception as e:
    st.error(f"BOM 로드 오류: {e}")

if bom_df is not None and plan_file is not None:
    plan_df, parse_log_df = parse_plan_workbook_auto(plan_file)
    if plan_df.empty:
        st.warning("주차별 생산계획을 자동 해석하지 못했습니다.")
    else:
        st.success(f"주차별 생산계획 해석 완료: {len(plan_df)}건")
    with st.expander("생산계획 해석 로그", expanded=True):
        st.dataframe(parse_log_df, use_container_width=True, height=220)
    with st.expander("해석된 생산계획 보기", expanded=True):
        st.dataframe(plan_df, use_container_width=True, height=260)

    if not plan_df.empty:
        board_preset = st.selectbox("원장 규격 선택", list(BOARD_PRESETS.keys()), index=0)
        bw, bh = (2440.0, 1220.0) if BOARD_PRESETS[board_preset] is None else BOARD_PRESETS[board_preset]
        c1, c2, c3 = st.columns(3)
        kerf = c1.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
        margin = c2.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
        rotate_allowed = c3.checkbox("회전 허용", value=True)
        mix_same = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=True)

        daily_rows, analysis_rows = [], []
        detail_results, workorder_summaries, placement_tables = {}, {}, {}

        for d, g in plan_df.groupby("date"):
            parts = []
            for _, row in g.iterrows():
                product_code = normalize_text(row["product_code"]); color = normalize_text(row.get("color")); plan_qty = to_int(row["plan_qty"], 0)
                matched = bom_df[(bom_df["product_code"] == product_code) & (bom_df["is_cutting_target"] == True)].copy()
                if color:
                    matched_color = matched[matched["color"] == color].copy()
                    if not matched_color.empty:
                        matched = matched_color
                if matched.empty:
                    continue
                matched["actual_cut_qty"] = plan_qty
                matched["qty"] = matched["bom_qty"].astype(int) * plan_qty
                parts.extend(matched.to_dict("records"))
            if not parts:
                continue

            result = optimize_parts(parts, bw, bh, float(kerf), float(margin), rotate_allowed, mix_same)
            detail_results[d] = result
            part_df = pd.DataFrame(parts)

            summary_df = (
                part_df.groupby(["color","thickness_mm"], dropna=False)
                .agg(총_재단_수량=("qty","sum"), 품목수=("part_code","count"))
                .reset_index()
                .rename(columns={"color":"색상","thickness_mm":"두께(mm)","총_재단_수량":"총 재단 수량"})
            )
            summary_df["사용 원장 수"] = None
            summary_df["수율(%)"] = None
            summary_df["자투리 면적"] = None

            for idx, row2 in summary_df.iterrows():
                gg = part_df[(part_df["color"] == row2["색상"]) & (part_df["thickness_mm"] == row2["두께(mm)"])]
                sub = optimize_parts(gg.to_dict("records"), bw, bh, float(kerf), float(margin), rotate_allowed, True)
                summary_df.loc[idx, "사용 원장 수"] = sub["used_boards"]
                summary_df.loc[idx, "수율(%)"] = sub["yield_rate"]
                summary_df.loc[idx, "자투리 면적"] = sub["waste_area"]
                daily_rows.append({"date": d, "color": row2["색상"], "thickness_mm": row2["두께(mm)"], "총 재단 수량": int(row2["총 재단 수량"]), "사용 원장 수": sub["used_boards"], "수율(%)": sub["yield_rate"], "자투리 면적": sub["waste_area"]})

            analysis_df_one, summary_text = analyze_alternatives(parts, bw, bh, float(kerf), float(margin), rotate_allowed, mix_same)
            best_row = analysis_df_one.sort_values(["수율(%)","사용 원장 수"], ascending=[False, True]).iloc[0]
            analysis_rows.append({"date": d, "추천 시나리오": best_row["시나리오"], "추천 수율(%)": best_row["수율(%)"], "추천 사용 원장 수": best_row["사용 원장 수"], "분석 요약": summary_text})

            workorder_summaries[d] = summary_df
            place_rows = []
            for s in result["sheets"]:
                for p in s["placements"]:
                    place_rows.append({"date": d, "sheet_no": s["sheet_no"], "group_name": s["group_name"], "part_code": p["part_code"], "product_code": p["product_code"], "color": p["color"], "thickness_mm": p["thickness_mm"], "x_mm": p["x_mm"], "y_mm": p["y_mm"], "width_mm": p["width_mm"], "height_mm": p["height_mm"]})
            placement_tables[d] = pd.DataFrame(place_rows)

        daily_df = pd.DataFrame(daily_rows, columns=RESULT_COLS)
        analysis_df_all = pd.DataFrame(analysis_rows)

        if daily_df.empty:
            st.warning("BOM과 생산계획이 매칭되지 않았습니다. 생산계획의 품목코드/색상 값과 BOM 값을 확인해 주세요.")
        else:
            daily_df = daily_df.sort_values(["date","thickness_mm","color"])
            st.subheader("날짜별 / 두께별 / 색상별 재단 계획")
            st.dataframe(daily_df, use_container_width=True, height=360)

            st.subheader("날짜별 재단 수율 분석")
            st.dataframe(analysis_df_all, use_container_width=True, height=220)

            workorder_bytes = build_workorder_excel(daily_df, analysis_df_all, workorder_summaries, placement_tables)
            st.download_button("작업지시서 엑셀 다운로드", data=workorder_bytes, file_name="wood_cutting_workorders.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            selected_date = st.selectbox("작업지시서 / 분할도 날짜 선택", sorted(detail_results.keys()))
            selected_result = detail_results[selected_date]; selected_summary = workorder_summaries[selected_date]
            selected_analysis = analysis_df_all[analysis_df_all["date"] == selected_date]
            selected_place = placement_tables[selected_date]

            st.subheader(f"{selected_date} 작업지시서")
            m1, m2, m3 = st.columns(3)
            m1.metric("사용 원장 수", selected_result["used_boards"])
            m2.metric("수율", f"{selected_result['yield_rate']}%")
            m3.metric("자투리 면적", f"{selected_result['waste_area']:,}")
            if not selected_analysis.empty:
                st.info(selected_analysis.iloc[0]["분석 요약"])
            st.markdown("#### 색상 / 두께별 작업 요약")
            st.dataframe(selected_summary, use_container_width=True, height=220)

            if selected_result["sheets"]:
                labels = [f"Sheet {s['sheet_no']} | {s.get('group_name','')}" for s in selected_result["sheets"]]
                selected_sheet_label = st.selectbox("시트 선택", labels)
                selected_sheet_no = int(selected_sheet_label.split("|")[0].replace("Sheet","").strip())
                selected_sheet = next(s for s in selected_result["sheets"] if s["sheet_no"] == selected_sheet_no)
                st.markdown("#### 분할도")
                components.html(make_svg(selected_sheet, selected_result["board_width_mm"], selected_result["board_height_mm"]), height=700, scrolling=True)

            st.markdown("#### 상세 배치 목록")
            st.dataframe(selected_place, use_container_width=True, height=260)
