
import io
import re
import traceback

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

DEFAULT_BOM_URL = "https://raw.githubusercontent.com/kjw26/wood-cutting-app-blob-main-app.py/main/BOM_DATA.xlsx"
BOARD_PRESETS = {
    "기본 4x8 (1220 x 2440)": (2440.0, 1220.0),
    "4x6 (1220 x 1830)": (1830.0, 1220.0),
}
MAX_PLAN_ROWS = 120
MAX_SHEETS_RENDER = 20

st.set_page_config(page_title="목재 재단 프로그램 Fast v19", layout="wide")


def normalize_text(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def to_int(v, default=0):
    try:
        if pd.isna(v):
            return default
        return int(float(v))
    except Exception:
        return default


def parse_spec(spec_raw):
    text = normalize_text(spec_raw)
    if not text:
        return None, None, None
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*$", text)
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))


def parse_date_like(val):
    try:
        if isinstance(val, pd.Timestamp) and not pd.isna(val):
            return val
    except Exception:
        pass
    text = normalize_text(val)
    if not text:
        return None
    cleaned = re.sub(r"\([^)]*\)", "", text)
    cleaned = cleaned.replace(".", "/").replace("-", "/")
    cleaned = re.sub(r"[가-힣A-Za-z]", "", cleaned)
    cleaned = re.sub(r"\s+", "", cleaned).strip()
    for cand in [cleaned, text]:
        if not cand:
            continue
        if re.match(r"^\d{2}/\d{2}$", cand):
            cand = "2026/" + cand
        ts = pd.to_datetime(cand, errors="coerce")
        if not pd.isna(ts):
            return ts
    return None


def read_bom(uploaded_file, url_text):
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file)
    url = normalize_text(url_text) or DEFAULT_BOM_URL
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    return pd.read_excel(io.BytesIO(resp.content))


def is_cutting_target(row, w, h, t):
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


def load_bom(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    items = []
    for idx, row in df.iterrows():
        raw = row.to_dict()
        w, h, t = parse_spec(raw.get("규격"))
        bom_qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        items.append({
            "row_no": int(idx) + 2,
            "product_code": normalize_text(raw.get("품목코드")),
            "part_code": normalize_text(raw.get("부품코드")),
            "part_name": normalize_text(raw.get("품목명")),
            "color": normalize_text(raw.get("색상")),
            "bom_qty": max(1, bom_qty),
            "qty": max(1, bom_qty),
            "spec_raw": normalize_text(raw.get("규격")),
            "width_mm": w,
            "height_mm": h,
            "thickness_mm": t,
            "is_cutting_target": is_cutting_target(raw, w, h, t),
        })
    return pd.DataFrame(items)


def find_header_layout(raw):
    rows = min(len(raw), 45)
    cols = min(len(raw.columns), 55)
    for r in range(rows):
        row_text = [normalize_text(raw.iat[r, c]).replace(" ", "") for c in range(cols)]
        product_candidates = [i for i, v in enumerate(row_text) if "품목코드" in v]
        if not product_candidates:
            continue
        product_col = product_candidates[0]
        color_candidates = [i for i, v in enumerate(row_text) if "색상" in v]
        color_col = color_candidates[0] if color_candidates else min(product_col + 1, cols - 1)
        date_cols = []
        for rr in [r, min(r + 1, len(raw) - 1), min(r + 2, len(raw) - 1)]:
            local = []
            for c in range(cols):
                ts = parse_date_like(raw.iat[rr, c])
                if ts is not None and ts.weekday() <= 5:
                    local.append((c, str(ts.date())))
            if len(local) >= 2:
                date_cols = local
                break
        if len(date_cols) >= 2:
            return {"data_start_row": r + 1, "product_col": product_col, "color_col": color_col, "date_cols": date_cols}
    return None


def parse_plan_workbook_auto(file):
    xls = pd.ExcelFile(file)
    all_rows = []
    logs = []
    for sheet in xls.sheet_names:
        try:
            raw = pd.read_excel(file, sheet_name=sheet, header=None)
            if raw.empty:
                logs.append({"sheet": sheet, "status": "skip", "reason": "빈 시트"})
                continue
            layout = find_header_layout(raw)
            if layout is None:
                logs.append({"sheet": sheet, "status": "skip", "reason": "헤더 인식 실패"})
                continue
            current_product = ""
            added = 0
            for r in range(layout["data_start_row"], len(raw)):
                product_cell = normalize_text(raw.iat[r, layout["product_col"]]) if layout["product_col"] < len(raw.columns) else ""
                if product_cell:
                    current_product = product_cell
                if not current_product or any(x in current_product for x in ["품목코드", "색상", "Packing", "Division", "주간계", "weekly"]):
                    continue
                color = normalize_text(raw.iat[r, layout["color_col"]]) if layout["color_col"] < len(raw.columns) else ""
                for c, d in layout["date_cols"]:
                    qty = to_int(raw.iat[r, c], 0)
                    if qty > 0:
                        all_rows.append({"sheet": str(sheet), "date": d, "product_code": current_product, "color": color, "plan_qty": qty})
                        added += 1
            logs.append({"sheet": sheet, "status": "ok" if added else "empty", "rows": added})
        except Exception as e:
            logs.append({"sheet": sheet, "status": "error", "reason": f"{type(e).__name__}: {e}"})
    plan_df = pd.DataFrame(all_rows)
    if plan_df.empty:
        plan_df = pd.DataFrame(columns=["sheet", "date", "product_code", "color", "plan_qty"])
    else:
        plan_df = plan_df.groupby(["date", "product_code", "color"], as_index=False)["plan_qty"].sum()
    return plan_df, pd.DataFrame(logs)


def aggregate_parts(parts_df, mix_same):
    if parts_df.empty:
        return pd.DataFrame(columns=["group_name", "color", "thickness_mm", "width_mm", "height_mm", "part_code", "qty"])
    if mix_same:
        group_cols = ["color", "thickness_mm", "width_mm", "height_mm", "part_code"]
        agg = parts_df.groupby(group_cols, dropna=False).agg(qty=("qty", "sum")).reset_index()
        agg["group_name"] = agg.apply(lambda r: f"색상:{r['color']} / 두께:{r['thickness_mm']}", axis=1)
    else:
        group_cols = ["product_code", "color", "thickness_mm", "width_mm", "height_mm", "part_code"]
        agg = parts_df.groupby(group_cols, dropna=False).agg(qty=("qty", "sum")).reset_index()
        agg["group_name"] = agg.apply(lambda r: f"제품:{r['product_code']} / 색상:{r['color']} / 두께:{r['thickness_mm']}", axis=1)
    return agg.sort_values(["group_name", "width_mm", "height_mm"], ascending=[True, False, False])


def expand_agg_parts(agg_df):
    expanded = []
    for _, p in agg_df.iterrows():
        repeat = min(max(1, to_int(p["qty"], 1)), 50)  # runaway protection
        for _ in range(repeat):
            expanded.append({
                "group_name": p["group_name"],
                "part_code": p["part_code"],
                "color": p["color"],
                "thickness_mm": float(p["thickness_mm"]),
                "width_mm": float(p["width_mm"]),
                "height_mm": float(p["height_mm"]),
            })
    expanded.sort(key=lambda x: x["width_mm"] * x["height_mm"], reverse=True)
    return expanded


def optimize_parts_fast(parts_df, board_width, board_height, kerf, margin, rotate_allowed, mix_same):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    agg_df = aggregate_parts(parts_df, mix_same)
    expanded = expand_agg_parts(agg_df)

    sheets = []
    for part in expanded:
        placed = False
        for sheet in sheets:
            free_rects = sheet["free_rects"]
            best = None
            variants = [(part["width_mm"], part["height_mm"])]
            if rotate_allowed and abs(part["width_mm"] - part["height_mm"]) > 1e-9:
                variants.append((part["height_mm"], part["width_mm"]))
            for idx, rect in enumerate(free_rects):
                for w, h in variants:
                    if w <= rect["w"] and h <= rect["h"]:
                        score = rect["w"] * rect["h"] - w * h
                        if best is None or score < best["score"]:
                            best = {"idx": idx, "w": w, "h": h, "score": score}
            if best is not None:
                rect = free_rects.pop(best["idx"])
                sheet["placements"].append({
                    "group_name": part["group_name"],
                    "part_code": part["part_code"],
                    "color": part["color"],
                    "thickness_mm": part["thickness_mm"],
                    "x_mm": round(rect["x"] + margin, 1),
                    "y_mm": round(rect["y"] + margin, 1),
                    "width_mm": round(best["w"], 1),
                    "height_mm": round(best["h"], 1),
                })
                right_w = rect["w"] - best["w"] - kerf
                bottom_h = rect["h"] - best["h"] - kerf
                if right_w > 0:
                    free_rects.append({"x": rect["x"] + best["w"] + kerf, "y": rect["y"], "w": right_w, "h": best["h"]})
                if bottom_h > 0:
                    free_rects.append({"x": rect["x"], "y": rect["y"] + best["h"] + kerf, "w": rect["w"], "h": bottom_h})
                placed = True
                break
        if not placed:
            if len(sheets) >= MAX_SHEETS_RENDER:
                break
            sheet = {
                "sheet_no": len(sheets) + 1,
                "placements": [],
                "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}],
            }
            sheets.append(sheet)
            rect = sheet["free_rects"].pop(0)
            sheet["placements"].append({
                "group_name": part["group_name"],
                "part_code": part["part_code"],
                "color": part["color"],
                "thickness_mm": part["thickness_mm"],
                "x_mm": round(rect["x"] + margin, 1),
                "y_mm": round(rect["y"] + margin, 1),
                "width_mm": round(part["width_mm"], 1),
                "height_mm": round(part["height_mm"], 1),
            })
            right_w = rect["w"] - part["width_mm"] - kerf
            bottom_h = rect["h"] - part["height_mm"] - kerf
            if right_w > 0:
                sheet["free_rects"].append({"x": rect["x"] + part["width_mm"] + kerf, "y": rect["y"], "w": right_w, "h": part["height_mm"]})
            if bottom_h > 0:
                sheet["free_rects"].append({"x": rect["x"], "y": rect["y"] + part["height_mm"] + kerf, "w": rect["w"], "h": bottom_h})

    total_area = sum(p["width_mm"] * p["height_mm"] for s in sheets for p in s["placements"])
    board_area = board_width * board_height
    used = len(sheets)
    yield_rate = round((total_area / (used * board_area)) * 100, 2) if used else 0.0
    return {"board_width_mm": board_width, "board_height_mm": board_height, "used_boards": used, "yield_rate": yield_rate, "waste_area": max(0.0, used * board_area - total_area), "sheets": sheets}


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


def build_workorder_excel(summary_df, placement_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="작업지시요약", index=False)
        if placement_df is not None and not placement_df.empty:
            placement_df.to_excel(writer, sheet_name="분할목록", index=False)
    output.seek(0)
    return output.getvalue()


try:
    st.subheader("BOM 데이터 소스")
    bom_url = st.text_input("BOM URL (기본 고정)", value=DEFAULT_BOM_URL)
    st.caption("기본값은 GitHub raw BOM 주소로 고정되어 있습니다.")
    bom_file = st.file_uploader("또는 BOM 엑셀 업로드", type=["xlsx", "xls"], key="bom")
    plan_file = st.file_uploader("주차별 생산계획 업로드", type=["xlsx", "xls"], key="plan")

    bom_df = None
    if bom_file is not None or normalize_text(bom_url):
        raw_bom = read_bom(bom_file, bom_url)
        bom_df, bom_errors = load_bom(raw_bom)
        st.success(f"BOM 로드 완료: {len(bom_df)}행")
        with st.expander("BOM 전체 데이터 보기", expanded=False):
            st.dataframe(bom_df.head(300), width="stretch", height=420)

    if bom_df is not None and plan_file is not None:
        plan_df, parse_log_df = parse_plan_workbook_auto(plan_file)
        if len(plan_df) > MAX_PLAN_ROWS:
            st.warning(f"생산계획 데이터가 많아 상위 {MAX_PLAN_ROWS}건만 사용합니다.")
            plan_df = plan_df.head(MAX_PLAN_ROWS)

        with st.expander("생산계획 해석 로그", expanded=True):
            st.dataframe(parse_log_df, width="stretch", height=220)
        with st.expander("해석된 생산계획 보기", expanded=True):
            st.dataframe(plan_df, width="stretch", height=260)

        if not plan_df.empty:
            board_preset = st.selectbox("원장 규격 선택", list(BOARD_PRESETS.keys()), index=0)
            bw, bh = BOARD_PRESETS[board_preset]
            c1, c2, c3 = st.columns(3)
            kerf = c1.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
            margin = c2.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
            rotate_allowed = c3.checkbox("회전 허용", value=True)
            mix_same = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=True)

            date_options = sorted(plan_df["date"].unique().tolist())
            selected_date = st.selectbox("작업 날짜 선택", date_options)

            date_plan = plan_df[plan_df["date"] == selected_date].copy()

            parts_rows = []
            for _, row in date_plan.iterrows():
                product_code = normalize_text(row["product_code"])
                color = normalize_text(row["color"])
                plan_qty = to_int(row["plan_qty"], 0)
                matched = bom_df[(bom_df["product_code"] == product_code) & (bom_df["is_cutting_target"] == True)].copy()
                if color:
                    matched_color = matched[matched["color"] == color].copy()
                    if not matched_color.empty:
                        matched = matched_color
                if matched.empty:
                    continue
                matched["qty"] = matched["bom_qty"].astype(int) * plan_qty
                parts_rows.append(matched)

            if parts_rows:
                parts_df = pd.concat(parts_rows, ignore_index=True)
                result = optimize_parts_fast(parts_df, bw, bh, float(kerf), float(margin), rotate_allowed, mix_same)

                summary_df = (
                    parts_df.groupby(["color", "thickness_mm"], dropna=False)
                    .agg(총_재단_수량=("qty", "sum"), 품목수=("part_code", "count"))
                    .reset_index()
                    .rename(columns={"color": "색상", "thickness_mm": "두께(mm)", "총_재단_수량": "총 재단 수량"})
                )
                summary_df["사용 원장 수"] = result["used_boards"]
                summary_df["수율(%)"] = result["yield_rate"]
                summary_df["자투리 면적"] = result["waste_area"]

                placement_rows = []
                for s in result["sheets"]:
                    for p in s["placements"]:
                        placement_rows.append({
                            "date": selected_date,
                            "sheet_no": s["sheet_no"],
                            "group_name": p["group_name"],
                            "part_code": p["part_code"],
                            "color": p["color"],
                            "thickness_mm": p["thickness_mm"],
                            "x_mm": p["x_mm"],
                            "y_mm": p["y_mm"],
                            "width_mm": p["width_mm"],
                            "height_mm": p["height_mm"],
                        })
                placement_df = pd.DataFrame(placement_rows)

                st.subheader(f"{selected_date} 작업지시서")
                m1, m2, m3 = st.columns(3)
                m1.metric("사용 원장 수", result["used_boards"])
                m2.metric("수율", f"{result['yield_rate']}%")
                m3.metric("자투리 면적", f"{result['waste_area']:,}")

                st.dataframe(summary_df, width="stretch", height=220)

                workorder_bytes = build_workorder_excel(summary_df, placement_df)
                st.download_button(
                    "작업지시서 엑셀 다운로드",
                    data=workorder_bytes,
                    file_name=f"wood_cutting_workorder_{selected_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )

                if result["sheets"]:
                    labels = [f"Sheet {s['sheet_no']}" for s in result["sheets"]]
                    selected_sheet_label = st.selectbox("시트 선택", labels)
                    selected_sheet_no = int(selected_sheet_label.replace("Sheet", "").strip())
                    selected_sheet = next(s for s in result["sheets"] if s["sheet_no"] == selected_sheet_no)
                    components.html(make_svg(selected_sheet, result["board_width_mm"], result["board_height_mm"]), height=700, scrolling=True)

                if not placement_df.empty:
                    with st.expander("상세 배치 목록", expanded=False):
                        st.dataframe(placement_df, width="stretch", height=260)
            else:
                st.warning("선택한 날짜에 BOM과 매칭된 재단 데이터가 없습니다.")
        else:
            st.warning("주차별 생산계획을 자동 해석하지 못했습니다.")

except Exception as e:
    st.error(f"앱 실행 오류: {type(e).__name__}: {e}")
    with st.expander("오류 상세", expanded=True):
        st.code(traceback.format_exc())
