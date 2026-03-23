import io
import re
import traceback
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

DEFAULT_BOM_URL = "https://raw.githubusercontent.com/kjw26/wood-cutting-app-blob-main-app.py/main/BOM_DATA.xlsx"
BOARD_PRESETS = {"4x8 (1220 x 2440)": (2440.0, 1220.0), "4x6 (1220 x 1830)": (1830.0, 1220.0)}
MAX_PLAN_ROWS = 500
MAX_EXPANDED_PARTS = 2000
MAX_SHEETS_RENDER = 40

st.set_page_config(page_title="목재 재단 프로그램 v23", layout="wide")

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
    candidates = [text]
    cleaned = re.sub(r"\([^)]*\)", "", text)
    cleaned = cleaned.replace(".", "/").replace("-", "/")
    cleaned = re.sub(r"\s+", "", cleaned)
    cleaned_no_letters = re.sub(r"[A-Za-z가-힣]", "", cleaned)
    candidates.extend([cleaned, cleaned_no_letters])
    for cand in candidates:
        cand = cand.strip()
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
    errors = []
    for idx, row in df.iterrows():
        raw = row.to_dict()
        w, h, t = parse_spec(raw.get("규격"))
        bom_qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        item = {
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
        }
        items.append(item)
        if not item["product_code"]:
            errors.append({"row": item["row_no"], "field": "품목코드", "message": "누락"})
        if not item["part_code"]:
            errors.append({"row": item["row_no"], "field": "부품코드", "message": "누락"})
    bom_df = pd.DataFrame(items)
    err_df = pd.DataFrame(errors) if errors else pd.DataFrame(columns=["row", "field", "message"])
    return bom_df, err_df

def detect_text_column(raw, keywords, rows=25, cols=60):
    best = None
    best_row = None
    rows = min(len(raw), rows)
    cols = min(len(raw.columns), cols)
    for r in range(rows):
        for c in range(cols):
            val = normalize_text(raw.iat[r, c]).replace(" ", "")
            if any(k in val for k in keywords):
                if best is None or r < best_row:
                    best = c
                    best_row = r
    return best, best_row

def detect_date_columns(raw, rows=25, cols=60):
    rows = min(len(raw), rows)
    cols = min(len(raw.columns), cols)
    row_hits = []
    for r in range(rows):
        hits = []
        for c in range(cols):
            ts = parse_date_like(raw.iat[r, c])
            if ts is not None and ts.weekday() <= 5:
                hits.append((c, str(ts.date())))
        unique_hits = []
        seen = set()
        for c, d in hits:
            if c not in seen:
                unique_hits.append((c, d))
                seen.add(c)
        if len(unique_hits) >= 2:
            row_hits.append((r, unique_hits))
    if not row_hits:
        return None, None
    row_hits.sort(key=lambda x: (-len(x[1]), x[0]))
    return row_hits[0][0], row_hits[0][1]

def detect_data_start_row(raw, product_col, date_cols, search_start):
    rows = len(raw)
    start = min(rows - 1, max(0, search_start))
    for r in range(start, rows):
        prod = normalize_text(raw.iat[r, product_col]) if product_col < len(raw.columns) else ""
        qty_hits = 0
        for c, _ in date_cols:
            if c >= len(raw.columns):
                continue
            if to_int(raw.iat[r, c], 0) > 0:
                qty_hits += 1
        if prod or qty_hits > 0:
            return r
    return min(start + 1, max(rows - 1, 0))

def find_header_layout(raw):
    rows = min(len(raw), 60)
    cols = min(len(raw.columns), 70)
    product_col, product_row = detect_text_column(raw, ["품목코드", "productcode", "product_code"], rows, cols)
    color_col, color_row = detect_text_column(raw, ["색상", "color"], rows, cols)
    date_row, date_cols = detect_date_columns(raw, rows, cols)
    if product_col is None or date_cols is None:
        return None
    if color_col is None:
        color_col = min(product_col + 1, cols - 1)
    header_anchor = min([x for x in [product_row, color_row, date_row] if x is not None])
    data_start_row = detect_data_start_row(raw, product_col, date_cols, max(header_anchor + 1, date_row))
    return {"product_col": product_col, "color_col": color_col, "date_cols": date_cols, "header_row": header_anchor, "date_row": date_row, "data_start_row": data_start_row}

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
                if not current_product:
                    continue
                if any(x in current_product for x in ["품목코드", "색상", "Packing", "Division", "주간계", "weekly", "Date", "작성일"]):
                    continue
                color = normalize_text(raw.iat[r, layout["color_col"]]) if layout["color_col"] < len(raw.columns) else ""
                if not color:
                    continue
                row_added = 0
                for c, d in layout["date_cols"]:
                    if c >= len(raw.columns):
                        continue
                    qty = to_int(raw.iat[r, c], 0)
                    if qty > 0:
                        all_rows.append({"sheet": str(sheet), "date": d, "product_code": current_product, "color": color, "plan_qty": qty})
                        row_added += 1
                        added += 1
                if row_added == 0 and not product_cell:
                    continue
            logs.append({"sheet": sheet, "status": "ok" if added else "empty", "header_row": layout["header_row"], "date_row": layout["date_row"], "data_start_row": layout["data_start_row"], "product_col": layout["product_col"], "color_col": layout["color_col"], "date_col_count": len(layout["date_cols"]), "rows": added})
        except Exception as e:
            logs.append({"sheet": sheet, "status": "error", "reason": f"{type(e).__name__}: {e}"})
    plan_df = pd.DataFrame(all_rows)
    if plan_df.empty:
        plan_df = pd.DataFrame(columns=["sheet", "date", "product_code", "color", "plan_qty"])
    else:
        plan_df = (plan_df.groupby(["date", "product_code", "color"], as_index=False)["plan_qty"].sum()
                   .sort_values(["date", "product_code", "color"])
                   .drop_duplicates(subset=["date", "product_code", "color"], keep="last")
                   .reset_index(drop=True))
    return plan_df, pd.DataFrame(logs)

def read_scrap(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame(columns=["scrap_id", "color", "thickness_mm", "width_mm", "height_mm", "qty"])
    df = pd.read_excel(uploaded_file).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    def find_col(*candidates):
        lowered = {c.lower(): c for c in df.columns}
        for cand in candidates:
            if cand in df.columns:
                return cand
            if cand.lower() in lowered:
                return lowered[cand.lower()]
        return None
    out = pd.DataFrame()
    id_col = find_col("scrap_id", "자투리ID", "ID")
    color_col = find_col("color", "색상")
    thick_col = find_col("thickness_mm", "두께", "두께(mm)")
    width_col = find_col("width_mm", "가로", "가로(mm)")
    height_col = find_col("height_mm", "세로", "세로(mm)")
    qty_col = find_col("qty", "수량")
    out["scrap_id"] = df[id_col] if id_col else [f"S{i+1}" for i in range(len(df))]
    out["color"] = df[color_col] if color_col else ""
    out["thickness_mm"] = pd.to_numeric(df[thick_col], errors="coerce").fillna(0) if thick_col else 0
    out["width_mm"] = pd.to_numeric(df[width_col], errors="coerce").fillna(0) if width_col else 0
    out["height_mm"] = pd.to_numeric(df[height_col], errors="coerce").fillna(0) if height_col else 0
    out["qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(1).astype(int) if qty_col else 1
    return out

def aggregate_parts(parts_df, mix_same):
    if parts_df.empty:
        return pd.DataFrame(columns=["group_name", "part_code", "color", "thickness_mm", "width_mm", "height_mm", "qty"])
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
        repeat = min(max(1, to_int(p["qty"], 1)), 50)
        for _ in range(repeat):
            expanded.append({"group_name": p["group_name"], "part_code": p["part_code"], "color": p["color"], "thickness_mm": float(p["thickness_mm"]), "width_mm": float(p["width_mm"]), "height_mm": float(p["height_mm"])})
    expanded.sort(key=lambda x: x["width_mm"] * x["height_mm"], reverse=True)
    return expanded[:MAX_EXPANDED_PARTS]

def try_place_in_rects(free_rects, part, rotate_allowed):
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
    return best

def optimize_parts_fast(parts_df, scrap_df, board_width, board_height, kerf, margin, rotate_allowed, mix_same):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    agg_df = aggregate_parts(parts_df, mix_same)
    expanded = expand_agg_parts(agg_df)
    sheets = []
    scrap_used = []
    leftovers = expanded.copy()
    if scrap_df is not None and not scrap_df.empty:
        scrap_rects = []
        for _, s in scrap_df.iterrows():
            qty = max(1, to_int(s.get("qty", 1), 1))
            for _ in range(qty):
                scrap_rects.append({"scrap_id": normalize_text(s.get("scrap_id")), "color": normalize_text(s.get("color")), "thickness_mm": float(s.get("thickness_mm", 0) or 0), "w": float(s.get("width_mm", 0) or 0), "h": float(s.get("height_mm", 0) or 0)})
        remaining = []
        for part in leftovers:
            placed = False
            for scrap in scrap_rects:
                if scrap["w"] <= 0 or scrap["h"] <= 0:
                    continue
                if normalize_text(scrap["color"]) and normalize_text(scrap["color"]) != normalize_text(part["color"]):
                    continue
                if float(scrap["thickness_mm"]) not in [0.0, float(part["thickness_mm"])]:
                    continue
                variants = [(part["width_mm"], part["height_mm"])]
                if rotate_allowed and abs(part["width_mm"] - part["height_mm"]) > 1e-9:
                    variants.append((part["height_mm"], part["width_mm"]))
                for w, h in variants:
                    if w <= scrap["w"] and h <= scrap["h"]:
                        scrap_used.append({"scrap_id": scrap["scrap_id"], "part_code": part["part_code"], "color": part["color"], "thickness_mm": part["thickness_mm"], "width_mm": w, "height_mm": h})
                        scrap["w"] = 0
                        scrap["h"] = 0
                        placed = True
                        break
                if placed:
                    break
            if not placed:
                remaining.append(part)
        leftovers = remaining
    for part in leftovers:
        placed = False
        for sheet in sheets:
            best = try_place_in_rects(sheet["free_rects"], part, rotate_allowed)
            if best is None:
                continue
            rect = sheet["free_rects"].pop(best["idx"])
            sheet["placements"].append({"group_name": part["group_name"], "part_code": part["part_code"], "color": part["color"], "thickness_mm": part["thickness_mm"], "x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1)})
            right_w = rect["w"] - best["w"] - kerf
            bottom_h = rect["h"] - best["h"] - kerf
            if right_w > 0:
                sheet["free_rects"].append({"x": rect["x"] + best["w"] + kerf, "y": rect["y"], "w": right_w, "h": best["h"]})
            if bottom_h > 0:
                sheet["free_rects"].append({"x": rect["x"], "y": rect["y"] + best["h"] + kerf, "w": rect["w"], "h": bottom_h})
            placed = True
            break
        if not placed:
            if len(sheets) >= MAX_SHEETS_RENDER:
                break
            sheet = {"sheet_no": len(sheets) + 1, "placements": [], "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}]}
            sheets.append(sheet)
            rect = sheet["free_rects"].pop(0)
            sheet["placements"].append({"group_name": part["group_name"], "part_code": part["part_code"], "color": part["color"], "thickness_mm": part["thickness_mm"], "x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(part["width_mm"], 1), "height_mm": round(part["height_mm"], 1)})
            right_w = rect["w"] - part["width_mm"] - kerf
            bottom_h = rect["h"] - part["height_mm"] - kerf
            if right_w > 0:
                sheet["free_rects"].append({"x": rect["x"] + part["width_mm"] + kerf, "y": rect["y"], "w": right_w, "h": part["height_mm"]})
            if bottom_h > 0:
                sheet["free_rects"].append({"x": rect["x"], "y": rect["y"] + part["height_mm"] + kerf, "w": rect["w"], "h": bottom_h})
    total_area = sum(p["width_mm"] * p["height_mm"] for s in sheets for p in s["placements"]) + sum(p["width_mm"] * p["height_mm"] for p in scrap_used)
    board_area = board_width * board_height
    used = len(sheets)
    yield_rate = round((total_area / (used * board_area)) * 100, 2) if used else 0.0
    return {"board_width_mm": board_width, "board_height_mm": board_height, "used_boards": used, "yield_rate": yield_rate, "waste_area": max(0.0, used * board_area - (total_area - sum(p["width_mm"] * p["height_mm"] for p in scrap_used))), "sheets": sheets, "scrap_used": scrap_used}

def scenario_table(parts_df, scrap_df, kerf, margin, rotate_allowed, mix_same):
    scenarios = [("현재 조건", 2440.0, 1220.0, margin), ("4x8 / 여유치 5", 2440.0, 1220.0, 5.0), ("4x8 / 여유치 10", 2440.0, 1220.0, 10.0), ("4x6 / 여유치 10", 1830.0, 1220.0, 10.0)]
    rows = []
    for name, bw, bh, mg in scenarios:
        r = optimize_parts_fast(parts_df, scrap_df, bw, bh, kerf, mg, rotate_allowed, mix_same)
        rows.append({"시나리오": name, "원장 가로": bw, "원장 세로": bh, "여유치": mg, "사용 원장 수": r["used_boards"], "수율(%)": r["yield_rate"], "자투리 사용 수": len(r["scrap_used"])})
    return pd.DataFrame(rows)

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

def build_workorder_excel(summary_df, placement_df, scrap_used_df, scenario_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="작업지시요약", index=False)
        if placement_df is not None and not placement_df.empty:
            placement_df.to_excel(writer, sheet_name="분할목록", index=False)
        if scrap_used_df is not None and not scrap_used_df.empty:
            scrap_used_df.to_excel(writer, sheet_name="자투리사용", index=False)
        if scenario_df is not None and not scenario_df.empty:
            scenario_df.to_excel(writer, sheet_name="시나리오분석", index=False)
    output.seek(0)
    return output.getvalue()

try:
    st.title("목재 재단 프로그램 v23")
    st.caption("고도화 파서 + 중복 제거 + 색상 공백 제거")
    st.subheader("입력")
    bom_url = st.text_input("BOM URL (기본 고정)", value=DEFAULT_BOM_URL)
    bom_file = st.file_uploader("또는 BOM 엑셀 업로드", type=["xlsx", "xls"], key="bom")
    plan_file = st.file_uploader("주차별 생산계획 업로드", type=["xlsx", "xls"], key="plan")
    scrap_file = st.file_uploader("자투리 재고 업로드 (선택)", type=["xlsx", "xls"], key="scrap")
    bom_df = None
    if bom_file is not None or normalize_text(bom_url):
        raw_bom = read_bom(bom_file, bom_url)
        bom_df, bom_errors = load_bom(raw_bom)
        st.success(f"BOM 로드 완료: {len(bom_df)}행")
        with st.expander("BOM 보기", expanded=False):
            st.dataframe(bom_df.head(300), width="stretch", height=320)
    if bom_df is not None and plan_file is not None:
        plan_df, parse_log_df = parse_plan_workbook_auto(plan_file)
        scrap_df = read_scrap(scrap_file)
        if len(plan_df) > MAX_PLAN_ROWS:
            st.warning(f"생산계획 데이터가 많아 상위 {MAX_PLAN_ROWS}건만 사용합니다.")
            plan_df = plan_df.head(MAX_PLAN_ROWS)
        with st.expander("생산계획 해석 로그", expanded=True):
            st.dataframe(parse_log_df, width="stretch", height=220)
        if plan_df.empty:
            st.warning("주차별 생산계획을 자동 해석하지 못했습니다.")
        else:
            st.success(f"생산계획 해석 완료: {len(plan_df)}건")
            with st.expander("해석된 생산계획 보기", expanded=False):
                st.dataframe(plan_df, width="stretch", height=260)
            c1, c2, c3, c4 = st.columns(4)
            board_preset = c1.selectbox("원장 규격", list(BOARD_PRESETS.keys()), index=0)
            kerf = c2.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
            margin = c3.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
            rotate_allowed = c4.checkbox("회전 허용", value=True)
            mix_same = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=True)
            bw, bh = BOARD_PRESETS[board_preset]
            date_options = sorted(plan_df["date"].unique().tolist())
            selected_date = st.selectbox("작업 날짜 선택", date_options)
            date_plan = plan_df[plan_df["date"] == selected_date].copy()
            parts_rows = []
            matching_rows = []
            for _, row in date_plan.iterrows():
                product_code = normalize_text(row["product_code"])
                color = normalize_text(row["color"])
                plan_qty = to_int(row["plan_qty"], 0)
                matched = bom_df[(bom_df["product_code"] == product_code) & (bom_df["is_cutting_target"] == True)].copy()
                if color:
                    matched_color = matched[matched["color"] == color].copy()
                    if not matched_color.empty:
                        matched = matched_color
                matching_rows.append({"date": selected_date, "product_code": product_code, "plan_color": color, "plan_qty": plan_qty, "matched_parts": len(matched)})
                if matched.empty:
                    continue
                matched["qty"] = matched["bom_qty"].astype(int) * plan_qty
                parts_rows.append(matched)
            st.subheader(f"{selected_date} 매칭 결과")
            st.dataframe(pd.DataFrame(matching_rows), width="stretch", height=180)
            if parts_rows:
                parts_df = pd.concat(parts_rows, ignore_index=True)
                result = optimize_parts_fast(parts_df, scrap_df, bw, bh, float(kerf), float(margin), rotate_allowed, mix_same)
                scenario_df = scenario_table(parts_df, scrap_df, float(kerf), float(margin), rotate_allowed, mix_same)
                summary_df = (parts_df.groupby(["color", "thickness_mm"], dropna=False)
                              .agg(총_재단_수량=("qty", "sum"), 품목수=("part_code", "count"))
                              .reset_index()
                              .rename(columns={"color": "색상", "thickness_mm": "두께(mm)", "총_재단_수량": "총 재단 수량"}))
                summary_df["사용 원장 수"] = result["used_boards"]
                summary_df["수율(%)"] = result["yield_rate"]
                summary_df["자투리 사용 수"] = len(result["scrap_used"])
                summary_df["자투리 면적"] = result["waste_area"]
                placement_rows = []
                for s in result["sheets"]:
                    for p in s["placements"]:
                        placement_rows.append({"date": selected_date, "sheet_no": s["sheet_no"], "group_name": p["group_name"], "part_code": p["part_code"], "color": p["color"], "thickness_mm": p["thickness_mm"], "x_mm": p["x_mm"], "y_mm": p["y_mm"], "width_mm": p["width_mm"], "height_mm": p["height_mm"]})
                placement_df = pd.DataFrame(placement_rows)
                scrap_used_df = pd.DataFrame(result["scrap_used"])
                st.subheader(f"{selected_date} 작업지시서")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("사용 원장 수", result["used_boards"])
                m2.metric("수율", f"{result['yield_rate']}%")
                m3.metric("자투리 사용 수", len(result["scrap_used"]))
                m4.metric("자투리 면적", f"{result['waste_area']:,}")
                st.markdown("#### 색상 / 두께별 작업 요약")
                st.dataframe(summary_df, width="stretch", height=220)
                st.markdown("#### 시나리오 분석")
                st.dataframe(scenario_df, width="stretch", height=200)
                workorder_bytes = build_workorder_excel(summary_df, placement_df, scrap_used_df, scenario_df)
                st.download_button("작업지시서 엑셀 다운로드", data=workorder_bytes, file_name=f"wood_cutting_workorder_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
                if result["sheets"]:
                    labels = [f"Sheet {s['sheet_no']}" for s in result["sheets"]]
                    selected_sheet_label = st.selectbox("시트 선택", labels)
                    selected_sheet_no = int(selected_sheet_label.replace("Sheet", "").strip())
                    selected_sheet = next(s for s in result["sheets"] if s["sheet_no"] == selected_sheet_no)
                    st.markdown("#### 분할도")
                    components.html(make_svg(selected_sheet, result["board_width_mm"], result["board_height_mm"]), height=700, scrolling=True)
                if not scrap_used_df.empty:
                    with st.expander("자투리 사용 내역", expanded=False):
                        st.dataframe(scrap_used_df, width="stretch", height=180)
                if not placement_df.empty:
                    with st.expander("상세 배치 목록", expanded=False):
                        st.dataframe(placement_df, width="stretch", height=260)
            else:
                st.warning("선택한 날짜에 BOM과 매칭된 재단 데이터가 없습니다.")
except Exception as e:
    st.error(f"앱 실행 오류: {type(e).__name__}: {e}")
    with st.expander("오류 상세", expanded=True):
        st.code(traceback.format_exc())
