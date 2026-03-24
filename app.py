import io, re, traceback
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

DEFAULT_BOM_URL = "https://raw.githubusercontent.com/kjw26/wood-cutting-app-blob-main-app.py/main/BOM_DATA.xlsx"
BOARD_PRESETS = {"4x8 (1220 x 2440)": (2440.0, 1220.0), "4x6 (1220 x 1830)": (1830.0, 1220.0)}
MAX_PLAN_ROWS = 800
MAX_EXPANDED_PARTS = 3000
MAX_SHEETS_RENDER = 80
st.set_page_config(page_title="목재 재단 프로그램 v27", layout="wide")

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
    cands = [text, re.sub(r"\([^)]*\)", "", text)]
    cands += [re.sub(r"[A-Za-z가-힣]", "", cands[-1]).replace(".", "/").replace("-", "/").strip()]
    for cand in cands:
        if not cand:
            continue
        cand = cand.replace(" ", "")
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
    items, errors = [], []
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
    err_df = pd.DataFrame(errors) if errors else pd.DataFrame(columns=["row", "field", "message"])
    return pd.DataFrame(items), err_df

def detect_text_column(raw, keywords, rows=25, cols=60):
    best, best_row = None, None
    rows = min(len(raw), rows)
    cols = min(len(raw.columns), cols)
    for r in range(rows):
        for c in range(cols):
            val = normalize_text(raw.iat[r, c]).replace(" ", "")
            if any(k in val for k in keywords):
                if best is None or r < best_row:
                    best, best_row = c, r
    return best, best_row

def detect_date_columns(raw, rows=25, cols=60):
    rows = min(len(raw), rows)
    cols = min(len(raw.columns), cols)
    row_hits = []
    for r in range(rows):
        uniq, seen = [], set()
        for c in range(cols):
            ts = parse_date_like(raw.iat[r, c])
            if ts is not None and ts.weekday() <= 5 and c not in seen:
                uniq.append((c, str(ts.date())))
                seen.add(c)
        if len(uniq) >= 2:
            row_hits.append((r, uniq))
    if not row_hits:
        return None, None
    row_hits.sort(key=lambda x: (-len(x[1]), x[0]))
    return row_hits[0][0], row_hits[0][1]

def detect_data_start_row(raw, product_col, date_cols, search_start):
    rows = len(raw)
    start = min(rows - 1, max(0, search_start))
    for r in range(start, rows):
        prod = normalize_text(raw.iat[r, product_col]) if product_col < len(raw.columns) else ""
        qty_hits = sum(1 for c, _ in date_cols if c < len(raw.columns) and to_int(raw.iat[r, c], 0) > 0)
        if prod or qty_hits > 0:
            return r
    return min(start + 1, max(rows - 1, 0))

def find_header_layout(raw):
    rows, cols = min(len(raw), 60), min(len(raw.columns), 70)
    product_col, product_row = detect_text_column(raw, ["품목코드", "productcode", "product_code"], rows, cols)
    color_col, color_row = detect_text_column(raw, ["색상", "color"], rows, cols)
    date_row, date_cols = detect_date_columns(raw, rows, cols)
    if product_col is None or date_cols is None:
        return None
    if color_col is None:
        color_col = min(product_col + 1, cols - 1)
    header_anchor = min(x for x in [product_row, color_row, date_row] if x is not None)
    data_start_row = detect_data_start_row(raw, product_col, date_cols, max(header_anchor + 1, date_row))
    return {"product_col": product_col, "color_col": color_col, "date_cols": date_cols, "header_row": header_anchor, "date_row": date_row, "data_start_row": data_start_row}

def parse_plan_workbook_auto(file):
    xls = pd.ExcelFile(file)
    all_rows, logs = [], []
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
            current_product, added = "", 0
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
                for c, d in layout["date_cols"]:
                    if c >= len(raw.columns):
                        continue
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
        plan_df = plan_df.groupby(["date", "product_code", "color"], as_index=False)["plan_qty"].sum().sort_values(["date", "product_code", "color"]).reset_index(drop=True)
    return plan_df, pd.DataFrame(logs)

def read_scrap(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame(columns=["scrap_id", "color", "thickness_mm", "width_mm", "height_mm", "qty"])
    df = pd.read_excel(uploaded_file).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    def find_col(*cands):
        lowered = {c.lower(): c for c in df.columns}
        for cand in cands:
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
        return pd.DataFrame(columns=["group_name", "part_code", "part_name", "product_code", "date", "color", "thickness_mm", "width_mm", "height_mm", "qty"])
    if mix_same:
        group_cols = ["color", "thickness_mm", "width_mm", "height_mm", "part_code", "part_name", "product_code", "date"]
        agg = parts_df.groupby(group_cols, dropna=False).agg(qty=("qty", "sum")).reset_index()
        agg["group_name"] = agg.apply(lambda r: f"색상:{r['color']} / 두께:{r['thickness_mm']}", axis=1)
    else:
        group_cols = ["product_code", "color", "thickness_mm", "width_mm", "height_mm", "part_code", "part_name", "date"]
        agg = parts_df.groupby(group_cols, dropna=False).agg(qty=("qty", "sum")).reset_index()
        agg["group_name"] = agg.apply(lambda r: f"제품:{r['product_code']} / 색상:{r['color']} / 두께:{r['thickness_mm']}", axis=1)
    return agg.sort_values(["group_name", "width_mm", "height_mm"], ascending=[True, False, False])

def expand_agg_parts(agg_df):
    expanded = []
    for _, p in agg_df.iterrows():
        repeat = min(max(1, to_int(p["qty"], 1)), 50)
        for _ in range(repeat):
            expanded.append({"group_name": p["group_name"], "part_code": p["part_code"], "part_name": p["part_name"], "product_code": p["product_code"], "date": p["date"], "color": p["color"], "thickness_mm": float(p["thickness_mm"]), "width_mm": float(p["width_mm"]), "height_mm": float(p["height_mm"])})
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
    usable_w, usable_h = board_width - margin * 2, board_height - margin * 2
    expanded = expand_agg_parts(aggregate_parts(parts_df, mix_same))
    sheets = []
    leftovers = expanded.copy()
    for part in leftovers:
        placed = False
        for sheet in sheets:
            best = try_place_in_rects(sheet["free_rects"], part, rotate_allowed)
            if best is None:
                continue
            rect = sheet["free_rects"].pop(best["idx"])
            sheet["placements"].append({"group_name": part["group_name"], "part_code": part["part_code"], "part_name": part["part_name"], "product_code": part["product_code"], "date": part["date"], "color": part["color"], "thickness_mm": part["thickness_mm"], "x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(best["w"], 1), "height_mm": round(best["h"], 1)})
            rw, bh = rect["w"] - best["w"] - kerf, rect["h"] - best["h"] - kerf
            if rw > 0:
                sheet["free_rects"].append({"x": rect["x"] + best["w"] + kerf, "y": rect["y"], "w": rw, "h": best["h"]})
            if bh > 0:
                sheet["free_rects"].append({"x": rect["x"], "y": rect["y"] + best["h"] + kerf, "w": rect["w"], "h": bh})
            placed = True
            break
        if not placed:
            if len(sheets) >= MAX_SHEETS_RENDER:
                break
            sheet = {"sheet_no": len(sheets) + 1, "placements": [], "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}]}
            sheets.append(sheet)
            rect = sheet["free_rects"].pop(0)
            sheet["placements"].append({"group_name": part["group_name"], "part_code": part["part_code"], "part_name": part["part_name"], "product_code": part["product_code"], "date": part["date"], "color": part["color"], "thickness_mm": part["thickness_mm"], "x_mm": round(rect["x"] + margin, 1), "y_mm": round(rect["y"] + margin, 1), "width_mm": round(part["width_mm"], 1), "height_mm": round(part["height_mm"], 1)})
            rw, bh = rect["w"] - part["width_mm"] - kerf, rect["h"] - part["height_mm"] - kerf
            if rw > 0:
                sheet["free_rects"].append({"x": rect["x"] + part["width_mm"] + kerf, "y": rect["y"], "w": rw, "h": part["height_mm"]})
            if bh > 0:
                sheet["free_rects"].append({"x": rect["x"], "y": rect["y"] + part["height_mm"] + kerf, "w": rect["w"], "h": bh})
    total_area = sum(p["width_mm"] * p["height_mm"] for s in sheets for p in s["placements"])
    board_area = board_width * board_height
    used = len(sheets)
    yield_rate = round((total_area / (used * board_area)) * 100, 2) if used else 0.0
    return {"board_width_mm": board_width, "board_height_mm": board_height, "used_boards": used, "yield_rate": yield_rate, "waste_area": max(0.0, used * board_area - total_area), "sheets": sheets, "scrap_used": []}

def exact_signature(sheet):
    sig = []
    for p in sorted(sheet["placements"], key=lambda x: (x["part_code"], x["part_name"], x["product_code"], x["date"], x["color"], x["thickness_mm"], x["width_mm"], x["height_mm"], x["x_mm"], x["y_mm"])):
        sig.append((p["part_code"], p["part_name"], p["product_code"], p["date"], p["color"], round(float(p["thickness_mm"]), 1), round(float(p["width_mm"]), 1), round(float(p["height_mm"]), 1), round(float(p["x_mm"]), 1), round(float(p["y_mm"]), 1)))
    return tuple(sig)

def compress_sheets_exact(sheets):
    groups = {}
    for s in sheets:
        sig = exact_signature(s)
        groups.setdefault(sig, {"count": 0, "sheet": s, "all_placements": []})
        groups[sig]["count"] += 1
        groups[sig]["all_placements"].extend(s["placements"])
    out = []
    for i, g in enumerate(groups.values(), start=1):
        out.append({"sheet_no": i, "placements": g["sheet"]["placements"], "count": g["count"], "all_placements": g["all_placements"]})
    return out

def scenario_table(parts_df, kerf, margin, rotate_allowed, mix_same):
    scenarios = [("현재 조건", 2440.0, 1220.0, margin), ("4x8 / 여유치 5", 2440.0, 1220.0, 5.0), ("4x8 / 여유치 10", 2440.0, 1220.0, 10.0), ("4x6 / 여유치 10", 1830.0, 1220.0, 10.0)]
    rows = []
    for name, bw, bh, mg in scenarios:
        r = optimize_parts_fast(parts_df, None, bw, bh, kerf, mg, rotate_allowed, mix_same)
        rows.append({"시나리오": name, "원장 가로": bw, "원장 세로": bh, "여유치": mg, "사용 원장 수": r["used_boards"], "수율(%)": r["yield_rate"]})
    return pd.DataFrame(rows)

def make_svg(sheet, board_width_mm, board_height_mm):
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width, svg_height = int(board_width_mm * scale), int(board_height_mm * scale)
    parts_svg = []
    for p in sheet["placements"]:
        x, y = p["x_mm"] * scale, p["y_mm"] * scale
        w, h = p["width_mm"] * scale, p["height_mm"] * scale
        t1 = f"{p['part_code']} / {p['product_code']}"
        t2 = f"{p['part_name'][:16]} / {p['color']} / {p['date'][5:]}"
        t3 = f"{p['width_mm']}x{p['height_mm']} / t{p['thickness_mm']}"
        parts_svg.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect><text x="{x+4}" y="{y+16}" font-size="11" fill="#111">{t1}</text><text x="{x+4}" y="{y+30}" font-size="11" fill="#111">{t2}</text><text x="{x+4}" y="{y+44}" font-size="11" fill="#111">{t3}</text></g>')
    badge = f'<text x="10" y="20" font-size="16" fill="#b91c1c">정밀형 동일 분할도 x{sheet.get("count",1)}장</text>' if sheet.get("count",1) > 1 else ""
    return f'<div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;"><svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>{badge}{"".join(parts_svg)}</svg></div>'

def build_pattern_workorder(sheet):
    df = pd.DataFrame(sheet["all_placements"] if "all_placements" in sheet else sheet["placements"])
    if df.empty:
        return pd.DataFrame(columns=["product_code", "part_code", "part_name", "color", "thickness_mm", "규격", "수량", "date"])
    summary = (df.groupby(["product_code", "part_code", "part_name", "color", "thickness_mm", "date"])
               .agg(수량=("part_code", "count"), 가로=("width_mm", "first"), 세로=("height_mm", "first"))
               .reset_index())
    summary["규격"] = summary["가로"].astype(str) + "x" + summary["세로"].astype(str)
    return summary[["product_code", "part_code", "part_name", "color", "thickness_mm", "규격", "수량", "date"]]

def export_pattern_workorders_excel(compressed_sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for s in compressed_sheets:
            build_pattern_workorder(s).to_excel(writer, sheet_name=f"Sheet{s['sheet_no']}"[:31], index=False)
    output.seek(0)
    return output.getvalue()

def build_workorder_excel(summary_df, placement_df, scenario_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="작업지시요약", index=False)
        if placement_df is not None and not placement_df.empty:
            placement_df.to_excel(writer, sheet_name="분할목록", index=False)
        if scenario_df is not None and not scenario_df.empty:
            scenario_df.to_excel(writer, sheet_name="시나리오분석", index=False)
    output.seek(0)
    return output.getvalue()

try:
    st.title("목재 재단 프로그램 v27")
    st.caption("정밀형 분류 + 두께 표시 + 패턴별 작업지시서")
    st.subheader("입력")
    bom_url = st.text_input("BOM URL (기본 고정)", value=DEFAULT_BOM_URL)
    bom_file = st.file_uploader("또는 BOM 엑셀 업로드", type=["xlsx", "xls"], key="bom")
    plan_file = st.file_uploader("주차별 생산계획 업로드", type=["xlsx", "xls"], key="plan")
    bom_df = None
    if bom_file is not None or normalize_text(bom_url):
        raw_bom = read_bom(bom_file, bom_url)
        bom_df, _ = load_bom(raw_bom)
        st.success(f"BOM 로드 완료: {len(bom_df)}행")
    if bom_df is not None and plan_file is not None:
        plan_df, parse_log_df = parse_plan_workbook_auto(plan_file)
        with st.expander("생산계획 해석 로그", expanded=True):
            st.dataframe(parse_log_df, width="stretch", height=220)
        if not plan_df.empty:
            c1, c2, c3, c4 = st.columns(4)
            board_preset = c1.selectbox("원장 규격", list(BOARD_PRESETS.keys()), index=0)
            kerf = c2.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
            margin = c3.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
            rotate_allowed = c4.checkbox("회전 허용", value=True)
            mix_same = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=True)
            integrate_all_dates = st.checkbox("모든 날짜 통합 혼합 재단", value=True)
            bw, bh = BOARD_PRESETS[board_preset]
            work_plan = plan_df.copy() if integrate_all_dates else plan_df[plan_df["date"] == st.selectbox("작업 날짜 선택", sorted(plan_df["date"].unique().tolist()))].copy()
            parts_rows = []
            for _, row in work_plan.iterrows():
                product_code = normalize_text(row["product_code"])
                color = normalize_text(row["color"])
                plan_qty = to_int(row["plan_qty"], 0)
                date_val = normalize_text(row["date"])
                matched = bom_df[(bom_df["product_code"] == product_code) & (bom_df["is_cutting_target"] == True)].copy()
                if color:
                    matched_color = matched[matched["color"] == color].copy()
                    if not matched_color.empty:
                        matched = matched_color
                if matched.empty:
                    continue
                matched["qty"] = matched["bom_qty"].astype(int) * plan_qty
                matched["date"] = date_val
                parts_rows.append(matched)
            if parts_rows:
                parts_df = pd.concat(parts_rows, ignore_index=True)
                result = optimize_parts_fast(parts_df, None, bw, bh, float(kerf), float(margin), rotate_allowed, mix_same)
                scenario_df = scenario_table(parts_df, float(kerf), float(margin), rotate_allowed, mix_same)
                compressed_sheets = compress_sheets_exact(result["sheets"])
                summary_df = (parts_df.groupby(["color", "thickness_mm"], dropna=False)
                              .agg(총_재단_수량=("qty", "sum"), 품목수=("part_code", "count"), 날짜수=("date", "nunique"))
                              .reset_index()
                              .rename(columns={"color": "색상", "thickness_mm": "두께(mm)", "총_재단_수량": "총 재단 수량"}))
                placement_rows = []
                for s in compressed_sheets:
                    for p in s["all_placements"]:
                        placement_rows.append({"sheet_no": s["sheet_no"], "정밀동일분할도수": s.get("count", 1), "date": p["date"], "product_code": p["product_code"], "part_code": p["part_code"], "part_name": p["part_name"], "color": p["color"], "thickness_mm": p["thickness_mm"], "x_mm": p["x_mm"], "y_mm": p["y_mm"], "width_mm": p["width_mm"], "height_mm": p["height_mm"]})
                placement_df = pd.DataFrame(placement_rows)

                st.download_button("작업지시서 엑셀 다운로드", data=build_workorder_excel(summary_df, placement_df, scenario_df), file_name="wood_cutting_workorder_v27.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
                st.download_button("패턴별 작업지시서 다운로드", data=export_pattern_workorders_excel(compressed_sheets), file_name="pattern_workorders_v27.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

                st.markdown("### 패턴별 작업지시서")
                for s in compressed_sheets:
                    st.markdown(f"#### Sheet {s['sheet_no']} (x{s.get('count',1)})")
                    st.dataframe(build_pattern_workorder(s), width="stretch", height=220)
                    components.html(make_svg(s, result["board_width_mm"], result["board_height_mm"]), height=520, scrolling=True)
except Exception as e:
    st.error(f"앱 실행 오류: {type(e).__name__}: {e}")
    with st.expander("오류 상세", expanded=True):
        st.code(traceback.format_exc())
