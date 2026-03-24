import io
import re
import traceback
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

DEFAULT_BOM_URL = "https://raw.githubusercontent.com/kjw26/wood-cutting-app-blob-main-app.py/main/BOM_DATA.xlsx"
BOARD_PRESETS = {"4x8 (1220 x 2440)": (2440.0, 1220.0), "4x6 (1220 x 1830)": (1830.0, 1220.0)}
MAX_PLAN_ROWS = 200
MAX_EXPANDED_PARTS = 500
MAX_SHEETS_RENDER = 20

st.set_page_config(page_title="목재 재단 프로그램 v31 Thickness Fix", layout="wide")

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
    candidates = [text, re.sub(r"\([^)]*\)", "", text)]
    candidates += [re.sub(r"[A-Za-z가-힣]", "", candidates[-1]).replace(".", "/").replace("-", "/").strip()]
    for cand in candidates:
        if not cand:
            continue
        cand = cand.replace(" ", "")
        if re.match(r"^\d{2}/\d{2}$", cand):
            cand = "2026/" + cand
        ts = pd.to_datetime(cand, errors="coerce")
        if not pd.isna(ts):
            return ts
    return None

@st.cache_data(show_spinner=False)
def fetch_bom_url(url):
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    return r.content

def read_bom(uploaded_file, url_text):
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file)
    url = normalize_text(url_text) or DEFAULT_BOM_URL
    return pd.read_excel(io.BytesIO(fetch_bom_url(url)))

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
    rows = []
    for idx, row in df.iterrows():
        raw = row.to_dict()
        w, h, t = parse_spec(raw.get("규격"))
        bom_qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        rows.append({
            "row_no": int(idx) + 2,
            "product_code": normalize_text(raw.get("품목코드")),
            "part_code": normalize_text(raw.get("부품코드")),
            "part_name": normalize_text(raw.get("품목명")),
            "color": normalize_text(raw.get("색상")),
            "bom_qty": max(1, bom_qty),
            "spec_raw": normalize_text(raw.get("규격")),
            "width_mm": w,
            "height_mm": h,
            "thickness_mm": t,
            "is_cutting_target": is_cutting_target(raw, w, h, t),
        })
    return pd.DataFrame(rows)

def detect_text_column(raw, keywords, rows=25, cols=60):
    best, best_row = None, None
    rows, cols = min(len(raw), rows), min(len(raw.columns), cols)
    for r in range(rows):
        for c in range(cols):
            val = normalize_text(raw.iat[r, c]).replace(" ", "")
            if any(k in val for k in keywords):
                if best is None or r < best_row:
                    best, best_row = c, r
    return best, best_row

def detect_date_columns(raw, rows=25, cols=60):
    rows, cols = min(len(raw), rows), min(len(raw.columns), cols)
    hits_by_row = []
    for r in range(rows):
        uniq, seen = [], set()
        for c in range(cols):
            ts = parse_date_like(raw.iat[r, c])
            if ts is not None and ts.weekday() <= 5 and c not in seen:
                uniq.append((c, str(ts.date())))
                seen.add(c)
        if len(uniq) >= 2:
            hits_by_row.append((r, uniq))
    if not hits_by_row:
        return None, None
    hits_by_row.sort(key=lambda x: (-len(x[1]), x[0]))
    return hits_by_row[0][0], hits_by_row[0][1]

def detect_data_start_row(raw, product_col, date_cols, search_start):
    start = min(len(raw) - 1, max(0, search_start))
    for r in range(start, len(raw)):
        prod = normalize_text(raw.iat[r, product_col]) if product_col < len(raw.columns) else ""
        qty_hits = sum(1 for c, _ in date_cols if c < len(raw.columns) and to_int(raw.iat[r, c], 0) > 0)
        if prod or qty_hits > 0:
            return r
    return start

def find_header_layout(raw):
    rows, cols = min(len(raw), 60), min(len(raw.columns), 70)
    product_col, product_row = detect_text_column(raw, ["품목코드", "productcode", "product_code"], rows, cols)
    color_col, color_row = detect_text_column(raw, ["색상", "color"], rows, cols)
    date_row, date_cols = detect_date_columns(raw, rows, cols)
    if product_col is None or date_cols is None:
        return None
    if color_col is None:
        color_col = min(product_col + 1, cols - 1)
    anchor = min(x for x in [product_row, color_row, date_row] if x is not None)
    return {
        "product_col": product_col,
        "color_col": color_col,
        "date_cols": date_cols,
        "data_start_row": detect_data_start_row(raw, product_col, date_cols, max(anchor + 1, date_row)),
    }

def parse_plan_workbook_auto(file):
    xls = pd.ExcelFile(file)
    rows, logs = [], []
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
                color = normalize_text(raw.iat[r, layout["color_col"]]) if layout["color_col"] < len(raw.columns) else ""
                if not color:
                    continue
                for c, d in layout["date_cols"]:
                    if c < len(raw.columns):
                        qty = to_int(raw.iat[r, c], 0)
                        if qty > 0:
                            rows.append({"sheet": str(sheet), "date": d, "product_code": current_product, "color": color, "plan_qty": qty})
                            added += 1
            logs.append({"sheet": sheet, "status": "ok" if added else "empty", "rows": added})
        except Exception as e:
            logs.append({"sheet": sheet, "status": "error", "reason": f"{type(e).__name__}: {e}"})
    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=["sheet", "date", "product_code", "color", "plan_qty"])
    else:
        df = df.groupby(["date", "product_code", "color"], as_index=False)["plan_qty"].sum().sort_values(["date", "product_code", "color"]).reset_index(drop=True)
    return df, pd.DataFrame(logs)

def aggregate_parts(parts_df):
    group_cols = ["product_code", "part_code", "part_name", "color", "thickness_mm", "width_mm", "height_mm", "date"]
    agg = parts_df.groupby(group_cols, dropna=False).agg(qty=("qty", "sum")).reset_index()
    return agg.sort_values(["width_mm", "height_mm"], ascending=[False, False])

def expand_agg_parts(agg_df):
    expanded = []
    for _, p in agg_df.iterrows():
        repeat = min(max(1, to_int(p["qty"], 1)), 25)
        for _ in range(repeat):
            expanded.append({
                "product_code": p["product_code"],
                "part_code": p["part_code"],
                "part_name": p["part_name"],
                "date": p["date"],
                "color": p["color"],
                "thickness_mm": float(p["thickness_mm"]),
                "width_mm": float(p["width_mm"]),
                "height_mm": float(p["height_mm"]),
            })
    expanded.sort(key=lambda x: x["width_mm"] * x["height_mm"], reverse=True)
    return expanded[:MAX_EXPANDED_PARTS]

def try_place(free_rects, part, rotate_allowed):
    best = None
    variants = [(part["width_mm"], part["height_mm"])]
    if rotate_allowed and abs(part["width_mm"] - part["height_mm"]) > 1e-9:
        variants.append((part["height_mm"], part["width_mm"]))
    for idx, rect in enumerate(free_rects):
        for w, h in variants:
            if w <= rect["w"] and h <= rect["h"]:
                waste = rect["w"] * rect["h"] - w * h
                if best is None or waste < best["waste"]:
                    best = {"idx": idx, "w": w, "h": h, "waste": waste}
    return best

def optimize_parts_fast(parts_df, board_width, board_height, kerf, margin, rotate_allowed):
    usable_w, usable_h = board_width - margin * 2, board_height - margin * 2
    expanded = expand_agg_parts(aggregate_parts(parts_df))
    sheets = []
    for part in expanded:
        placed = False
        for sheet in sheets:
            best = try_place(sheet["free_rects"], part, rotate_allowed)
            if best is None:
                continue
            rect = sheet["free_rects"].pop(best["idx"])
            sheet["placements"].append({
                "product_code": part["product_code"],
                "part_code": part["part_code"],
                "part_name": part["part_name"],
                "date": part["date"],
                "color": part["color"],
                "thickness_mm": part["thickness_mm"],
                "x_mm": round(rect["x"] + margin, 1),
                "y_mm": round(rect["y"] + margin, 1),
                "width_mm": round(best["w"], 1),
                "height_mm": round(best["h"], 1),
            })
            rw = rect["w"] - best["w"] - kerf
            bh = rect["h"] - best["h"] - kerf
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
            sheet["placements"].append({
                "product_code": part["product_code"],
                "part_code": part["part_code"],
                "part_name": part["part_name"],
                "date": part["date"],
                "color": part["color"],
                "thickness_mm": part["thickness_mm"],
                "x_mm": round(rect["x"] + margin, 1),
                "y_mm": round(rect["y"] + margin, 1),
                "width_mm": round(part["width_mm"], 1),
                "height_mm": round(part["height_mm"], 1),
            })
            rw = rect["w"] - part["width_mm"] - kerf
            bh = rect["h"] - part["height_mm"] - kerf
            if rw > 0:
                sheet["free_rects"].append({"x": rect["x"] + part["width_mm"] + kerf, "y": rect["y"], "w": rw, "h": part["height_mm"]})
            if bh > 0:
                sheet["free_rects"].append({"x": rect["x"], "y": rect["y"] + part["height_mm"] + kerf, "w": rect["w"], "h": bh})
    area = sum(p["width_mm"] * p["height_mm"] for s in sheets for p in s["placements"])
    used = len(sheets)
    yield_rate = round(area / (used * board_width * board_height) * 100, 2) if used else 0.0
    return {"board_width_mm": board_width, "board_height_mm": board_height, "used_boards": used, "yield_rate": yield_rate, "sheets": sheets}

def build_pattern_workorder(sheet):
    df = pd.DataFrame(sheet["placements"])
    if df.empty:
        return pd.DataFrame()
    out = (df.groupby(["product_code", "part_code", "part_name", "color", "thickness_mm", "date"])
           .agg(수량=("part_code", "count"), 가로=("width_mm", "first"), 세로=("height_mm", "first"))
           .reset_index())
    out["규격"] = out["가로"].astype(str) + "x" + out["세로"].astype(str)
    return out[["product_code", "part_code", "part_name", "color", "thickness_mm", "규격", "수량", "date"]]

def make_svg(sheet, board_width_mm, board_height_mm):
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width, svg_height = int(board_width_mm * scale), int(board_height_mm * scale)
    parts_svg = []
    for p in sheet["placements"]:
        x, y = p["x_mm"] * scale, p["y_mm"] * scale
        w, h = p["width_mm"] * scale, p["height_mm"] * scale
        label = f"{p['part_code']} / {p['part_name'][:12]} / t{p['thickness_mm']}"
        parts_svg.append(f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect><text x="{x+4}" y="{y+16}" font-size="11" fill="#111">{label}</text></g>')
    return f'<div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;"><svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>{"".join(parts_svg)}</svg></div>'

def export_grouped_workorders_excel(group_results):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for g in group_results:
            sheet_name = f"t{g['thickness_mm']}_{g['sheet']['sheet_no']}"[:31]
            build_pattern_workorder(g["sheet"]).to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()

st.title("목재 재단 프로그램 v31 Thickness Fix")
st.caption("두께별 분리 재단 고정 버전")

bom_url = st.text_input("BOM URL (기본 고정)", value=DEFAULT_BOM_URL)
bom_file = st.file_uploader("또는 BOM 엑셀 업로드", type=["xlsx", "xls"], key="bom")
plan_file = st.file_uploader("주차별 생산계획 업로드", type=["xlsx", "xls"], key="plan")

try:
    bom_df = None
    if bom_file is not None or normalize_text(bom_url):
        bom_df = load_bom(read_bom(bom_file, bom_url))
        st.success(f"BOM 로드 완료: {len(bom_df)}행")

    if bom_df is not None and plan_file is not None:
        plan_df, parse_log_df = parse_plan_workbook_auto(plan_file)
        with st.expander("생산계획 해석 로그", expanded=False):
            st.dataframe(parse_log_df, width="stretch", height=180)

        if plan_df.empty:
            st.warning("주차별 생산계획을 자동 해석하지 못했습니다.")
        else:
            if len(plan_df) > MAX_PLAN_ROWS:
                st.warning(f"생산계획이 커서 상위 {MAX_PLAN_ROWS}건만 처리합니다.")
                plan_df = plan_df.head(MAX_PLAN_ROWS)

            c1, c2, c3, c4 = st.columns(4)
            board_preset = c1.selectbox("원장 규격", list(BOARD_PRESETS.keys()), index=0)
            kerf = c2.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
            margin = c3.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
            rotate_allowed = c4.checkbox("회전 허용", value=True)

            selected_date = st.selectbox("작업 날짜 선택", sorted(plan_df["date"].unique().tolist()))
            date_plan = plan_df[plan_df["date"] == selected_date].copy()
            valid_products = sorted(date_plan["product_code"].unique().tolist())
            selected_product = st.selectbox("품목코드 선택", valid_products)

            work_plan = date_plan[date_plan["product_code"] == selected_product].copy()
            bw, bh = BOARD_PRESETS[board_preset]

            if work_plan.empty:
                st.warning("선택한 날짜의 해당 품목코드 계획이 없습니다.")
            else:
                row = work_plan.iloc[0]
                product_code = normalize_text(row["product_code"])
                color = normalize_text(row["color"])
                plan_qty = to_int(row["plan_qty"], 0)

                matched_all = bom_df[bom_df["product_code"] == product_code].copy()
                matched_color = matched_all[matched_all["color"] == color].copy() if color else matched_all.copy()
                matched_used = matched_color if not matched_color.empty else matched_all.copy()

                st.markdown("### BOM 매칭 현황")
                c1, c2, c3 = st.columns(3)
                c1.metric("BOM 전체 부품 수", len(matched_all))
                c2.metric("색상 일치 부품 수", len(matched_color))
                c3.metric("재단 대상 부품 수", len(matched_used[matched_used["is_cutting_target"] == True]))

                parts_df = matched_used[matched_used["is_cutting_target"] == True].copy()
                parts_df["qty"] = parts_df["bom_qty"].astype(int) * plan_qty
                parts_df["date"] = selected_date

                if parts_df.empty:
                    st.warning("재단 대상 부품이 없습니다.")
                else:
                    st.markdown("### 실제 패턴 계산에 들어간 부품")
                    st.dataframe(parts_df.head(MAX_MATCH_ROWS), width="stretch", height=240)

                    group_results = []
                    for thickness in sorted(parts_df["thickness_mm"].dropna().unique().tolist()):
                        sub = parts_df[parts_df["thickness_mm"] == thickness].copy()
                        result = optimize_parts_fast(sub, bw, bh, float(kerf), float(margin), rotate_allowed)
                        for s in result["sheets"]:
                            group_results.append({"thickness_mm": thickness, "board_width_mm": result["board_width_mm"], "board_height_mm": result["board_height_mm"], "sheet": s})

                    st.warning("두께가 다른 부품은 이제 같은 분할도에 들어가지 않도록 두께별로 강제 분리했습니다.")

                    summary_rows = []
                    for g in group_results:
                        summary_rows.append({"두께(mm)": g["thickness_mm"], "시트번호": g["sheet"]["sheet_no"], "배치부품수": len(g["sheet"]["placements"])})
                    st.dataframe(pd.DataFrame(summary_rows), width="stretch", height=220)

                    st.download_button("두께별 패턴 작업지시서 다운로드", data=export_grouped_workorders_excel(group_results), file_name="pattern_workorders_v31_thickness_fix.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

                    labels = [f"t{g['thickness_mm']} / Sheet {g['sheet']['sheet_no']}" for g in group_results]
                    if labels:
                        selected_label = st.selectbox("시트 선택", labels)
                        idx = labels.index(selected_label)
                        chosen = group_results[idx]
                        st.markdown(f"### 선택 분할도 - 두께 {chosen['thickness_mm']}mm")
                        st.dataframe(build_pattern_workorder(chosen["sheet"]), width="stretch", height=220)
                        components.html(make_svg(chosen["sheet"], chosen["board_width_mm"], chosen["board_height_mm"]), height=520, scrolling=True)
except Exception as e:
    st.error(f"앱 실행 오류: {type(e).__name__}: {e}")
    with st.expander("오류 상세", expanded=True):
        st.code(traceback.format_exc())
