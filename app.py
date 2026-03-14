import math
import re
import traceback
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="목재 재단 프로그램 Fast v4", layout="wide")

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
    if qty <= 0:
        return False
    if image_flag == "Y":
        return False
    exclude_words = ["BOX", "포장", "철물", "경첩"]
    if any(word in material for word in exclude_words):
        return False
    return True


def load_bom_from_dataframe(df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    items: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []

    for idx, row in df.iterrows():
        raw = row.to_dict()
        width, height, thickness = parse_spec(raw.get("규격"))
        product_code = str(raw.get("품목코드") or "").strip()
        part_code = str(raw.get("부품코드") or "").strip()
        part_name = str(raw.get("품목명") or "").strip()
        qty = to_int(raw.get("정소요량"), 0) or to_int(raw.get("실소요량"), 0)
        color = str(raw.get("색상") or "").strip()
        material_name = str(raw.get("재질") or "").strip()

        item = {
            "selected": True,
            "row_no": int(idx) + 2,
            "product_code": product_code,
            "part_code": part_code,
            "part_name": part_name,
            "color": color,
            "qty": max(1, qty),
            "spec_raw": str(raw.get("규격") or "").strip(),
            "width_mm": width,
            "height_mm": height,
            "thickness_mm": thickness,
            "material_name": material_name,
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


def apply_cut_count(parts: List[Dict[str, Any]], cut_count: int) -> List[Dict[str, Any]]:
    adjusted: List[Dict[str, Any]] = []
    for p in parts:
        row = dict(p)
        row["qty"] = max(1, to_int(row.get("qty"), 1)) * cut_count
        adjusted.append(row)
    return adjusted


def expand_parts(parts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    expanded: List[Dict[str, Any]] = []
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


def prune_rects(rects: List[Dict[str, float]]) -> List[Dict[str, float]]:
    kept: List[Dict[str, float]] = []
    for i, r in enumerate(rects):
        if r["w"] <= 0 or r["h"] <= 0:
            continue
        contained = False
        for j, other in enumerate(rects):
            if i == j:
                continue
            if (
                r["x"] >= other["x"] - 1e-9
                and r["y"] >= other["y"] - 1e-9
                and r["x"] + r["w"] <= other["x"] + other["w"] + 1e-9
                and r["y"] + r["h"] <= other["y"] + other["h"] + 1e-9
            ):
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


def split_rect_guillotine(rect: Dict[str, float], placed_w: float, placed_h: float, kerf: float) -> List[Dict[str, float]]:
    right_w = rect["w"] - placed_w - kerf
    bottom_h = rect["h"] - placed_h - kerf
    new_rects: List[Dict[str, float]] = []

    if right_w > 0:
        new_rects.append({
            "x": rect["x"] + placed_w + kerf,
            "y": rect["y"],
            "w": right_w,
            "h": placed_h,
        })

    if bottom_h > 0:
        new_rects.append({
            "x": rect["x"],
            "y": rect["y"] + placed_h + kerf,
            "w": rect["w"],
            "h": bottom_h,
        })

    return prune_rects(new_rects)


def try_place_part(free_rects: List[Dict[str, float]], part: Dict[str, Any], rotate_allowed: bool):
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
                    best = {
                        "score": score,
                        "rect_index": idx,
                        "w": w,
                        "h": h,
                        "rotated": rotated,
                    }
    return best


def optimize_group(parts: List[Dict[str, Any]], board_width: float, board_height: float, kerf: float, margin: float, rotate_allowed: bool, group_name: str):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    if usable_w <= 0 or usable_h <= 0:
        raise ValueError("원장 크기보다 여유치가 큽니다.")

    sheets: List[Dict[str, Any]] = []
    unplaced: List[Dict[str, Any]] = []

    for part in expand_parts(parts):
        placed = False

        for sheet in sheets:
            best = try_place_part(sheet["free_rects"], part, rotate_allowed)
            if best is None:
                continue

            rect = sheet["free_rects"].pop(best["rect_index"])
            placement = {
                "x_mm": round(rect["x"] + margin, 1),
                "y_mm": round(rect["y"] + margin, 1),
                "width_mm": round(best["w"], 1),
                "height_mm": round(best["h"], 1),
                "rotated": best["rotated"],
                "part_code": part["part_code"],
                "part_name": part["part_name"],
                "product_code": part["product_code"],
                "color": part["color"],
                "thickness_mm": round(part["thickness_mm"], 1),
                "material_name": part["material_name"],
            }
            sheet["placements"].append(placement)
            sheet["free_rects"].extend(split_rect_guillotine(rect, best["w"], best["h"], kerf))
            sheet["free_rects"] = prune_rects(sheet["free_rects"])
            placed = True
            break

        if not placed:
            new_sheet = {
                "sheet_no": len(sheets) + 1,
                "group_name": group_name,
                "placements": [],
                "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}],
            }
            best = try_place_part(new_sheet["free_rects"], part, rotate_allowed)
            if best is None:
                unplaced.append(part)
            else:
                rect = new_sheet["free_rects"].pop(best["rect_index"])
                placement = {
                    "x_mm": round(rect["x"] + margin, 1),
                    "y_mm": round(rect["y"] + margin, 1),
                    "width_mm": round(best["w"], 1),
                    "height_mm": round(best["h"], 1),
                    "rotated": best["rotated"],
                    "part_code": part["part_code"],
                    "part_name": part["part_name"],
                    "product_code": part["product_code"],
                    "color": part["color"],
                    "thickness_mm": round(part["thickness_mm"], 1),
                    "material_name": part["material_name"],
                }
                new_sheet["placements"].append(placement)
                new_sheet["free_rects"].extend(split_rect_guillotine(rect, best["w"], best["h"], kerf))
                new_sheet["free_rects"] = prune_rects(new_sheet["free_rects"])
                sheets.append(new_sheet)

    return sheets, unplaced


def build_groups(parts: List[Dict[str, Any]], mix_same_color_thickness: bool):
    groups: Dict[Tuple, List[Dict[str, Any]]] = {}
    if mix_same_color_thickness:
        for p in parts:
            key = (p["color"], float(p["thickness_mm"]))
            groups.setdefault(key, []).append(p)
    else:
        for p in parts:
            key = (p["product_code"], p["color"], float(p["thickness_mm"]))
            groups.setdefault(key, []).append(p)
    return groups


def optimize_parts(parts: List[Dict[str, Any]], board_width: float, board_height: float, kerf: float, margin: float, rotate_allowed: bool, mix_same_color_thickness: bool):
    groups = build_groups(parts, mix_same_color_thickness)
    all_sheets: List[Dict[str, Any]] = []
    all_unplaced: List[Dict[str, Any]] = []

    for key, group_parts in groups.items():
        if mix_same_color_thickness:
            group_name = f"색상:{key[0]} / 두께:{key[1]}"
        else:
            group_name = f"제품:{key[0]} / 색상:{key[1]} / 두께:{key[2]}"

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

    return {
        "board_width_mm": round(board_width, 1),
        "board_height_mm": round(board_height, 1),
        "used_boards": used_boards,
        "total_part_area": round(total_part_area, 1),
        "waste_area": round(waste_area, 1),
        "yield_rate": yield_rate,
        "unplaced_count": len(all_unplaced),
        "unplaced_parts": all_unplaced,
        "sheets": all_sheets,
    }


def analyze_alternatives(parts: List[Dict[str, Any]], current_board_w: float, current_board_h: float, kerf: float, margin: float, rotate_allowed: bool, mix_same_color_thickness: bool):
    scenarios = [
        ("현재 조건", current_board_w, current_board_h, margin),
        ("여유치 5.0", current_board_w, current_board_h, 5.0),
        ("여유치 8.0", current_board_w, current_board_h, 8.0),
        ("여유치 10.0", current_board_w, current_board_h, 10.0),
        ("원장 4x6 / 여유치 10.0", 1830.0, 1220.0, 10.0),
    ]
    rows = []
    for name, bw, bh, mg in scenarios:
        try:
            result = optimize_parts(parts, bw, bh, kerf, mg, rotate_allowed, mix_same_color_thickness)
            rows.append({
                "시나리오": name,
                "원장 가로": bw,
                "원장 세로": bh,
                "여유치": mg,
                "사용 원장 수": result["used_boards"],
                "수율(%)": result["yield_rate"],
                "자투리 면적": result["waste_area"],
            })
        except Exception:
            rows.append({
                "시나리오": name,
                "원장 가로": bw,
                "원장 세로": bh,
                "여유치": mg,
                "사용 원장 수": None,
                "수율(%)": None,
                "자투리 면적": None,
            })
    df = pd.DataFrame(rows)
    valid = df.dropna(subset=["수율(%)"])
    summary = ""
    if not valid.empty:
        best = valid.sort_values(["수율(%)", "사용 원장 수"], ascending=[False, True]).iloc[0]
        summary = (
            f"분석 결과, '{best['시나리오']}' 조건이 가장 높은 수율을 보였습니다. "
            f"예상 수율은 {best['수율(%)']}%이며, 사용 원장 수는 {int(best['사용 원장 수'])}장입니다."
        )
    return df, summary


def make_svg(sheet: Dict[str, Any], board_width_mm: float, board_height_mm: float, kerf: float = 0.0) -> str:
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale)
    svg_height = int(board_height_mm * scale)

    parts_svg: List[str] = []
    kerf_svg: List[str] = []

    for p in sheet["placements"]:
        x = p["x_mm"] * scale
        y = p["y_mm"] * scale
        w = p["width_mm"] * scale
        h = p["height_mm"] * scale
        label = f'{p["part_code"]} ({p["width_mm"]}x{p["height_mm"]})'

        parts_svg.append(
            f"""
            <g>
                <rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect>
                <text x="{x + 4}" y="{y + 16}" font-size="12" fill="#111">{label}</text>
            </g>
            """
        )

        if kerf > 0:
            k = kerf * scale
            kerf_svg.append(f'<rect x="{x + w}" y="{y}" width="{k}" height="{h}" fill="#fca5a5" fill-opacity="0.35"></rect>')
            kerf_svg.append(f'<rect x="{x}" y="{y + h}" width="{w}" height="{k}" fill="#fca5a5" fill-opacity="0.35"></rect>')

    return f"""
    <div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;">
      <svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg">
          <rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="white" stroke="#333" stroke-width="2"></rect>
          {''.join(kerf_svg)}
          {''.join(parts_svg)}
      </svg>
    </div>
    """


st.title("목재 재단 프로그램 Fast v4")
st.caption("품목 다중 선택 + 원장 매수/재단 매수 포함 + 분석 제안 포함")

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
bom_errors = st.session_state.get("bom_errors", [])

if bom_items:
    c1, c2, c3 = st.columns(3)
    c1.metric("업로드 행 수", len(bom_items))
    c2.metric("제품 수", len(sorted({x["product_code"] for x in bom_items if x["product_code"]})))
    c3.metric("재단 대상 행 수", sum(1 for x in bom_items if x["is_cutting_target"]))

    if bom_errors:
        with st.expander(f"검증 오류 {len(bom_errors)}건", expanded=False):
            st.dataframe(pd.DataFrame(bom_errors), use_container_width=True)

    products = sorted({x["product_code"] for x in bom_items if x["product_code"]})
    query = st.text_input("제품코드 입력", value=st.session_state.get("product_query", ""))
    st.session_state["product_query"] = query

    matched_products = [p for p in products if query.strip().upper() in p.upper()] if query.strip() else products[:50]
    selected_products = st.multiselect("조회 결과 / 재단 대상 제품 선택", matched_products, default=matched_products[:1] if matched_products else [])

    if selected_products:
        target_items = [
            x for x in bom_items
            if x["product_code"] in selected_products and x["is_cutting_target"]
            and x.get("width_mm") and x.get("height_mm") and x.get("thickness_mm")
        ]

        left, right = st.columns([1.2, 1.25])

        with left:
            st.subheader("BOM 목록")
            mix_same_color_thickness = st.checkbox("같은 색상 + 같은 두께 혼합 재단", value=False)
            st.caption("selected 체크된 품목만 재단합니다.")

            editable_rows = []
            for item in target_items:
                row = dict(item)
                row["selected"] = True
                row["qty"] = max(1, to_int(row.get("qty"), 1))
                editable_rows.append(row)

            edited_df = st.data_editor(
                pd.DataFrame(editable_rows),
                use_container_width=True,
                height=500,
                disabled=[
                    "row_no", "product_code", "part_code", "part_name", "color", "spec_raw",
                    "width_mm", "height_mm", "thickness_mm", "material_name", "process_name", "is_cutting_target"
                ],
                column_config={
                    "selected": st.column_config.CheckboxColumn("선택", default=True),
                    "qty": st.column_config.NumberColumn("재단 수량", min_value=1, step=1, format="%d"),
                    "width_mm": st.column_config.NumberColumn("가로(mm)", format="%.1f"),
                    "height_mm": st.column_config.NumberColumn("세로(mm)", format="%.1f"),
                    "thickness_mm": st.column_config.NumberColumn("두께(mm)", format="%.1f"),
                },
            )

        with right:
            st.subheader("최적화 조건")
            preset = st.selectbox("원장 규격 선택", list(BOARD_PRESETS.keys()), index=0)

            if BOARD_PRESETS[preset] is None:
                preset_w, preset_h = 2440.0, 1220.0
            else:
                preset_w, preset_h = BOARD_PRESETS[preset]

            r1, r2 = st.columns(2)
            with r1:
                board_width = st.number_input("원장 가로(mm)", min_value=100.0, value=float(preset_w), step=1.0, format="%.1f")
                kerf = st.number_input("톱날폭(mm)", min_value=0.0, value=4.8, step=0.1, format="%.1f")
                board_batch = st.number_input("한 번에 재단할 원장 매수", min_value=1, value=1, step=1)
            with r2:
                board_height = st.number_input("원장 세로(mm)", min_value=100.0, value=float(preset_h), step=1.0, format="%.1f")
                margin = st.number_input("여유치(mm)", min_value=0.0, value=10.0, step=0.1, format="%.1f")
                cut_count = st.number_input("재단 매수", min_value=1, value=1, step=1)

            rotate_allowed = st.checkbox("회전 허용", value=True)

            if st.button("최적화 실행", type="primary", use_container_width=True):
                try:
                    optimized_items = edited_df.to_dict("records")
                    optimized_items = [
                        x for x in optimized_items
                        if bool(x.get("selected")) and x.get("width_mm") and x.get("height_mm") and x.get("thickness_mm")
                    ]

                    for item in optimized_items:
                        item["qty"] = max(1, to_int(item.get("qty"), 1))
                        item["width_mm"] = to_float(item.get("width_mm"), 0.0)
                        item["height_mm"] = to_float(item.get("height_mm"), 0.0)
                        item["thickness_mm"] = to_float(item.get("thickness_mm"), 0.0)

                    optimized_items = apply_cut_count(optimized_items, int(cut_count))
                    total_qty = sum(item["qty"] for item in optimized_items)

                    if not optimized_items:
                        st.error("선택된 재단 품목이 없습니다.")
                    elif total_qty > 3000:
                        st.error("총 재단 수량이 너무 많습니다. 수량을 줄여주세요.")
                    else:
                        result = optimize_parts(
                            optimized_items,
                            float(board_width),
                            float(board_height),
                            float(kerf),
                            float(margin),
                            rotate_allowed,
                            mix_same_color_thickness,
                        )
                        result["kerf_mm"] = float(kerf)
                        result["board_batch"] = int(board_batch)
                        result["cut_count"] = int(cut_count)
                        result["runs_required"] = math.ceil(result["used_boards"] / max(1, int(board_batch)))
                        st.session_state["opt_result"] = result

                        analysis_df, analysis_summary = analyze_alternatives(
                            optimized_items,
                            float(board_width),
                            float(board_height),
                            float(kerf),
                            float(margin),
                            rotate_allowed,
                            mix_same_color_thickness,
                        )
                        st.session_state["analysis_df"] = analysis_df
                        st.session_state["analysis_summary"] = analysis_summary

                except Exception as exc:
                    st.error(f"최적화 중 오류: {exc}")
                    st.code(traceback.format_exc())

            opt_result = st.session_state.get("opt_result")
            if opt_result:
                st.subheader("최적화 결과")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("사용 원장 수", opt_result["used_boards"])
                m2.metric("수율", f"{opt_result['yield_rate']}%")
                m3.metric("자투리 면적", f"{opt_result['waste_area']:,}")
                m4.metric("미배치 수", opt_result["unplaced_count"])

                x1, x2, x3 = st.columns(3)
                x1.metric("한 번에 재단할 원장 매수", opt_result.get("board_batch", 1))
                x2.metric("재단 매수", opt_result.get("cut_count", 1))
                x3.metric("예상 재단 횟수", opt_result.get("runs_required", 1))

                if opt_result["sheets"]:
                    labels = [f"Sheet {s['sheet_no']} | {s.get('group_name', '')}" for s in opt_result["sheets"]]
                    selected_sheet_label = st.selectbox("시트 선택", labels)
                    selected_sheet_no = int(selected_sheet_label.split("|")[0].replace("Sheet", "").strip())
                    selected_sheet = next(s for s in opt_result["sheets"] if s["sheet_no"] == selected_sheet_no)

                    st.caption("빨간색은 톱날폭(kerf) 영역입니다.")
                    st.write(f"그룹: {selected_sheet.get('group_name', '-')}")
                    components.html(
                        make_svg(
                            selected_sheet,
                            opt_result["board_width_mm"],
                            opt_result["board_height_mm"],
                            opt_result.get("kerf_mm", 0.0),
                        ),
                        height=700,
                        scrolling=True,
                    )
                    st.dataframe(pd.DataFrame(selected_sheet["placements"]), use_container_width=True, height=260)

                analysis_df = st.session_state.get("analysis_df")
                analysis_summary = st.session_state.get("analysis_summary", "")
                if analysis_df is not None:
                    st.subheader("수율 개선 분석")
                    if analysis_summary:
                        st.info(analysis_summary)
                    st.dataframe(analysis_df, use_container_width=True)
                    st.caption("예: 여유치를 변경하거나, 원장을 4x6(1220 x 1830)로 바꿨을 때의 수율 비교")
    else:
        st.info("제품코드를 입력하거나 선택하세요.")
else:
    st.info("BOM 엑셀 파일을 업로드하면 제품 조회와 최적화를 사용할 수 있습니다.")
