import math
import re
import traceback
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Wood Cutting Optimizer", layout="wide")


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
    material = str(normalize_value(row.get("재질")) or "").strip()
    image_flag = str(normalize_value(row.get("대표이미지")) or "").strip().upper()
    qty = to_int(row.get("정소요량"), 0) or to_int(row.get("실소요량"), 0)

    if not width or not height or not thickness:
        return False
    if qty <= 0:
        return False
    if image_flag == "Y":
        return False
    if material and any(x in material for x in ["BOX", "포장", "철물", "경첩"]):
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

        item = {
            "row_no": int(idx) + 2,
            "product_code": product_code,
            "part_code": part_code,
            "part_name": part_name,
            "color": str(raw.get("색상") or "").strip(),
            "qty": max(1, qty),
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


def expand_parts(parts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    expanded: List[Dict[str, Any]] = []
    for p in parts:
        repeat = max(1, to_int(p.get("qty"), 1))
        for _ in range(repeat):
            expanded.append({
                "product_code": p["product_code"],
                "part_code": p["part_code"],
                "part_name": p["part_name"],
                "width_mm": float(p["width_mm"]),
                "height_mm": float(p["height_mm"]),
                "thickness_mm": float(p["thickness_mm"]),
                "material_name": p["material_name"],
            })
    expanded.sort(key=lambda x: x["width_mm"] * x["height_mm"], reverse=True)
    return expanded


def try_place_in_free_rect(part: Dict[str, Any], free_rect: Dict[str, float], kerf: float, rotate_allowed: bool):
    variants = [(part["width_mm"], part["height_mm"], False)]
    if rotate_allowed and part["width_mm"] != part["height_mm"]:
        variants.append((part["height_mm"], part["width_mm"], True))

    for w, h, rotated in variants:
        if w <= free_rect["w"] and h <= free_rect["h"]:
            placement = {
                "x_mm": round(free_rect["x"], 1),
                "y_mm": round(free_rect["y"], 1),
                "width_mm": round(w, 1),
                "height_mm": round(h, 1),
                "rotated": rotated,
                "part_code": part["part_code"],
                "part_name": part["part_name"],
            }

            right_w = free_rect["w"] - w - kerf
            bottom_h = free_rect["h"] - h - kerf

            right = {
                "x": free_rect["x"] + w + kerf,
                "y": free_rect["y"],
                "w": right_w,
                "h": h,
            }
            bottom = {
                "x": free_rect["x"],
                "y": free_rect["y"] + h + kerf,
                "w": free_rect["w"],
                "h": bottom_h,
            }
            bottom_right = {
                "x": free_rect["x"] + w + kerf,
                "y": free_rect["y"] + h + kerf,
                "w": right_w,
                "h": bottom_h,
            }

            new_rects = [r for r in (right, bottom, bottom_right) if r["w"] > 10 and r["h"] > 10]
            return placement, new_rects
    return None, None


def optimize_parts(parts: List[Dict[str, Any]], board_width: float, board_height: float, kerf: float, margin: float, rotate_allowed: bool):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    if usable_w <= 0 or usable_h <= 0:
        raise ValueError("원판 크기보다 여유치가 큽니다.")

    sheets = []
    unplaced = []

    for part in expand_parts(parts):
        placed = False

        for sheet in sheets:
            best_idx = None
            best_score = None
            best_result = None

            for idx, rect in enumerate(sheet["free_rects"]):
                placement, new_rects = try_place_in_free_rect(part, rect, kerf, rotate_allowed)
                if placement:
                    waste_score = rect["w"] * rect["h"] - placement["width_mm"] * placement["height_mm"]
                    if best_score is None or waste_score < best_score:
                        best_score = waste_score
                        best_idx = idx
                        best_result = (placement, new_rects)

            if best_result:
                placement, new_rects = best_result
                placement["x_mm"] = round(placement["x_mm"] + margin, 1)
                placement["y_mm"] = round(placement["y_mm"] + margin, 1)
                sheet["placements"].append(placement)
                sheet["free_rects"].pop(best_idx)
                sheet["free_rects"].extend(new_rects)
                placed = True
                break

        if not placed:
            new_sheet = {
                "sheet_no": len(sheets) + 1,
                "placements": [],
                "free_rects": [{"x": 0.0, "y": 0.0, "w": usable_w, "h": usable_h}],
            }
            placement, new_rects = try_place_in_free_rect(part, new_sheet["free_rects"][0], kerf, rotate_allowed)
            if placement:
                placement["x_mm"] = round(placement["x_mm"] + margin, 1)
                placement["y_mm"] = round(placement["y_mm"] + margin, 1)
                new_sheet["placements"].append(placement)
                new_sheet["free_rects"] = new_rects
                sheets.append(new_sheet)
            else:
                unplaced.append(part)

    total_part_area = sum(p["width_mm"] * p["height_mm"] for s in sheets for p in s["placements"])
    board_area = board_width * board_height
    used_boards = len(sheets)
    total_board_area = used_boards * board_area if used_boards else 0
    waste_area = max(0, total_board_area - total_part_area)
    yield_rate = round((total_part_area / total_board_area) * 100, 2) if total_board_area else 0

    return {
        "board_width_mm": round(board_width, 1),
        "board_height_mm": round(board_height, 1),
        "used_boards": used_boards,
        "total_part_area": round(total_part_area, 1),
        "waste_area": round(waste_area, 1),
        "yield_rate": yield_rate,
        "unplaced_count": len(unplaced),
        "unplaced_parts": unplaced,
        "sheets": [{"sheet_no": s["sheet_no"], "placements": s["placements"]} for s in sheets],
    }


def make_svg(sheet: Dict[str, Any], board_width_mm: float, board_height_mm: float, kerf: float = 0.0) -> str:
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale)
    svg_height = int(board_height_mm * scale)

    parts_svg = []
    cut_lines = []

    for p in sheet["placements"]:
        x = p["x_mm"] * scale
        y = p["y_mm"] * scale
        w = p["width_mm"] * scale
        h = p["height_mm"] * scale
        label = f'{p["part_code"]} ({p["width_mm"]}x{p["height_mm"]})'
        parts_svg.append(
            f"""
            <g>
                <rect x="{x}" y="{y}" width="{w}" height="{h}"
                      fill="#dbeafe" stroke="#1d4ed8" stroke-width="1.2"></rect>
                <text x="{x + 4}" y="{y + 16}" font-size="12" fill="#111">{label}</text>
            </g>
            """
        )
        if kerf > 0:
            k = kerf * scale
            cut_lines.append(f'<rect x="{x + w}" y="{y}" width="{k}" height="{h}" fill="#fca5a5" fill-opacity="0.65"></rect>')
            cut_lines.append(f'<rect x="{x}" y="{y + h}" width="{w}" height="{k}" fill="#fca5a5" fill-opacity="0.65"></rect>')

    return f"""
    <div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;">
      <svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg">
          <rect x="0" y="0" width="{svg_width}" height="{svg_height}"
                fill="white" stroke="#333" stroke-width="2"></rect>
          {''.join(cut_lines)}
          {''.join(parts_svg)}
      </svg>
    </div>
    """


st.title("목재 재단 프로그램")
st.caption("분할도 계산 보정 버전: 재단 수량 직접 수정, 톱날폭/여유치 소수점, 톱날선 표시")

with st.expander("지원 BOM 컬럼", expanded=False):
    st.markdown(\"\"\"
    - 품목코드
    - 부품코드
    - 품목명
    - 색상
    - 정소요량
    - 실소요량
    - 규격
    - 재질
    - 소요공정
    - 대표이미지
    \"\"\")

uploaded_file = st.file_uploader("BOM 엑셀 업로드", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        items, errors = load_bom_from_dataframe(df)
        st.session_state["bom_items"] = items
        st.session_state["bom_errors"] = errors
        st.success(f"업로드 완료: {len(items)}행, 오류 {len(errors)}건")
    except Exception as exc:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {exc}")

bom_items = st.session_state.get("bom_items", [])
bom_errors = st.session_state.get("bom_errors", [])

if bom_items:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("업로드 행 수", len(bom_items))
    with col2:
        st.metric("제품 수", len(sorted({x['product_code'] for x in bom_items if x['product_code']})))
    with col3:
        st.metric("재단 대상 행 수", sum(1 for x in bom_items if x["is_cutting_target"]))

    if bom_errors:
        with st.expander(f"검증 오류 {len(bom_errors)}건", expanded=False):
            st.dataframe(pd.DataFrame(bom_errors), use_container_width=True)

    products = sorted({x["product_code"] for x in bom_items if x["product_code"]})
    if not products:
        st.warning("품목코드가 있는 데이터가 없습니다.")
    else:
        st.subheader("제품 조회")
        selected_product = st.selectbox("제품코드 선택", products)

        product_items = [x for x in bom_items if x["product_code"] == selected_product]
        cutting_items = [
            x for x in product_items
            if x["is_cutting_target"] and x.get("width_mm") and x.get("height_mm") and x.get("thickness_mm")
        ]

        left, right = st.columns([1.2, 1.3])

        with left:
            st.markdown("### BOM 목록")
            show_cutting_only = st.checkbox("재단 대상만 보기", value=True)
            base_items = cutting_items if show_cutting_only else product_items
            editable_items = []

            for item in base_items:
                row = dict(item)
                row["qty"] = max(1, to_int(row.get("qty"), 1))
                editable_items.append(row)

            edited_df = st.data_editor(
                pd.DataFrame(editable_items),
                use_container_width=True,
                height=460,
                disabled=[
                    "row_no", "product_code", "part_code", "part_name", "color", "spec_raw",
                    "width_mm", "height_mm", "thickness_mm", "material_name", "process_name",
                    "is_cutting_target"
                ],
                column_config={
                    "qty": st.column_config.NumberColumn("재단 수량", min_value=1, step=1, format="%d"),
                    "width_mm": st.column_config.NumberColumn("가로(mm)", format="%.1f"),
                    "height_mm": st.column_config.NumberColumn("세로(mm)", format="%.1f"),
                    "thickness_mm": st.column_config.NumberColumn("두께(mm)", format="%.1f"),
                }
            )

        with right:
            st.markdown("### 최적화 조건")
            c1, c2 = st.columns(2)
            with c1:
                board_width = st.number_input("원판 가로(mm)", min_value=100.0, value=2440.0, step=1.0, format="%.1f")
                kerf = st.number_input("톱날폭(mm)", min_value=0.0, value=3.0, step=0.1, format="%.1f")
            with c2:
                board_height = st.number_input("원판 세로(mm)", min_value=100.0, value=1220.0, step=1.0, format="%.1f")
                margin = st.number_input("여유치(mm)", min_value=0.0, value=20.0, step=0.1, format="%.1f")

            rotate_allowed = st.checkbox("회전 허용", value=True)

            if st.button("최적화 실행", type="primary", use_container_width=True):
                try:
                    optimized_items = edited_df.to_dict("records")
                    optimized_items = [x for x in optimized_items if bool(x.get("is_cutting_target")) and x.get("width_mm") and x.get("height_mm")]

                    for item in optimized_items:
                        item["qty"] = max(1, to_int(item.get("qty"), 1))
                        item["width_mm"] = to_float(item.get("width_mm"), 0.0)
                        item["height_mm"] = to_float(item.get("height_mm"), 0.0)
                        item["thickness_mm"] = to_float(item.get("thickness_mm"), 0.0)

                    if not optimized_items:
                        st.error("최적화 가능한 재단 대상 부품이 없습니다.")
                    else:
                        result = optimize_parts(
                            optimized_items,
                            float(board_width),
                            float(board_height),
                            float(kerf),
                            float(margin),
                            rotate_allowed,
                        )
                        result["kerf_mm"] = float(kerf)
                        st.session_state["opt_result"] = result
                        st.session_state["opt_product"] = selected_product
                except Exception as exc:
                    st.error(f"최적화 중 오류: {exc}")
                    st.code(traceback.format_exc())

            opt_result = st.session_state.get("opt_result")
            opt_product = st.session_state.get("opt_product")

            if opt_result and opt_product == selected_product:
                st.markdown("### 최적화 결과")
                r1, r2, r3, r4 = st.columns(4)
                with r1:
                    st.metric("사용 원판 수", opt_result["used_boards"])
                with r2:
                    st.metric("수율", f"{opt_result['yield_rate']}%")
                with r3:
                    st.metric("자투리 면적", f"{opt_result['waste_area']:,}")
                with r4:
                    st.metric("미배치 수", opt_result["unplaced_count"])

                if opt_result["sheets"]:
                    labels = [f"Sheet {s['sheet_no']}" for s in opt_result["sheets"]]
                    selected_sheet_label = st.selectbox("시트 선택", labels)
                    selected_sheet_no = int(selected_sheet_label.replace("Sheet ", ""))
                    selected_sheet = next(s for s in opt_result["sheets"] if s["sheet_no"] == selected_sheet_no)

                    st.caption("빨간색은 톱날폭(kerf) 영역입니다.")
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

                if opt_result["unplaced_parts"]:
                    with st.expander("미배치 부품", expanded=False):
                        st.dataframe(pd.DataFrame(opt_result["unplaced_parts"]), use_container_width=True)

else:
    st.info("BOM 엑셀 파일을 업로드하면 제품 조회와 최적화를 사용할 수 있습니다.")
