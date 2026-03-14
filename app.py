import math
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(page_title="Wood Cutting Optimizer", layout="wide")


def parse_spec(spec_raw: Any) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    if spec_raw is None or (isinstance(spec_raw, float) and math.isnan(spec_raw)):
        return None, None, None
    text = str(spec_raw).strip()
    match = re.match(r"^\s*(\d+)\s*[*xX]\s*(\d+)\s*[*xX]\s*(\d+)\s*$", text)
    if not match:
        return None, None, None
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


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


def is_cutting_target(row: Dict[str, Any], width: Optional[int], height: Optional[int], thickness: Optional[int]) -> bool:
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
            "qty": qty,
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
        repeat = max(1, int(p["qty"]))
        for _ in range(repeat):
            expanded.append({
                "product_code": p["product_code"],
                "part_code": p["part_code"],
                "part_name": p["part_name"],
                "width_mm": p["width_mm"],
                "height_mm": p["height_mm"],
                "thickness_mm": p["thickness_mm"],
                "material_name": p["material_name"],
            })
    expanded.sort(key=lambda x: x["width_mm"] * x["height_mm"], reverse=True)
    return expanded


def try_place_in_free_rect(part: Dict[str, Any], free_rect: Dict[str, int], kerf: int, rotate_allowed: bool):
    variants = [(part["width_mm"], part["height_mm"], False)]
    if rotate_allowed and part["width_mm"] != part["height_mm"]:
        variants.append((part["height_mm"], part["width_mm"], True))

    for w, h, rotated in variants:
        need_w = w + kerf
        need_h = h + kerf
        if need_w <= free_rect["w"] and need_h <= free_rect["h"]:
            placement = {
                "x_mm": free_rect["x"],
                "y_mm": free_rect["y"],
                "width_mm": w,
                "height_mm": h,
                "rotated": rotated,
                "part_code": part["part_code"],
                "part_name": part["part_name"],
            }
            right = {
                "x": free_rect["x"] + need_w,
                "y": free_rect["y"],
                "w": free_rect["w"] - need_w,
                "h": h,
            }
            bottom = {
                "x": free_rect["x"],
                "y": free_rect["y"] + need_h,
                "w": free_rect["w"],
                "h": free_rect["h"] - need_h,
            }
            new_rects = [r for r in (right, bottom) if r["w"] > 20 and r["h"] > 20]
            return placement, new_rects
    return None, None


def optimize_parts(parts: List[Dict[str, Any]], board_width: int, board_height: int, kerf: int, margin: int, rotate_allowed: bool):
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
                placement["x_mm"] += margin
                placement["y_mm"] += margin
                sheet["placements"].append(placement)
                sheet["free_rects"].pop(best_idx)
                sheet["free_rects"].extend(new_rects)
                placed = True
                break

        if not placed:
            new_sheet = {
                "sheet_no": len(sheets) + 1,
                "placements": [],
                "free_rects": [{"x": 0, "y": 0, "w": usable_w, "h": usable_h}],
            }
            placement, new_rects = try_place_in_free_rect(part, new_sheet["free_rects"][0], kerf, rotate_allowed)
            if placement:
                placement["x_mm"] += margin
                placement["y_mm"] += margin
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
        "board_width_mm": board_width,
        "board_height_mm": board_height,
        "used_boards": used_boards,
        "total_part_area": total_part_area,
        "waste_area": waste_area,
        "yield_rate": yield_rate,
        "unplaced_count": len(unplaced),
        "unplaced_parts": unplaced,
        "sheets": [{"sheet_no": s["sheet_no"], "placements": s["placements"]} for s in sheets],
    }


def make_svg(sheet: Dict[str, Any], board_width_mm: int, board_height_mm: int) -> str:
    scale = min(900 / board_width_mm, 600 / board_height_mm)
    svg_width = int(board_width_mm * scale)
    svg_height = int(board_height_mm * scale)

    parts_svg = []
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

    return f"""
    <div style="overflow:auto; border:1px solid #ddd; padding:12px; background:#fff;">
      <svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg">
          <rect x="0" y="0" width="{svg_width}" height="{svg_height}"
                fill="white" stroke="#333" stroke-width="2"></rect>
          {''.join(parts_svg)}
      </svg>
    </div>
    """


st.title("목재 재단 프로그램")
st.caption("Streamlit Cloud에 바로 올릴 수 있는 단일 파일 버전입니다.")

with st.expander("지원 BOM 컬럼", expanded=False):
    st.markdown("""
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
    """)

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
        st.metric("제품 수", len(sorted({x["product_code"] for x in bom_items if x["product_code"]})))
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
        cutting_items = [x for x in product_items if x["is_cutting_target"]]

        left, right = st.columns([1.1, 1.4])

        with left:
            st.markdown("### BOM 목록")
            show_cutting_only = st.checkbox("재단 대상만 보기", value=True)
            view_items = cutting_items if show_cutting_only else product_items
            st.dataframe(pd.DataFrame(view_items), use_container_width=True, height=420)

        with right:
            st.markdown("### 최적화 조건")
            c1, c2 = st.columns(2)
            with c1:
                board_width = st.number_input("원판 가로(mm)", min_value=100, value=2440, step=10)
                kerf = st.number_input("톱날폭(mm)", min_value=0, value=3, step=1)
            with c2:
                board_height = st.number_input("원판 세로(mm)", min_value=100, value=1220, step=10)
                margin = st.number_input("여유치(mm)", min_value=0, value=5, step=1)

            rotate_allowed = st.checkbox("회전 허용", value=True)

            if st.button("최적화 실행", type="primary", use_container_width=True):
                if not cutting_items:
                    st.error("최적화 가능한 재단 대상 부품이 없습니다.")
                else:
                    try:
                        result = optimize_parts(
                            cutting_items,
                            int(board_width),
                            int(board_height),
                            int(kerf),
                            int(margin),
                            rotate_allowed,
                        )
                        st.session_state["opt_result"] = result
                        st.session_state["opt_product"] = selected_product
                    except Exception as exc:
                        st.error(str(exc))

            opt_result = st.session_state.get("opt_result")
            opt_product = st.session_state.get("opt_product")

            if opt_result and opt_product == selected_product:
                st.markdown("### 최적화 결과")
                r1, r2, r3, r4 = st.columns(4)
                with r1:
                    st.metric("사용 원판 수", opt_result["used_boards"])
                with r2:
                    st.metric("수율", f'{opt_result["yield_rate"]}%')
                with r3:
                    st.metric("자투리 면적", f'{opt_result["waste_area"]:,}')
                with r4:
                    st.metric("미배치 수", opt_result["unplaced_count"])

                if opt_result["sheets"]:
                    labels = [f'Sheet {s["sheet_no"]}' for s in opt_result["sheets"]]
                    selected_sheet_label = st.selectbox("시트 선택", labels)
                    selected_sheet_no = int(selected_sheet_label.replace("Sheet ", ""))
                    selected_sheet = next(s for s in opt_result["sheets"] if s["sheet_no"] == selected_sheet_no)

                    components.html(
                        make_svg(selected_sheet, opt_result["board_width_mm"], opt_result["board_height_mm"]),
                        height=700,
                        scrolling=True,
                    )

                    st.dataframe(pd.DataFrame(selected_sheet["placements"]), use_container_width=True, height=260)

                if opt_result["unplaced_parts"]:
                    with st.expander("미배치 부품", expanded=False):
                        st.dataframe(pd.DataFrame(opt_result["unplaced_parts"]), use_container_width=True)

else:
    st.info("왼쪽 위에서 BOM 엑셀 파일을 업로드하면 제품 조회와 최적화를 사용할 수 있습니다.")
