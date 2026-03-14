from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse
import pandas as pd
import math
import re
from typing import List, Dict, Any, Optional

app = FastAPI(title="Wood Cutting Optimizer", version="1.0.0")

# In-memory storage for a simple GitHub-ready single-file app
BOM_ITEMS: List[Dict[str, Any]] = []


HTML_PAGE = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Wood Cutting Optimizer</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 24px; background: #f7f7f7; color: #222; }
        .wrap { max-width: 1100px; margin: 0 auto; }
        .card { background: white; border-radius: 12px; padding: 18px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
        h1, h2 { margin-top: 0; }
        input, button, select { padding: 10px 12px; margin: 4px 0; }
        input[type="text"], input[type="number"] { width: 240px; }
        button { cursor: pointer; border: 1px solid #ccc; border-radius: 8px; background: #fff; }
        button:hover { background: #f0f0f0; }
        .row { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; }
        .muted { color: #666; font-size: 14px; }
        #result, #products, #bom { white-space: pre-wrap; }
        svg { background: white; border: 1px solid #ddd; }
        table { border-collapse: collapse; width: 100%; margin-top: 12px; font-size: 14px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .pill { display: inline-block; padding: 4px 8px; border-radius: 999px; background: #eef4ff; margin-right: 6px; }
    </style>
</head>
<body>
<div class="wrap">
    <div class="card">
        <h1>목재 재단 프로그램</h1>
        <div class="muted">단일 파일 FastAPI 앱입니다. 엑셀 BOM 업로드, 제품 조회, 최적화, 분할도 확인이 가능합니다.</div>
    </div>

    <div class="card">
        <h2>1. BOM 업로드</h2>
        <div class="row">
            <input type="file" id="fileInput" accept=".xlsx,.xls" />
            <button onclick="uploadBom()">업로드</button>
        </div>
        <div id="uploadStatus" class="muted"></div>
    </div>

    <div class="card">
        <h2>2. 제품 조회</h2>
        <div class="row">
            <button onclick="loadProducts()">제품 목록 불러오기</button>
        </div>
        <div id="products"></div>
    </div>

    <div class="card">
        <h2>3. 제품 BOM 조회</h2>
        <div class="row">
            <input type="text" id="productCode" placeholder="제품코드 입력" />
            <button onclick="loadBom()">BOM 조회</button>
        </div>
        <div id="bom"></div>
    </div>

    <div class="card">
        <h2>4. 재단 최적화</h2>
        <div class="row">
            <input type="text" id="optProductCode" placeholder="제품코드 입력" />
            <input type="number" id="boardWidth" value="2440" placeholder="원판 가로" />
            <input type="number" id="boardHeight" value="1220" placeholder="원판 세로" />
            <input type="number" id="kerf" value="3" placeholder="톱날폭" />
            <input type="number" id="margin" value="5" placeholder="여유치" />
            <label><input type="checkbox" id="rotateAllowed" checked /> 회전 허용</label>
            <button onclick="optimize()">최적화 실행</button>
        </div>
        <div id="result"></div>
        <div id="svgWrap"></div>
    </div>
</div>

<script>
async function uploadBom() {
    const fileInput = document.getElementById("fileInput");
    if (!fileInput.files.length) {
        alert("엑셀 파일을 선택하세요.");
        return;
    }
    const fd = new FormData();
    fd.append("file", fileInput.files[0]);

    const res = await fetch("/upload-bom", { method: "POST", body: fd });
    const data = await res.json();
    document.getElementById("uploadStatus").textContent = JSON.stringify(data, null, 2);
}

async function loadProducts() {
    const res = await fetch("/products");
    const data = await res.json();
    document.getElementById("products").innerHTML =
        (data.products || []).map(p => `<span class="pill">${p}</span>`).join(" ");
}

async function loadBom() {
    const code = document.getElementById("productCode").value.trim();
    if (!code) return alert("제품코드를 입력하세요.");
    const res = await fetch(`/bom/${encodeURIComponent(code)}`);
    const data = await res.json();
    document.getElementById("bom").textContent = JSON.stringify(data, null, 2);
}

function renderSvg(result) {
    const scale = 0.28;
    let html = "";
    for (const sheet of result.sheets) {
        let parts = "";
        for (const p of sheet.placements) {
            parts += `
                <rect x="${p.x_mm * scale}" y="${p.y_mm * scale}" width="${p.width_mm * scale}" height="${p.height_mm * scale}"
                      fill="#dbeafe" stroke="#1e40af" stroke-width="1"></rect>
                <text x="${p.x_mm * scale + 4}" y="${p.y_mm * scale + 16}" font-size="12" fill="#111">${p.part_code}</text>
                <text x="${p.x_mm * scale + 4}" y="${p.y_mm * scale + 30}" font-size="10" fill="#444">${p.width_mm}x${p.height_mm}</text>
            `;
        }
        html += `
            <div style="margin-top:18px;">
                <h3>Sheet ${sheet.sheet_no}</h3>
                <svg width="${result.board_width_mm * scale}" height="${result.board_height_mm * scale}">
                    <rect x="0" y="0" width="${result.board_width_mm * scale}" height="${result.board_height_mm * scale}"
                          fill="white" stroke="#333" stroke-width="2"></rect>
                    ${parts}
                </svg>
            </div>
        `;
    }
    document.getElementById("svgWrap").innerHTML = html;
}

async function optimize() {
    const payload = {
        product_code: document.getElementById("optProductCode").value.trim(),
        board_width_mm: Number(document.getElementById("boardWidth").value),
        board_height_mm: Number(document.getElementById("boardHeight").value),
        kerf_mm: Number(document.getElementById("kerf").value),
        margin_mm: Number(document.getElementById("margin").value),
        rotate_allowed: document.getElementById("rotateAllowed").checked
    };
    const res = await fetch("/optimize", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(payload)
    });
    const data = await res.json();
    document.getElementById("result").textContent = JSON.stringify(data, null, 2);
    if (data.sheets) renderSvg(data);
}
</script>
</body>
</html>
"""


def parse_spec(spec_raw: Any):
    if spec_raw is None or (isinstance(spec_raw, float) and math.isnan(spec_raw)):
        return None, None, None
    text = str(spec_raw).strip()
    match = re.match(r"^\s*(\d+)\s*[*xX]\s*(\d+)\s*[*xX]\s*(\d+)\s*$", text)
    if not match:
        return None, None, None
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


def normalize_value(value: Any):
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def to_int(value: Any, default: int = 0) -> int:
    value = normalize_value(value)
    if value is None or value == "":
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


def load_bom_from_dataframe(df: pd.DataFrame) -> Dict[str, Any]:
    global BOM_ITEMS

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
            "is_cutting_target": is_cutting_target(raw, width, height, thickness)
        }

        if not product_code:
            errors.append({"row": int(idx) + 2, "field": "품목코드", "message": "제품코드 누락"})
        if not part_code:
            errors.append({"row": int(idx) + 2, "field": "부품코드", "message": "부품코드 누락"})
        if item["spec_raw"] and width is None:
            errors.append({"row": int(idx) + 2, "field": "규격", "message": f"규격 파싱 실패: {item['spec_raw']}"})

        items.append(item)

    BOM_ITEMS = items

    product_count = len(sorted({x["product_code"] for x in items if x["product_code"]}))
    return {
        "imported_rows": len(items),
        "product_count": product_count,
        "cutting_target_rows": sum(1 for x in items if x["is_cutting_target"]),
        "error_rows": len(errors),
        "errors": errors[:50]
    }


def expand_parts(parts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    expanded = []
    for p in parts:
        for _ in range(max(1, int(p["qty"]))):
            expanded.append({
                "product_code": p["product_code"],
                "part_code": p["part_code"],
                "part_name": p["part_name"],
                "width_mm": p["width_mm"],
                "height_mm": p["height_mm"],
                "thickness_mm": p["thickness_mm"],
                "material_name": p["material_name"]
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
                "part_name": part["part_name"]
            }

            right = {
                "x": free_rect["x"] + need_w,
                "y": free_rect["y"],
                "w": free_rect["w"] - need_w,
                "h": h
            }
            bottom = {
                "x": free_rect["x"],
                "y": free_rect["y"] + need_h,
                "w": free_rect["w"],
                "h": free_rect["h"] - need_h
            }
            new_rects = [r for r in (right, bottom) if r["w"] > 20 and r["h"] > 20]
            return placement, new_rects
    return None, None


def optimize_parts(parts: List[Dict[str, Any]], board_width: int, board_height: int, kerf: int, margin: int, rotate_allowed: bool):
    usable_w = board_width - margin * 2
    usable_h = board_height - margin * 2
    if usable_w <= 0 or usable_h <= 0:
        raise HTTPException(status_code=400, detail="원판 크기보다 여유치가 큽니다.")

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
                used_rect = sheet["free_rects"].pop(best_idx)
                sheet["free_rects"].extend(new_rects)
                placed = True
                break

        if not placed:
            new_sheet = {
                "sheet_no": len(sheets) + 1,
                "placements": [],
                "free_rects": [{"x": 0, "y": 0, "w": usable_w, "h": usable_h}]
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
        "sheets": [{"sheet_no": s["sheet_no"], "placements": s["placements"]} for s in sheets]
    }


@app.get("/", response_class=HTMLResponse)
def home():
    return HTML_PAGE


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/upload-bom")
async def upload_bom(file: UploadFile = File(...)):
    try:
        df = pd.read_excel(file.file)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {exc}") from exc

    result = load_bom_from_dataframe(df)
    return {
        "success": True,
        "file_name": file.filename,
        **result
    }


@app.get("/products")
def get_products(keyword: str = Query(default="")):
    keyword = keyword.strip().lower()
    products = sorted({x["product_code"] for x in BOM_ITEMS if x["product_code"]})
    if keyword:
        products = [p for p in products if keyword in p.lower()]
    return {"count": len(products), "products": products}


@app.get("/bom/{product_code}")
def get_bom(product_code: str, cutting_only: bool = Query(default=False)):
    items = [x for x in BOM_ITEMS if x["product_code"] == product_code]
    if cutting_only:
        items = [x for x in items if x["is_cutting_target"]]
    if not items:
        raise HTTPException(status_code=404, detail="해당 제품코드의 BOM이 없습니다.")
    return {
        "product_code": product_code,
        "count": len(items),
        "items": items
    }


@app.post("/optimize")
def optimize(payload: Dict[str, Any]):
    product_code = str(payload.get("product_code") or "").strip()
    if not product_code:
        raise HTTPException(status_code=400, detail="product_code가 필요합니다.")

    board_width = to_int(payload.get("board_width_mm"), 2440)
    board_height = to_int(payload.get("board_height_mm"), 1220)
    kerf = to_int(payload.get("kerf_mm"), 3)
    margin = to_int(payload.get("margin_mm"), 5)
    rotate_allowed = bool(payload.get("rotate_allowed", True))

    parts = [x for x in BOM_ITEMS if x["product_code"] == product_code and x["is_cutting_target"]]
    if not parts:
        raise HTTPException(status_code=404, detail="최적화 가능한 재단 대상 부품이 없습니다.")

    result = optimize_parts(parts, board_width, board_height, kerf, margin, rotate_allowed)
    result["product_code"] = product_code
    return JSONResponse(result)
