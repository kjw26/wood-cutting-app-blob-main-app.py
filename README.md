# Wood Cutting Optimizer

단일 파일 구조로 만든 GitHub 업로드용 목재 재단 프로그램입니다.

구성 파일은 3개입니다.

- `README.md`
- `app.py`
- `requirements.txt`

## 기능

- BOM 엑셀 업로드
- 제품코드 목록 조회
- 제품 BOM 조회
- 재단 대상 필터링
- 원판 기준 재단 최적화
- 브라우저에서 분할도 확인

## 지원 BOM 컬럼

아래 컬럼을 기준으로 동작합니다.

- `품목코드`
- `부품코드`
- `품목명`
- `색상`
- `정소요량`
- `실소요량`
- `규격`
- `재질`
- `소요공정`
- `대표이미지`

`규격`은 아래 형식을 기대합니다.

```text
가로*세로*두께
예) 1814*394*18
```

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
uvicorn app:app --reload
```

실행 후 접속:

```text
http://127.0.0.1:8000
```

Swagger 문서:

```text
http://127.0.0.1:8000/docs
```

## 사용 순서

### 1. BOM 업로드
첫 화면에서 엑셀 파일을 업로드합니다.

또는 API:

```bash
curl -X POST "http://127.0.0.1:8000/upload-bom" \
  -F "file=@BOM_DATA.xlsx"
```

### 2. 제품 목록 조회

```bash
curl "http://127.0.0.1:8000/products"
```

### 3. 제품 BOM 조회

```bash
curl "http://127.0.0.1:8000/bom/CAB0085N?cutting_only=true"
```

### 4. 최적화 실행

```bash
curl -X POST "http://127.0.0.1:8000/optimize" \
  -H "Content-Type: application/json" \
  -d '{
    "product_code": "CAB0085N",
    "board_width_mm": 2440,
    "board_height_mm": 1220,
    "kerf_mm": 3,
    "margin_mm": 5,
    "rotate_allowed": true
  }'
```

## 최적화 방식

현재 버전은 단일 파일 구조에 맞춘 경량 휴리스틱입니다.

- 면적 큰 순 정렬
- free rectangle 기반 배치
- 회전 허용 옵션 지원
- kerf, margin 반영
- 시트별 배치 결과 반환

## 한계

- DB 저장 없음
- 로그인 없음
- PDF 출력 없음
- 고급 MaxRects 전체 구현은 아님

## GitHub 업로드 방법

```bash
git init
git add .
git commit -m "init wood cutting optimizer"
git branch -M main
git remote add origin https://github.com/USERNAME/REPO.git
git push -u origin main
```

## 권장 다음 단계

- SQLite 또는 PostgreSQL 저장
- 결과 PDF 출력
- 부품명/치수 라벨 인쇄
- 자투리 재사용 관리
- React 프론트엔드 분리
