# Wood Cutting Optimizer - Streamlit

Streamlit Cloud에서 바로 실행할 수 있는 목재 재단 프로그램입니다.

구성 파일은 3개입니다.

- `README.md`
- `app.py`
- `requirements.txt`

## 기능

- BOM 엑셀 업로드
- 제품코드 선택
- BOM 목록 조회
- 재단 대상 필터링
- 원판 기준 재단 최적화
- 시트별 분할도 표시

## BOM 컬럼

아래 컬럼을 기준으로 동작합니다.

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

규격 형식 예시:

```text
1814*394*18
```

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포 방법

### 1. GitHub 저장소 생성
파일 3개를 저장소 루트에 올립니다.

### 2. Streamlit Cloud에서 New app 클릭
- Repository: 본인 저장소 선택
- Branch: `main`
- Main file path: `app.py`

### 3. App URL 입력
여기에는 GitHub 주소를 넣지 않습니다.

예:
```text
wood-cutting-app
```

넣으면 안 되는 예:
```text
https://github.com/username/repo
```

### 4. Deploy 클릭
배포가 완료되면 `https://앱이름.streamlit.app` 형태로 접속됩니다.

## 주의
이 버전은 Streamlit용이므로 FastAPI 서버가 아닙니다.

## 향후 개선
- PDF 출력
- SQLite 저장
- 자투리 재사용 관리
- 더 고급 배치 알고리즘
