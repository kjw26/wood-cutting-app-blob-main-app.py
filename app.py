import streamlit as st
import pandas as pd

st.set_page_config(page_title="Wood Cutting Optimizer Test", layout="wide")

st.title("목재 재단 프로그램 - 최소 안정 버전")
st.success("앱이 정상 실행 중입니다.")

st.write("이 버전은 Streamlit 배포 확인용 최소 버전입니다.")

uploaded_file = st.file_uploader("BOM 엑셀 업로드", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"엑셀 업로드 성공: {len(df)}행")
        st.dataframe(df, use_container_width=True)
    except Exception as exc:
        st.error(f"엑셀 파일 읽기 오류: {exc}")

st.info("이 화면이 보이면 Streamlit 배포와 Python 환경은 정상입니다.")
