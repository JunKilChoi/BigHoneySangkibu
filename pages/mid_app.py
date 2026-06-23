import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(
    page_title="중학교 간편 문장 생성",
    page_icon="🍯",
    layout="wide",
)

st.title("🍯 중학교 간편 생기부 문장 생성")
st.caption("활동과 수준별 결과를 바탕으로 비슷하지만 조금씩 다른 생기부 문장을 생성합니다.")

st.info("이 페이지는 중학교용 간편 생성 페이지입니다.")
