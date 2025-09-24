import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="대학 지원 현황 - 막대그래프", layout="wide")
st.title("대학 지원 현황 시각화 (막대그래프 전용)")

uploaded = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"])

def safe_read_excel(file):
    try:
        df = pd.read_excel(file, dtype=str)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        return df
    except Exception as e:
        st.error(f"엑셀을 읽는 중 오류: {e}")
        return None

def default_col_by_letter(df, letter):
    """A=1 기준, G=7 → G열"""
    pos = ord(letter.upper()) - ord('A') + 1
    if 1 <= pos <= len(df.columns):
        return df.columns[pos-1]
    return None

if uploaded is not None:
    df = safe_read_excel(uploaded)
    if df is not None and not df.empty:
        univ_col = default_col_by_letter(df, "G")
        if univ_col is None:
            univ_col = st.selectbox("대학 컬럼 선택", df.columns)

        st.subheader("대학별 지원 빈도")
        univ_counts = (
            df[univ_col]
            .fillna("미기재")
            .replace("", "미기재")
            .value_counts()
            .reset_index()
            .rename(columns={"index": "대학", univ_col: "지원수"})
        )

        fig = px.bar(
            univ_counts,
            x="대학",
            y="지원수",
            text="지원수",
            title="대학별 지원 빈도 (G열 기준)"
        )
        fig.update_layout(
            xaxis_tickangle=-45,
            xaxis_title=None,
            yaxis_title=None,
            margin=dict(l=10, r=10, t=60, b=10)
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(univ_counts)
    else:
        st.warning("파일이 비어 있거나 읽을 수 없습니다.")
else:
    st.info("엑셀 파일을 업로드하면 바로 결과를 볼 수 있습니다.")
