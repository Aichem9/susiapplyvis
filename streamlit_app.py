import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="대학 지원 현황 - 막대그래프", layout="wide")

st.title("대학 지원 현황 시각화 (막대그래프 전용)")

st.markdown("""
**사용 안내**  
- 이 앱은 엑셀 파일을 업로드하면 **G열(대학명)**을 기준으로 **대학별 지원 빈도수**를 막대그래프로 보여줍니다.  
- X축은 **대학명**, Y축은 **빈도수(지원 건수)**입니다.  
- 기본값으로 **7번째 열(G열)**을 대학 열로 추정하며, 다를 경우 화면의 선택박스에서 직접 지정하세요.  
- 공백/결측 값은 **"미기재"**로 처리합니다.
""")

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

def build_univ_counts(df, univ_col):
    series = df[univ_col].astype(str)
    series = series.replace({"": "미기재", "NaN": "미기재", "nan": "미기재", "None": "미기재"}).fillna("미기재")
    series = series.apply(lambda x: x.strip() if isinstance(x, str) else x)
    vc = series.value_counts(dropna=False)
    out = vc.rename_axis("대학").reset_index(name="지원수")
    out = out.sort_values("지원수", ascending=False, kind="mergesort").reset_index(drop=True)
    return out

if uploaded is not None:
    df = safe_read_excel(uploaded)
    if df is not None and not df.empty:
        default_univ_col = default_col_by_letter(df, "G")
        if default_univ_col is None:
            default_univ_col = df.columns[0]

        univ_col = st.selectbox(
            "대학(빈도) 컬럼 선택",
            options=list(df.columns),
            index=(list(df.columns).index(default_univ_col) if default_univ_col in df.columns else 0),
            help="보통 G열(7번째 열)이 대학명입니다."
        )

        try:
            univ_counts = build_univ_counts(df, univ_col)
        except KeyError:
            st.error(f"선택한 컬럼을 찾을 수 없습니다: {univ_col}")
            st.stop()

        if univ_counts.empty:
            st.warning("집계 결과가 비어 있습니다. 대학 컬럼 선택을 확인해 주세요.")
            st.dataframe(df.head(20))
            st.stop()

        # 상위 N개 옵션
        c1, c2 = st.columns([1, 3])
        with c1:
            top_n = st.number_input("상위 N개만 표시 (0=전체)", min_value=0, max_value=int(len(univ_counts)), value=min(20, int(len(univ_counts))))
        with c2:
            sort_desc = st.checkbox("빈도 내림차순 정렬", value=True)

        plot_df = univ_counts.copy()
        if sort_desc:
            plot_df = plot_df.sort_values("지원수", ascending=False, kind="mergesort")
        if top_n and top_n > 0:
            plot_df = plot_df.head(int(top_n))

        # 막대그래프
        fig = px.bar(
            plot_df,
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

        # 데이터 표 & 다운로드
        with st.expander("데이터 표 보기"):
            st.dataframe(univ_counts, use_container_width=True)

        st.download_button(
            "빈도표 CSV 다운로드",
            data=univ_counts.to_csv(index=False).encode("utf-8-sig"),
            file_name="대학별_지원빈도.csv",
            mime="text/csv"
        )
    else:
        st.warning("파일이 비어 있거나 읽을 수 없습니다.")
else:
    st.info("엑셀 파일을 업로드하면 바로 결과를 볼 수 있습니다.")
