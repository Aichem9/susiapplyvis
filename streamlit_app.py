import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="대학 지원 현황 - 막대그래프", layout="wide")

st.title("대학 지원 현황 시각화 (막대그래프·컬러풀)")

st.markdown("""
**사용 안내**  
- 이 앱은 업로드한 엑셀 파일을 읽어 **G열(대학명)** 기준으로 지원 빈도를 집계하여 막대그래프로 보여줍니다.  
- 그래프 제목은 **C, D, B열** 데이터를 조합하여 자동 생성됩니다.  
  예) `2025학년도 3학년 6반 수시 지원 대학 시각화`  
- 공백/결측 값은 `"미기재"`로 처리합니다.  
- 각 대학 막대는 **다채로운 색상 팔레트**를 사용해 표시됩니다.  

📂 **엑셀 파일 저장 방법**  
👉 나이스 > 대입전형 > 제공현황 조회 > **엑셀파일로 저장**
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

def make_title(df):
    try:
        c_val = str(df.iloc[0, 2]) if df.shape[1] > 2 else ""
        d_val = str(df.iloc[0, 3]) if df.shape[1] > 3 else ""
        b_val = str(df.iloc[0, 1]) if df.shape[1] > 1 else ""
        return f"{c_val} {d_val} {b_val} 수시 지원 대학 시각화"
    except Exception:
        return "대학별 지원 빈도 시각화"

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

        graph_title = make_title(df)

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

        # 🎨 더 컬러풀한 팔레트 지정
        palette = px.colors.qualitative.Set3 + px.colors.qualitative.Vivid + px.colors.qualitative.Dark24

        fig = px.bar(
            plot_df,
            x="대학",
            y="지원수",
            color="대학",
            text="지원수",
            title=graph_title,
            color_discrete_sequence=palette
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            xaxis_tickangle=-45,
            xaxis_title=None,
            yaxis_title=None,
            margin=dict(l=10, r=10, t=60, b=10),
            showlegend=False
        )
        st.plotly_chart(fig, use_container_width=True)

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
