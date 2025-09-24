import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="대학 지원 현황 - 다중 파일 합산", layout="wide")
st.title("대입 전형자료 조회 데이터 기반 지원 현황 시각화 (다중 파일·막대그래프·컬러풀)")

st.markdown("""
**사용 안내**  
- 같은 양식의 엑셀 파일을 **여러 개 업로드**하면 **모든 파일을 합산**해 대학(G열)별 지원 빈도 막대그래프를 보여줍니다.  
- 그래프 제목은 **단일 파일 업로드 시** C, D, B열(예: `2025학년도 3학년 6반`)을 조합해 자동 생성됩니다. **여러 파일 업로드 시**엔 `전체(다중 파일)`로 표시합니다.  
- 공백/결측은 `"미기재"`로 처리합니다.  
- 각 대학 막대는 **다채로운 색상 팔레트**로 표시됩니다.  
- 인창고 AIchem 제작 : ssac9@sen.go.kr
📂 **엑셀 파일 저장 방법**  
👉 **나이스 > 대입전형 > 제공현황 조회 > 엑셀파일로 저장**
""")

uploaded_files = st.file_uploader("엑셀 파일(.xlsx)을 하나 이상 업로드하세요", type=["xlsx"], accept_multiple_files=True)

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

def build_univ_counts_from_series(series: pd.Series) -> pd.DataFrame:
    s = series.astype(str)
    s = s.replace({"": "미기재", "NaN": "미기재", "nan": "미기재", "None": "미기재"}).fillna("미기재")
    s = s.apply(lambda x: x.strip() if isinstance(x, str) else x)
    vc = s.value_counts(dropna=False)
    out = vc.rename_axis("대학").reset_index(name="지원수")
    out = out.sort_values("지원수", ascending=False, kind="mergesort").reset_index(drop=True)
    return out

def make_title_from_df(df):
    try:
        c_val = str(df.iloc[0, 2]) if df.shape[1] > 2 else ""
        d_val = str(df.iloc[0, 3]) if df.shape[1] > 3 else ""
        b_val = str(df.iloc[0, 1]) if df.shape[1] > 1 else ""
        base = " ".join([v for v in [c_val, d_val, b_val] if v])
        if base.strip():
            return f"{base} 수시 지원 대학 시각화"
    except Exception:
        pass
    return "대학별 지원 빈도 시각화"

if uploaded_files:
    # 첫 파일로 기본 컬럼 추정
    first_df = safe_read_excel(uploaded_files[0])
    if first_df is None or first_df.empty:
        st.warning("첫 번째 파일이 비어 있거나 읽을 수 없습니다.")
        st.stop()

    default_univ_col = default_col_by_letter(first_df, "G") or first_df.columns[0]
    univ_col = st.selectbox(
        "대학(빈도) 컬럼 선택 (모든 파일에 동일하게 적용)",
        options=list(first_df.columns),
        index=(list(first_df.columns).index(default_univ_col) if default_univ_col in first_df.columns else 0),
        help="보통 G열(7번째 열)이 대학명입니다."
    )

    # 단일/다중에 따른 제목
    if len(uploaded_files) == 1:
        graph_title = make_title_from_df(first_df)
    else:
        graph_title = "전체(다중 파일) 수시 지원 대학 시각화"

    # 모든 파일 로드 & 합산
    per_file_counts = []   # 각 파일별 집계 저장 (검증용)
    all_univ_values = []   # 합산용 시리즈 모음

    for f in uploaded_files:
        df = safe_read_excel(f)
        if df is None or df.empty:
            st.warning(f"비어 있거나 읽을 수 없는 파일이 있습니다: {getattr(f, 'name', '파일')}")
            continue
        if univ_col not in df.columns:
            st.warning(f"선택한 컬럼 '{univ_col}'이 없는 파일이 있습니다: {getattr(f, 'name', '파일')}")
            continue

        # 합산을 위해 원시 시리즈만 모으고, 개별 표도 생성
        s = df[univ_col]
        all_univ_values.append(s)
        per_file_counts.append({
            "file": getattr(f, "name", "파일"),
            "counts": build_univ_counts_from_series(s)
        })

    if not all_univ_values:
        st.error("유효한 데이터가 없습니다. 컬럼 선택 또는 파일을 확인해 주세요.")
        st.stop()

    merged_series = pd.concat(all_univ_values, ignore_index=True)
    total_counts = build_univ_counts_from_series(merged_series)

    # 상위 N개 옵션
    c1, c2 = st.columns([1, 3])
    with c1:
        top_n = st.number_input("상위 N개만 표시 (0=전체)", min_value=0, max_value=int(len(total_counts)), value=min(20, int(len(total_counts))))
    with c2:
        sort_desc = st.checkbox("빈도 내림차순 정렬", value=True)

    plot_df = total_counts.copy()
    if sort_desc:
        plot_df = plot_df.sort_values("지원수", ascending=False, kind="mergesort")
    if top_n and top_n > 0:
        plot_df = plot_df.head(int(top_n))

    # 팔레트 (더 컬러풀)
    palette = px.colors.qualitative.Set3 + px.colors.qualitative.Vivid + px.colors.qualitative.Dark24

    # 막대그래프 (전체 합산)
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

    # 전체 합산 표 & 다운로드
    with st.expander("전체 합산 표 보기"):
        st.dataframe(total_counts, use_container_width=True)

    st.download_button(
        "전체 합산 CSV 다운로드",
        data=total_counts.to_csv(index=False).encode("utf-8-sig"),
        file_name="대학별_지원빈도_전체합산.csv",
        mime="text/csv"
    )

    # (선택) 파일별 집계도 확인
    with st.expander("파일별 집계 표 보기"):
        for item in per_file_counts:
            st.markdown(f"**파일:** {item['file']}")
            st.dataframe(item["counts"], use_container_width=True)
            st.markdown("---")
else:
    st.info("엑셀 파일을 1개 이상 업로드하면 전체 합산 결과를 볼 수 있습니다.")
