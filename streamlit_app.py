import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="대학 지원 현황 - 다중 파일 합산", layout="wide")
st.title("대학 지원 현황 시각화 (여러 파일 합산 · 막대그래프)")

st.markdown("""
**사용 안내**  
- 같은 양식의 엑셀 파일을 **여러 개** 업로드하면 **G열(대학명)** 기준으로 **전체 합산 빈도**를 막대그래프로 보여줍니다.  
- 그래프 제목은 **C, D, B열**의 값을 읽어 자동 생성합니다.  
  예) `2025학년도 3학년 6반 수시 지원 대학 시각화` (여러 파일이면 맥락을 요약해서 표시)  
- 공백/결측 값은 `"미기재"`로 처리합니다.  
- 각 대학 막대는 **다채로운 색상 팔레트**를 사용해 표시됩니다.  

📂 **엑셀 파일 저장 방법**  
👉 **나이스 > 대입전형 > 제공현황 조회 > 엑셀파일로 저장**
""")

uploaded_files = st.file_uploader(
    "엑셀 파일(.xlsx)을 하나 이상 업로드하세요",
    type=["xlsx"],
    accept_multiple_files=True
)

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

def make_title_from_many(dfs):
    """여러 파일의 C(2), D(3), B(1) 열 값을 요약하여 제목 생성"""
    c_vals, d_vals, b_vals = set(), set(), set()
    for df in dfs:
        try:
            if df.shape[1] > 2 and pd.notna(df.iloc[0, 2]):
                c_vals.add(str(df.iloc[0, 2]))
            if df.shape[1] > 3 and pd.notna(df.iloc[0, 3]):
                d_vals.add(str(df.iloc[0, 3]))
            if df.shape[1] > 1 and pd.notna(df.iloc[0, 1]):
                b_vals.add(str(df.iloc[0, 1]))
        except Exception:
            continue

    def join_vals(vals):
        if len(vals) == 0:
            return ""
        if len(vals) == 1:
            return list(vals)[0]
        # 여러 값이면 범위를 요약
        return f"{list(vals)[0]} 외"

    c_part = join_vals(c_vals)
    d_part = join_vals(d_vals)
    b_part = join_vals(b_vals)

    parts = [p for p in [c_part, d_part, b_part] if p]
    prefix = " ".join(parts) if parts else "통합"
    return f"{prefix} 수시 지원 대학 시각화 (여러 파일 합산)"

if uploaded_files:
    # 1) 모든 파일 읽기
    dfs = []
    bad_files = []
    for f in uploaded_files:
        df = safe_read_excel(f)
        if df is not None and not df.empty:
            dfs.append(df)
        else:
            bad_files.append(f.name)

    if bad_files:
        st.warning("읽지 못한 파일: " + ", ".join(bad_files))

    if len(dfs) == 0:
        st.stop()

    # 2) 기준 컬럼(G열) 자동 추정 (첫 번째 파일 기준), 필요시 변경 가능
    first_df = dfs[0]
    default_univ_col = default_col_by_letter(first_df, "G") or first_df.columns[0]

    # 컬럼 선택 UI (모든 파일이 같은 구조라고 가정)
    univ_col = st.selectbox(
        "대학(빈도) 컬럼 선택",
        options=list(first_df.columns),
        index=(list(first_df.columns).index(default_univ_col) if default_univ_col in first_df.columns else 0),
        help="보통 G열(7번째 열)이 대학명입니다."
    )

    # 3) 각 파일에서 대학 빈도 집계 → 모두 합산
    per_file_counts = []
    for df in dfs:
        if univ_col not in df.columns:
            st.error(f"선택한 컬럼({univ_col})이 없는 파일이 있습니다. 해당 파일은 건너뜁니다.")
            continue
        cnt = build_univ_counts(df, univ_col)
        cnt["파일"] = "합산대상"
        per_file_counts.append(cnt[["대학", "지원수"]])

    if len(per_file_counts) == 0:
        st.error("선택한 컬럼으로 집계할 수 있는 파일이 없습니다.")
        st.stop()

    # 전체 합산
    total_df = pd.concat(per_file_counts, ignore_index=True)
    total_counts = total_df.groupby("대학", as_index=False)["지원수"].sum()
    total_counts = total_counts.sort_values("지원수", ascending=False, kind="mergesort").reset_index(drop=True)

    # 4) 제목 생성
    graph_title = make_title_from_many(dfs)

    # 5) 상위
