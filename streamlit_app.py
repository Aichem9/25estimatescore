import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="추정 분할 점수 생성기", layout="wide")
st.title("추정 분할 점수 생성기 (고정 규칙 버전)")

st.info("먼저 **사이드바에서 A~E 수준별 정답률(%)**을 설정한 뒤, **문항 정보표 엑셀**을 업로드하세요.", icon="ℹ️")

with st.expander("규칙 요약 및 사용 순서", expanded=False):
    st.markdown("""
    **엑셀 규칙(고정):**
    - **헤더:** 8행 + 9행을 결합하여 컬럼명으로 사용
    - **데이터 시작:** 10행부터
    - **문항 번호:** 10행 A열부터
    - **난이도:** 10행부터
        - **D열에 O표시 → 상**
        - **E열에 O표시 → 중**
        - **F열에 O표시 → 하**
        - (동시에 여러 칸 표시 시 우선순위 **D > E > F**)
    - **배점:** **G열**
    - 슬라이더: **5 단위**로 정답률(%) 입력

    **사용 순서**
    1) 사이드바에서 **A~E 수준 × (상·중·하)** 정답률(%)을 먼저 설정합니다.
    2) **문항 정보표 엑셀(8·9행 헤더, 10행 데이터 시작)**을 업로드합니다.
    3) 전처리 결과와 **수준별 기대 총점**을 확인하고, **분할 점수(컷)**을 검토합니다.
    """)

# ---------- Helpers ----------
def clean_consecutive_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    mask_dup = df.shift(1).eq(df).all(axis=1)
    return df.loc[~mask_dup].reset_index(drop=True)

def is_o_mark(val):
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s in {"o","Ｏ","○","◯","✓","✔","y","yes","true","1","표시","o표시"}

def ensure_numeric(col: pd.Series, default=0.0) -> pd.Series:
    s = pd.to_numeric(col, errors="coerce").fillna(default)
    return s

# ---------- Sidebar: rates ----------
st.sidebar.header("A~E 수준 정답률(%) 입력 (5 단위)")
default_rates = {
    "상": {"A": 70, "B": 55, "C": 40, "D": 25, "E": 10},
    "중": {"A": 85, "B": 70, "C": 55, "D": 40, "E": 25},
    "하": {"A": 95, "B": 85, "C": 70, "D": 55, "E": 40},
}
rates = {}
for band in ["상","중","하"]:
    with st.sidebar.expander(f"{band} 영역 정답률(%)", expanded=(band=="중")):
        rates[band] = {}
        for level in ["A","B","C","D","E"]:
            rates[band][level] = st.slider(
                f"{level} ({band})", 0, 100,
                int(default_rates[band][level]), 5
            ) / 100.0

# ---------- File upload ----------
uploaded = st.file_uploader("문항 정보표 엑셀 업로드 (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("엑셀 파일을 업로드하면 분석이 시작됩니다.")
    st.stop()

# ---------- Read Excel with fixed rules ----------
try:
    uploaded.seek(0)
    raw_all = pd.read_excel(uploaded, header=None, engine="openpyxl")
except Exception as e:
    st.error(f"엑셀을 읽는 중 오류: {e}")
    st.stop()

if len(raw_all) < 10:
    st.error("엑셀 행 수가 부족합니다. 8·9행 헤더, 10행 데이터가 있어야 합니다.")
    st.stop()

# Build header from rows 8 and 9 (0-based 7 and 8)
h1 = raw_all.iloc[7, :].astype(str).fillna("")
h2 = raw_all.iloc[8, :].astype(str).fillna("")
def comb(a,b):
    a = str(a).strip(); b = str(b).strip()
    if a and b: return f"{a} {b}"
    return a or b or ""
header = [comb(a,b) for a,b in zip(h1, h2)]

# Data from row 10 (0-based 9)
df = raw_all.iloc[9:, :].copy()
df.columns = header

# Keep also raw positional columns for D/E/F/G and A
pos_A = raw_all.columns[0] if raw_all.shape[1] > 0 else None
pos_D = raw_all.columns[3] if raw_all.shape[1] > 3 else None
pos_E = raw_all.columns[4] if raw_all.shape[1] > 4 else None
pos_F = raw_all.columns[5] if raw_all.shape[1] > 5 else None
pos_G = raw_all.columns[6] if raw_all.shape[1] > 6 else None

# Trim and deduplicate
df = df.dropna(how="all").dropna(axis=1, how="all")
df = clean_consecutive_duplicates(df)

# ---------- Extract required fields ----------
# Item number from column A (pos 0) starting row 10
items = raw_all.iloc[9:, 0] if pos_A is not None else pd.Series(range(1, len(df)+1))
items = items.iloc[:len(df)].reset_index(drop=True)
items.name = "문항번호"

# Difficulty band from D/E/F O-marks with priority D>E>F
def derive_band(irow):
    r = 9 + irow  # offset from row 10
    hard = raw_all.iloc[r, 3] if pos_D is not None and r < len(raw_all) else None
    mid  = raw_all.iloc[r, 4] if pos_E is not None and r < len(raw_all) else None
    easy = raw_all.iloc[r, 5] if pos_F is not None and r < len(raw_all) else None
    if is_o_mark(hard): return "상"
    if is_o_mark(mid):  return "중"
    if is_o_mark(easy): return "하"
    return None

band = pd.Series([derive_band(i) for i in range(len(df))], name="난이도영역")

# Points from column G starting row 10
points_raw = raw_all.iloc[9:, 6] if pos_G is not None else pd.Series([np.nan]*len(df))
points = ensure_numeric(points_raw, default=np.nan).iloc[:len(df)].reset_index(drop=True)
points.name = "배점"

# Build working table
work = pd.DataFrame({
    "문항번호": items,
    "난이도영역": band,
    "배점": points,
})

# Filter valid rows
valid = work.dropna(subset=["난이도영역","배점"]).reset_index(drop=True)

# ---------- Calculate expected scores ----------
levels = ["A","B","C","D","E"]

rows = []
for _, row in valid.iterrows():
    ent = {"문항번호": row["문항번호"], "난이도영역": row["난이도영역"], "배점": row["배점"]}
    for lvl in levels:
        ent[f"{lvl}_기대득점"] = rates[row["난이도영역"]][lvl] * row["배점"]
    rows.append(ent)
df_item = pd.DataFrame(rows)

exp_totals = {lvl: float(df_item[f"{lvl}_기대득점"].sum()) for lvl in levels}

def mid(a,b): return (a+b)/2.0
cuts = {
    "A/B 컷": mid(exp_totals["A"], exp_totals["B"]),
    "B/C 컷": mid(exp_totals["B"], exp_totals["C"]),
    "C/D 컷": mid(exp_totals["C"], exp_totals["D"]),
    "D/E 컷": mid(exp_totals["D"], exp_totals["E"]),
}

# ---------- Show Cut Scores FIRST (with color coding) ----------
st.subheader("추정 분할 점수 (최우선)")

cuts_order = ["A/B 컷", "B/C 컷", "C/D 컷", "D/E 컷"]
df_cuts = pd.DataFrame([[cuts.get(k) for k in cuts_order]], columns=cuts_order)

_color_map = {
    "A/B 컷": "#E8F5E9",  # green
    "B/C 컷": "#E3F2FD",  # blue
    "C/D 컷": "#FFF3E0",  # orange
    "D/E 컷": "#FFEBEE",  # red
}
def _style_cuts(df):
    return pd.DataFrame([[f"background-color: {_color_map.get(c, '')}" for c in df.columns]], index=df.index)

st.dataframe(df_cuts.style.apply(_style_cuts, axis=None).format("{:.2f}"), use_container_width=True)

m1, m2, m3, m4 = st.columns(4)
m1.metric("A/B 컷", f"{df_cuts.iloc[0,0]:.2f}")
m2.metric("B/C 컷", f"{df_cuts.iloc[0,1]:.2f}")
m3.metric("C/D 컷", f"{df_cuts.iloc[0,2]:.2f}")
m4.metric("D/E 컷", f"{df_cuts.iloc[0,3]:.2f}")

# ---------- Summary ----------
c1,c2,c3 = st.columns(3)
with c1: st.metric("총 행(데이터)", len(df))
with c2: st.metric("유효 문항(난이도+배점)", len(valid))
with c3:
    cnt = valid["난이도영역"].value_counts().reindex(["상","중","하"], fill_value=0)
    st.metric("상/중/하 분포", f"상 {cnt['상']} / 중 {cnt['중']} / 하 {cnt['하']}")

st.subheader("수준별 기대 총점")
st.dataframe(pd.DataFrame([exp_totals]), use_container_width=True)

st.subheader("문항별 수준별 기대득점")
st.dataframe(df_item, use_container_width=True)

st.subheader("전처리 결과 미리보기")
st.dataframe(valid.head(30), use_container_width=True)

# ---------- Download ----------
out = {
    "전처리_문항정보": valid,
    "수준별_기대총점": pd.DataFrame([exp_totals]),
    "추정_분할점수": pd.DataFrame([cuts]),
    "문항별_수준별_기대득점": df_item,
}
def to_excel_bytes(sheets: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.read()

xlsx = to_excel_bytes(out)
st.download_button("결과 엑셀 내려받기", data=xlsx,
                   file_name="추정_분할_점수_결과.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.success("완료: 고정 규칙에 따라 난이도·배점·문항을 해석하고 컷 점수를 계산했습니다.")
