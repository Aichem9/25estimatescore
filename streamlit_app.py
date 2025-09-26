
import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="추정 분할 점수 생성기", layout="wide")

st.title("추정 분할 점수 생성기 (Cut Score Estimator)")

with st.expander("도움말 / 개요", expanded=False):
    st.markdown("""
    이 도구는 문항 정보표(엑셀)를 업로드하면 **난이도(상/중/하)** 와 **배점**을 추출하고,
    **A~E 수준 학생의 정답률**을 입력하여 **추정 분할 점수(컷 점수)** 를 산출합니다.

    - 데이터는 **엑셀의 10행부터** 시작한다고 가정합니다. (즉, `skiprows=9`)
    - 동일한 행이 연속해서 나타나는 **연속 중복 행은 자동 제거**합니다.
    - 난이도 값은 텍스트(상/중/하) 또는 수치일 수 있습니다.
      - 수치 난이도는 **분위 기반 3분할(상/중/하)** 로 자동 매핑합니다. (상=어려움, 하=쉬움)
      - 텍스트 난이도는 `상/중/하` 혹은 이와 유사 키워드를 인식합니다.
    - A~E 수준별 **난이도 영역(상/중/하)** 정답률(%)을 입력하면,
      각 수준의 **기대 총점**을 계산하고 **인접 수준 간 중간값**을 컷 점수로 제안합니다.
      - 예: E-D 컷 = (E 기대총점 + D 기대총점) / 2

    필요에 따라 결과를 CSV로 내려받을 수 있습니다.
    """)

# 파일 업로드
uploaded = st.file_uploader("문항 정보표 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

# 파싱 보조 함수들
TEXT_DIFFICULTY_MAP = {
    "상": "상", "중": "중", "하": "하",
    "어려움": "상", "보통": "중", "쉬움": "하",
    "상등급": "상", "중등급": "중", "하등급": "하",
    "상급": "상", "중급": "중", "하급": "하",
}

def is_checked(val):
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    # common marks for checked
    return (
        s in {"1","y","yes","true","t","o","ok","v","✓","✔","●","■","checked","check","예","ㅇ","응"} or
        s == "x" or
        s == "○" or
        s == "●" or
        s == "✔" or
        s == "✓" or
        (isinstance(val, bool) and val is True)
    )

def normalize_colname(col: str) -> str:
    c = str(col).strip()
    # 우선순위 키워드 매핑(여러 경우를 대비)
    replacements = {
        "문항": ["문항", "문항번호", "번호", "item", "question", "id"],
        "난이도": ["난이도", "difficulty", "난도", "난이"],
        "배점": ["배점", "점수", "득점", "점"],
        "정답": ["정답", "답", "answer", "key"],
        "단원": ["단원", "영역", "topic", "domain"]
    }
    for std, keys in replacements.items():
        for k in keys:
            if k.lower() in c.lower():
                return std
    return c  # 원본 유지

RAW_G_INDEX = 6  # Excel column G (0-based index)

def find_cols(df: pd.DataFrame):
    norm = {col: normalize_colname(col) for col in df.columns}
    inv = {}
    for orig, n in norm.items():
        inv.setdefault(n, []).append(orig)
    col_item = (inv.get("문항") or [df.columns[0]])[0]
    col_diff = (inv.get("난이도") or [None])[0]
    col_point = (inv.get("배점") or [None])[0]
    return col_item, col_diff, col_point

def is_number(x):
    try:
        float(x)
        return True
    except Exception:
        return False

def map_numeric_to_band(series: pd.Series):
    # 낮은 값 = 쉬움(하), 높은 값 = 어려움(상)이라고 가정
    s = pd.to_numeric(series, errors="coerce")
    q1 = s.quantile(1/3)
    q2 = s.quantile(2/3)
    band = pd.Series(index=series.index, dtype="object")
    band[s <= q1] = "하"
    band[(s > q1) & (s <= q2)] = "중"
    band[s > q2] = "상"
    return band

def map_text_to_band(series: pd.Series):
    def f(x):
        if pd.isna(x):
            return None
        sx = str(x).strip()
        # 정확 매칭
        if sx in TEXT_DIFFICULTY_MAP:
            return TEXT_DIFFICULTY_MAP[sx]
        # 부분 키워드
        for k, v in TEXT_DIFFICULTY_MAP.items():
            if k in sx:
                return v
        return None
    return series.apply(f)

def clean_consecutive_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    # 직전 행과 **완전히 동일**한 행은 제거 (연속 중복 제거)
    # (요청 문구가 다소 모호하여 '연속 중복 제거'로 구현)
    mask_dup = df.shift(1).eq(df).all(axis=1)
    return df.loc[~mask_dup].reset_index(drop=True)

def ensure_points(col: pd.Series) -> pd.Series:
    pts = pd.to_numeric(col, errors="coerce")
    # 배점 결측은 1점으로 보정
    pts = pts.fillna(1.0)
    return pts

def summarize_counts(df_band: pd.Series):
    return df_band.value_counts(dropna=False).reindex(["상", "중", "하"], fill_value=0)

# 사이드바 입력: A~E 수준별 정답률(상/중/하 별)
st.sidebar.header("A~E 수준 정답률(%) 입력")
default_rates = {
    # 가이드 라인(예시 값): 수준이 높을수록 어려운 문항에서의 정답률이 상대적으로 낮게 감소
    # 필요 시 교사가 자유 조정
    "상": {"A": 70, "B": 55, "C": 40, "D": 25, "E": 10},
    "중": {"A": 85, "B": 70, "C": 55, "D": 40, "E": 25},
    "하": {"A": 95, "B": 85, "C": 70, "D": 55, "E": 40},
}
rates = {}
for band in ["상", "중", "하"]:
    with st.sidebar.expander(f"{band} 영역 정답률(%)", expanded=(band=="중")):
        rates[band] = {}
        for level in ["A", "B", "C", "D", "E"]:
            val = st.slider(f"{level} ({band})", 0, 100, int(default_rates[band][level]), 5)
            rates[band][level] = val / 100.0

if uploaded is None:
    st.info("엑셀 파일을 업로드하면 분석이 시작됩니다.")
    st.stop()

# 엑셀 읽기: 10행부터 데이터 시작
try:
    src = pd.read_excel(uploaded, skiprows=9, engine="openpyxl")
except Exception as e:
    st.error(f"엑셀을 읽는 중 오류: {e}")
    st.stop()

# 완전 비어있는 열/행 제거
df = src.copy()
df = df.dropna(how="all").dropna(axis=1, how="all")

# 연속 중복 행 제거
df = clean_consecutive_duplicates(df)

# 핵심 컬럼 찾기
col_item, col_diff, col_point = find_cols(df)

if col_diff is None or col_point is None:
    st.warning("난이도 또는 배점 컬럼을 자동으로 찾지 못했습니다. 컬럼 헤더를 확인해 주세요.")
    st.write("인식된 컬럼:", list(df.columns))
    st.stop()

# 난이도 밴드 생성

# 체크박스 기반 난이도 해석 (어려움/보통/쉬움 열이 모두 존재하는 경우 우선 적용)
cols_lower = {c.lower(): c for c in df.columns}
has_checkbox_cols = all(k in cols_lower for k in ["어려움", "보통", "쉬움"])

if has_checkbox_cols:
    col_hard = cols_lower["어려움"]
    col_mid = cols_lower["보통"]
    col_easy = cols_lower["쉬움"]
    def band_from_checks(row):
        if is_checked(row.get(col_hard)):
            return "상"
        if is_checked(row.get(col_mid)):
            return "중"
        if is_checked(row.get(col_easy)):
            return "하"
        return None
    band = df.apply(band_from_checks, axis=1)
else:
    diff_col = df[col_diff]
    if diff_col.apply(is_number).mean() > 0.6:
        band = map_numeric_to_band(diff_col)
    else:
        band = map_text_to_band(diff_col)


# 강제: 배점은 원본 엑셀의 G열에서 가져오기
# 업로드 버퍼를 다시 읽어 원본 기준으로 G열 배점을 확보
uploaded.seek(0)
try:
    raw = pd.read_excel(uploaded, header=None, engine="openpyxl")
    pts_series = pd.to_numeric(raw.iloc[9:, RAW_G_INDEX].reset_index(drop=True), errors="coerce")
    # 전처리된 표(valid 생성 전)의 길이에 맞춰 자르거나 보정
    # 우선 df와 동일 길이로 맞추기
    pts_series = pts_series.iloc[:len(df)].reindex(range(len(df)))
    points = ensure_points(pts_series)
except Exception:
    # 폴백: 자동 인식된 배점 컬럼 사용
    points = ensure_points(df[col_point])
items = df[col_item] if col_item in df.columns else pd.Series(range(1, len(df)+1), name="문항")

data = pd.DataFrame({
    "문항": items,
    "난이도원본": diff_col,
    "난이도영역": band,
    "배점": points,
})

# 유효 문항만 필터 (난이도영역 결측 제외)
valid = data.dropna(subset=["난이도영역"]).reset_index(drop=True)

# 요약
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("총 행(원자료)", len(src))
with c2:
    st.metric("유효 문항(난이도영역 식별)", len(valid))
with c3:
    counts = summarize_counts(valid["난이도영역"])
    st.metric("상/중/하 분포", f"상 {counts['상']}, 중 {counts['중']}, 하 {counts['하']}")

st.subheader("전처리 결과 미리보기")
st.dataframe(valid.head(30), use_container_width=True)

# 수준별 기대 총점 계산
bands = ["상", "중", "하"]
levels = ["A", "B", "C", "D", "E"]

# 각 문항에 대해 수준별 정답확률 할당
prob_matrix = {lvl: [] for lvl in levels}
for _, row in valid.iterrows():
    b = row["난이도영역"]
    p = row["배점"]
    for lvl in levels:
        pr = rates[b][lvl]  # 입력된 확률
        prob_matrix[lvl].append(pr * p)

expected_scores = {lvl: float(np.sum(prob_matrix[lvl])) for lvl in levels}

# 컷 점수: 인접 수준의 기대 총점 중간값
def cut_between(low_level, high_level):
    return (expected_scores[low_level] + expected_scores[high_level]) / 2.0

# E-D, D-C, C-B, B-A 컷 제안
cuts = {
    "E/D 컷": cut_between("E", "D"),
    "D/C 컷": cut_between("D", "C"),
    "C/B 컷": cut_between("C", "B"),
    "B/A 컷": cut_between("B", "A"),
}

st.subheader("수준별 기대 총점 & 추정 분할 점수(컷)")
colA, colB = st.columns(2)

with colA:
    st.markdown("**수준별 기대 총점**")
    df_exp = pd.DataFrame([expected_scores])
    st.dataframe(df_exp, use_container_width=True)

with colB:
    st.markdown("**추정 분할 점수 제안(중간값 방식)**")
    df_cuts = pd.DataFrame([cuts])
    st.dataframe(df_cuts, use_container_width=True)

# 문항별 수준별 기대득점(= 배점 × 정답확률) 테이블도 제공
rows = []
for i, row in valid.iterrows():
    ent = {"문항": row["문항"], "난이도영역": row["난이도영역"], "배점": row["배점"]}
    for lvl in levels:
        ent[f"{lvl}_기대득점"] = rates[row["난이도영역"]][lvl] * row["배점"]
    rows.append(ent)
df_item_level = pd.DataFrame(rows)

st.subheader("문항별 수준별 기대득점")
st.dataframe(df_item_level, use_container_width=True)

# 다운로드: 결과 묶음
out = {
    "전처리_문항정보": valid,
    "수준별_기대총점": pd.DataFrame([expected_scores]),
    "추정_분할점수": pd.DataFrame([cuts]),
    "문항별_수준별_기대득점": df_item_level,
}

def to_excel_bytes(sheets: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, d in sheets.items():
            d.to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.read()

xlsx_bytes = to_excel_bytes(out)
st.download_button(
    label="결과 엑셀 내려받기",
    data=xlsx_bytes,
    file_name="추정_분할_점수_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("완료: 입력한 정답률과 문항정보를 기반으로 추정 분할 점수(컷)를 계산했습니다.")
