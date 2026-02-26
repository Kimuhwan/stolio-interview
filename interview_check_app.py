import os
import io
import time
from datetime import datetime
import pandas as pd
import streamlit as st

APP_VERSION = "1.5.0"

DEFAULT_INPUT_FILE = "Stolio_5기_면접.xlsx"
DEFAULT_OUTPUT_DIR = "outputs"

EVAL_COLUMNS = [
    "timestamp", "app_version", "interviewer",
    "candidate_id", "name", "student_id", "mark",
    "category", "level",
    "score_rules_fit", "score_output_evidence", "score_collaboration", "score_self_driven", "score_role_skill", "score_overall",
    "flag_evidence_risk", "flag_schedule_risk", "flag_attitude_risk", "flag_comm_risk", "flag_other_risk",
    "memo_strength", "memo_concern", "memo_followup", "memo_summary",
    "recommendation",
]

def now_string():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_output_dir(path):
    os.makedirs(path, exist_ok=True)

def safe_str(x):
    if pd.isna(x): return ""
    return str(x)

def student_prefix(sid: str) -> str:
    sid = (sid or "").strip()
    return sid[:2] if len(sid) >= 2 else ""

def is_21_25(sid: str) -> bool:
    return student_prefix(sid) in {"21","22","23","24","25"}

def load_candidates(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    # 최소 컬럼 체크/정규화
    if "이름" not in df.columns:
        # 혹시 다른 시트/컬럼이면 여기서 더 확장 가능
        raise ValueError("엑셀에 '이름' 컬럼이 없습니다. 쉬운버전 엑셀을 사용하세요.")

    if "학번" not in df.columns:
        df["학번"] = ""

    if "학번표시" not in df.columns:
        # 비26 표시가 없다면 간단 생성
        df["학번표시"] = df["학번"].astype(str).apply(lambda x: "" if student_prefix(x) == "26" else "⚠️ 26학번 아님")

    if "예상레벨" not in df.columns:
        # 호환: 레벨추정이 있으면 사용
        if "레벨추정" in df.columns:
            df["예상레벨"] = df["레벨추정"]
        else:
            df["예상레벨"] = ""

    if "분류" not in df.columns:
        df["분류"] = ""

    # 후보 ID
    df["학번"] = df["학번"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["_candidate_id"] = (df["학번"].fillna("").astype(str).str.strip() + "_" + df["이름"].fillna("").astype(str).str.strip()).str.strip("_")

    return df

def candidate_label(r: pd.Series) -> str:
    name = safe_str(r.get("이름",""))
    sid  = safe_str(r.get("학번",""))
    mark = safe_str(r.get("학번표시",""))
    cat  = safe_str(r.get("분류",""))
    lvl  = safe_str(r.get("예상레벨", r.get("레벨추정","")))
    prefix = "⚠️ " if ("⚠️" in mark) else ""
    tail = []
    if cat: tail.append(cat)
    if lvl: tail.append(lvl)
    tail_str = f" - {' / '.join(tail)}" if tail else ""
    return f"{prefix}{name} ({sid}){tail_str}".strip()

def candidate_label_with_status(r: pd.Series, evaluated_ids: set) -> str:
    """평가 완료 여부(✅/❌)를 포함한 라벨"""
    cid = safe_str(r.get("_candidate_id",""))
    base = candidate_label(r)
    status = "✅" if cid in evaluated_ids else "❌"
    return f"{status} {base}"

def empty_evals():
    return pd.DataFrame(columns=EVAL_COLUMNS)

def load_results(path: str) -> pd.DataFrame:
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, sheet_name="Evaluations")
        except Exception:
            return empty_evals()
        for c in EVAL_COLUMNS:
            if c not in df.columns:
                df[c] = ""
        return df[EVAL_COLUMNS].copy()
    return empty_evals()

def upsert_eval(evals: pd.DataFrame, row: dict) -> pd.DataFrame:
    if evals.empty:
        return pd.DataFrame([row], columns=EVAL_COLUMNS)
    mask = (evals["interviewer"] == row["interviewer"]) & (evals["candidate_id"] == row["candidate_id"])
    if mask.any():
        idx = evals.index[mask][0]
        for k,v in row.items():
            evals.at[idx, k] = v
        return evals
    return pd.concat([evals, pd.DataFrame([row], columns=EVAL_COLUMNS)], ignore_index=True)

def save_results(path: str, evals: pd.DataFrame, candidates_snapshot: pd.DataFrame):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        evals.to_excel(writer, index=False, sheet_name="Evaluations")
        snap_cols = ["_candidate_id","이름","학번","학번표시","분류","예상레벨","중복지원","이메일","전화번호"]
        snap_cols = [c for c in snap_cols if c in candidates_snapshot.columns]
        candidates_snapshot[snap_cols].to_excel(writer, index=False, sheet_name="CandidatesSnapshot")
        pd.DataFrame([{"app_version": APP_VERSION, "generated_at": now_string()}]).to_excel(writer, index=False, sheet_name="Meta")

def auto_avg(nums):
    vals = [float(x) for x in nums if isinstance(x,(int,float)) and float(x) > 0]
    if not vals: return 0.0
    return round(sum(vals)/len(vals), 2)

# ---------------- UI ----------------
st.set_page_config(page_title="Stolio 면접 체크", layout="wide", page_icon="📋")

# ---- Custom CSS ----
st.markdown("""
<style>
    /* 전체 배경 및 폰트 */
    .block-container { padding-top: 1.5rem; }
    
    /* 헤더 스타일 */
    h1 { color: #1a73e8; font-weight: 700; letter-spacing: -0.5px; }
    h2, h3, h4 { color: #333; }
    
    /* 카드 스타일 */
    .card {
        background: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .card-info {
        background: linear-gradient(135deg, #e8f0fe 0%, #f0f4ff 100%);
        border: 1px solid #c4d7f5;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    .card-answer {
        background: #fafbfc;
        border-left: 4px solid #1a73e8;
        border-radius: 0 8px 8px 0;
        padding: 0.8rem 1rem;
        margin-bottom: 0.6rem;
    }
    .card-question {
        background: #fff8e1;
        border-left: 4px solid #f9a825;
        border-radius: 0 8px 8px 0;
        padding: 0.8rem 1rem;
        margin-bottom: 0.6rem;
    }
    
    /* 사이드바 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #2B8045 0%, darkgreen 100%);
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3, [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] .stCaption {
        color: #ffffff !important;
    }
    /* 사이드바 토글/체크박스/라디오 라벨 */
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] label,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] span {
        color: #ffffff !important;
    }
    /* 사이드바 인풋 필드 텍스트 */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        color: black !important;
        background-color: rgba(255,255,255,0.12) !important;
        border-color: rgba(255,255,255,0.3) !important;
    }
    [data-testid="stSidebar"] input::placeholder,
    [data-testid="stSidebar"] textarea::placeholder {
        color: rgba(255,255,255,0.5) !important;
    }
    /* 사이드바 숫자 인풋 버튼 */
    [data-testid="stSidebar"] button {
        color: #ffffff !important;
        border-color: rgba(255,255,255,0.3) !important;
    }
    /* 사이드바 토글 텍스트 */
    [data-testid="stSidebar"] [data-testid="stCheckbox"] span,
    [data-testid="stSidebar"] .st-emotion-cache-1gulkj5,
    [data-testid="stSidebar"] .st-emotion-cache-nahz7x {
        color: #ffffff !important;
    }
    /* 사이드바 divider */
    [data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.2) !important;
    }
    /* 사이드바 help 아이콘 */
    [data-testid="stSidebar"] .stTooltipIcon svg {
        fill: rgba(255,255,255,0.7) !important;
    }
    
    /* 버튼 스타일 */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    
    /* 슬라이더 */
    .stSlider > div > div > div > div {
        background-color: #1a73e8;
    }
    
    /* 프로그레스바 */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #1a73e8, #4285f4);
    }
    
    /* 성공/경고 메시지 */
    .stSuccess { border-radius: 8px; }
    .stWarning { border-radius: 8px; }
    
    /* 텍스트 영역 */
    .stTextArea textarea {
        border-radius: 8px;
        border: 1px solid #ddd;
    }
    .stTextArea textarea:focus {
        border-color: #1a73e8;
        box-shadow: 0 0 0 1px #1a73e8;
    }
    
    /* 구분선 */
    hr { border-color: #e8eaed; }
    
    /* 탭 헤더 스타일 */
    .badge-pass { background: #e6f4ea; color: #137333; padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 0.85rem; }
    .badge-hold { background: #fef7e0; color: #b45309; padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 0.85rem; }
    .badge-fail { background: #fce8e6; color: #c5221f; padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 0.85rem; }
    .badge-none { background: #f1f3f4; color: #5f6368; padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 0.85rem; }
</style>
""", unsafe_allow_html=True)

st.title("📋 Stolio 면접 체크 프로그램")
st.caption(f"v{APP_VERSION} · 쉬운질문 / 검색 / 정렬 / 점수저장 / 병합 / 타이머")

with st.sidebar:
    st.header("설정")
    input_file = st.text_input("지원자 엑셀 경로", value=DEFAULT_INPUT_FILE)
    interviewer = st.text_input("면접관 이름(필수)", value="")
    output_dir = st.text_input("저장 폴더", value=DEFAULT_OUTPUT_DIR)
    result_filename = st.text_input("결과 파일명(기본)", value="interview_results.xlsx")
    st.divider()

    st.subheader("지원자 리스트")
    search = st.text_input("지원자 검색(이름/학번)", value="", placeholder="예: 김경환 / 260123")
    pin_21_25_top = st.toggle("21~25학번 위로 올리기", value=False)

    st.divider()
    st.subheader("타이머")
    enable_timer = st.toggle("타이머 사용", value=True)
    minutes = st.number_input("면접 시간(분)", min_value=1, max_value=30, value=8, step=1)
    live_timer = st.toggle("실시간 갱신(1초)", value=True, help="외부 패키지 없이 sleep+rerun으로 1초 갱신합니다.")

if not interviewer.strip():
    st.warning("사이드바에서 **면접관 이름**을 입력하세요.")
    st.stop()

try:
    candidates = load_candidates(input_file)
except Exception as e:
    st.error(f"지원자 엑셀 로드 실패: {e}")
    st.stop()

ensure_output_dir(output_dir)
result_path = os.path.join(output_dir, f"{os.path.splitext(result_filename)[0]}_{interviewer}.xlsx")
evals = load_results(result_path)

# ---- filter/sort candidates ----
view_df = candidates.copy()

# 검색
if search.strip():
    s = search.strip().lower()
    view_df = view_df[
        view_df["이름"].astype(str).str.lower().str.contains(s, na=False)
        | view_df["학번"].astype(str).str.lower().str.contains(s, na=False)
    ].copy()

# 21~25 pin
if pin_21_25_top:
    view_df["_pin"] = view_df["학번"].astype(str).str[:2].isin(["21","22","23","24","25"])
    view_df = view_df.sort_values(by=["_pin","이름"], ascending=[False, True]).drop(columns=["_pin"])
else:
    view_df = view_df.sort_values(by=["이름"], ascending=[True])

if view_df.empty:
    st.warning("검색/정렬 조건에 해당하는 지원자가 없습니다.")
    st.stop()

labels = view_df.apply(candidate_label, axis=1).tolist()
label_to_index = {labels[i]: int(view_df.index[i]) for i in range(len(labels))}

# 평가 완료된 candidate_id 집합
evaluated_ids = set(evals[evals["interviewer"] == interviewer]["candidate_id"].unique()) if not evals.empty else set()
labels_with_status = view_df.apply(lambda r: candidate_label_with_status(r, evaluated_ids), axis=1).tolist()
status_label_to_index = {labels_with_status[i]: int(view_df.index[i]) for i in range(len(labels_with_status))}

# -------- Layout (stable) --------

# [수정 핵심] 변수 정의를 컬럼 나누기 '전'에 수행하여 에러 원천 차단
st.subheader("지원자 선택")
selected_label = st.selectbox("지원자", labels_with_status, index=0)
row_idx = status_label_to_index[selected_label]
r = candidates.loc[row_idx]

# ★★★ 여기서 변수를 미리 다 만들어야 '저장 버튼'이 에러가 안 납니다 ★★★
candidate_id = safe_str(r.get("_candidate_id",""))
name = safe_str(r.get("이름",""))
sid  = safe_str(r.get("학번",""))
mark = safe_str(r.get("학번표시",""))
cat  = safe_str(r.get("분류",""))
lvl  = safe_str(r.get("예상레벨","")) 
dup  = safe_str(r.get("중복지원",""))
# ------------------------------------------------------------------

left, right = st.columns([1,2], gap="large")

with left:
    # 이미 위에서 변수를 만들었으니 여기선 출력만 합니다.
    st.markdown("#### 👤 기본 정보")
    info_lines = f"""<div class='card-info'>
    <b style='font-size:1.15em;'>{name}</b> <span style='color:#666;'>({sid})</span>
    """
    if mark: info_lines += f"<br/>📌 {mark}"
    if cat: info_lines += f"<br/>📂 분류: {cat}"
    if lvl: info_lines += f"<br/>📊 예상레벨: {lvl}"
    if dup: info_lines += f"<br/>⚠️ 중복지원: {dup}"
    info_lines += "</div>"
    st.markdown(info_lines, unsafe_allow_html=True)

    st.divider()

    # -------- Timer (변수 사용) --------
    if enable_timer:
        st.markdown("#### ⏱️ 면접 타이머")
        total = int(minutes) * 60
        k_running = f"timer_running_{candidate_id}"
        k_started = f"timer_started_{candidate_id}"
        k_elapsed = f"timer_elapsed_{candidate_id}"

        if k_running not in st.session_state:
            st.session_state[k_running] = False
            st.session_state[k_started] = 0.0
            st.session_state[k_elapsed] = 0.0

        cA, cB, cC = st.columns(3)
        with cA:
            if st.button("▶️ 시작/재개", use_container_width=True, key=f"btn_start_{candidate_id}"):
                if not st.session_state[k_running]:
                    st.session_state[k_running] = True
                    st.session_state[k_started] = time.time()
        with cB:
            if st.button("⏸️ 일시정지", use_container_width=True, key=f"btn_pause_{candidate_id}"):
                if st.session_state[k_running]:
                    st.session_state[k_elapsed] += max(0.0, time.time() - float(st.session_state[k_started]))
                    st.session_state[k_running] = False
        with cC:
            if st.button("🔁 리셋", use_container_width=True, key=f"btn_reset_{candidate_id}"):
                st.session_state[k_running] = False
                st.session_state[k_started] = 0.0
                st.session_state[k_elapsed] = 0.0

        elapsed = float(st.session_state[k_elapsed])
        if st.session_state[k_running]:
            elapsed += max(0.0, time.time() - float(st.session_state[k_started]))

        remaining = max(0, total - elapsed)
        st.progress(min(1.0, elapsed / total) if total > 0 else 0.0)
        timer_color = "#c5221f" if remaining < 60 else ("#f9a825" if remaining < 120 else "#137333")
        timer_status = "🟢 진행중" if st.session_state[k_running] else "🔴 일시정지"
        st.markdown(f"<div style='text-align:center; font-size:1.5em; font-weight:700; color:{timer_color};'>{int(remaining//60)}:{int(remaining%60):02d}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:center; color:#666;'>{timer_status}</div>", unsafe_allow_html=True)
        
        # 실시간 갱신
        if live_timer and st.session_state[k_running] and remaining > 0:
            time.sleep(1)
            st.rerun()

    st.divider()

    # -------- Existing evaluation preview --------
    mask = (evals["interviewer"] == interviewer) & (evals["candidate_id"] == candidate_id)
    if mask.any():
        last = evals[mask].iloc[-1]
        st.markdown("#### 📝 저장된 평가")
        rec_val = safe_str(last.get('recommendation',''))
        badge_cls = {"합격": "badge-pass", "보류": "badge-hold", "불합": "badge-fail"}.get(rec_val, "badge-none")
        score_val = safe_str(last.get('score_overall',''))
        summ = safe_str(last.get("memo_summary",""))
        eval_html = f"""<div class='card'>
            <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:0.5rem;'>
                <span>종합 점수: <b>{score_val}</b> / 5.0</span>
                <span class='{badge_cls}'>{rec_val if rec_val else '미정'}</span>
            </div>
            <div style='color:#888; font-size:0.85em;'>🕐 {safe_str(last.get('timestamp',''))}</div>
        """
        if summ:
            eval_html += f"<div style='margin-top:0.5rem; padding:0.5rem; background:#f8f9fa; border-radius:6px; font-size:0.9em;'>📌 {summ}</div>"
        eval_html += "</div>"
        st.markdown(eval_html, unsafe_allow_html=True)
    else:
        st.info("아직 저장된 평가가 없습니다.")

with right:
    st.subheader("📄 지원서 답변 & 면접 질문")

    st.markdown("#### 💬 지원서 답변")
    for q_label, q_key in [("지원동기", "지원서답변1(동기)"), ("기대/매력", "지원서답변2(기대/매력)"), ("관심/경험", "지원서답변3(관심/경험)")]:
        answer_text = safe_str(r.get(q_key, ""))
        if answer_text:
            st.markdown(f"<div class='card-answer'><b>{q_label}</b><br/>{answer_text}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='card-answer'><b>{q_label}</b><br/><span style='color:#999;'>작성 내용 없음</span></div>", unsafe_allow_html=True)

    st.markdown("#### 🎯 면접 질문")
    q_items = [
        ("공통Q1", "공통Q1"), ("공통Q2", "공통Q2"), ("공통Q3", "공통Q3"),
        ("맞춤Q1 (심화)", "맞춤Q1(심화)"), ("맞춤Q2 (규정/운영)", "맞춤Q2(규정/운영 연결)"), ("맞춤Q3 (관심/경험)", "맞춤Q3(관심/경험 기반)")
    ]
    for q_label, q_key in q_items:
        q_text = safe_str(r.get(q_key, ""))
        if q_text:
            st.markdown(f"<div class='card-question'><b>{q_label}</b><br/>{q_text}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='card-question'><b>{q_label}</b><br/><span style='color:#999;'>질문 없음</span></div>", unsafe_allow_html=True)

    st.divider()

    score_header_col, clear_col = st.columns([3, 1])
    with score_header_col:
        st.subheader("✏️ 점수 & 메모 입력")
    with clear_col:
        st.markdown("<div style='height: 0.5rem;'></div>", unsafe_allow_html=True)
        clear_key = f"clear_{candidate_id}"
        if st.button("🗑️ 입력 내용 비우기", use_container_width=True, key=clear_key, type="secondary"):
            # 위젯 key에 연결된 session_state 값을 초기값으로 직접 설정
            for k in [f"sr_{candidate_id}", f"so_{candidate_id}",
                      f"sc_{candidate_id}", f"ss_{candidate_id}",
                      f"sro_{candidate_id}", f"som_{candidate_id}"]:
                st.session_state[k] = 0
            for k in [f"fe_{candidate_id}", f"fs_{candidate_id}",
                      f"fa_{candidate_id}", f"fc_{candidate_id}", f"fo_{candidate_id}"]:
                st.session_state[k] = False
            for k in [f"ms_{candidate_id}", f"mc_{candidate_id}",
                      f"mf_{candidate_id}", f"msum_{candidate_id}"]:
                st.session_state[k] = ""
            st.session_state[f"rec_{candidate_id}"] = "미정"
            st.toast("입력 내용이 초기화되었습니다.")
            st.rerun()

    existing = evals[mask].iloc[-1].to_dict() if mask.any() else {}

    def pre_i(key, widget_key, default=0):
        """session_state에 위젯 키가 이미 있으면 그 값을 쓰고, 없으면 existing에서 초기값 세팅"""
        if widget_key in st.session_state:
            return st.session_state[widget_key]
        v = existing.get(key, default)
        try:
            if v == "" or pd.isna(v): return default
            return int(float(v))
        except Exception:
            return default

    def pre_s(key, widget_key=None, default=""):
        """session_state에 위젯 키가 이미 있으면 그 값을 쓰고, 없으면 existing에서 초기값 세팅"""
        if widget_key and widget_key in st.session_state:
            return st.session_state[widget_key]
        v = existing.get(key, default)
        return "" if pd.isna(v) else str(v)

    def pre_b(key, widget_key):
        """체크박스용: session_state에 키가 있으면 bool 반환, 없으면 existing에서 판단"""
        if widget_key in st.session_state:
            return bool(st.session_state[widget_key])
        return str(existing.get(key, "")) == "True"

    c1, c2, c3 = st.columns(3)
    with c1:
        score_rules = st.slider("규정 적합도(1~5)", 0, 5, value=pre_i("score_rules_fit", f"sr_{candidate_id}"), key=f"sr_{candidate_id}")
        score_output = st.slider("증빙/산출물 의지(1~5)", 0, 5, value=pre_i("score_output_evidence", f"so_{candidate_id}"), key=f"so_{candidate_id}")
    with c2:
        score_collab = st.slider("협업/소통(1~5)", 0, 5, value=pre_i("score_collaboration", f"sc_{candidate_id}"), key=f"sc_{candidate_id}")
        score_self = st.slider("자기주도/문제해결(1~5)", 0, 5, value=pre_i("score_self_driven", f"ss_{candidate_id}"), key=f"ss_{candidate_id}")
    with c3:
        score_role = st.slider("역할 적합/역량(1~5)", 0, 5, value=pre_i("score_role_skill", f"sro_{candidate_id}"), key=f"sro_{candidate_id}")
        score_overall_manual = st.slider("종합(직접)", 0, 5, value=pre_i("score_overall", f"som_{candidate_id}"), help="0이면 자동 평균이 들어갑니다.", key=f"som_{candidate_id}")

    avg = auto_avg([score_rules, score_output, score_collab, score_self, score_role])
    st.caption(f"자동 평균(5개): **{avg} / 5.0**")

    st.markdown("#### ⚠️ 리스크 플래그")
    f1,f2,f3,f4,f5 = st.columns(5)
    with f1: flag_evidence = st.checkbox("증빙 리스크", value=pre_b("flag_evidence_risk", f"fe_{candidate_id}"), key=f"fe_{candidate_id}")
    with f2: flag_schedule = st.checkbox("일정 리스크", value=pre_b("flag_schedule_risk", f"fs_{candidate_id}"), key=f"fs_{candidate_id}")
    with f3: flag_attitude = st.checkbox("태도 리스크", value=pre_b("flag_attitude_risk", f"fa_{candidate_id}"), key=f"fa_{candidate_id}")
    with f4: flag_comm = st.checkbox("소통 리스크", value=pre_b("flag_comm_risk", f"fc_{candidate_id}"), key=f"fc_{candidate_id}")
    with f5: flag_other = st.checkbox("기타", value=pre_b("flag_other_risk", f"fo_{candidate_id}"), key=f"fo_{candidate_id}")

    memo_strength = st.text_area("강점", value=pre_s("memo_strength", f"ms_{candidate_id}"), height=80, key=f"ms_{candidate_id}")
    memo_concern = st.text_area("우려/근거", value=pre_s("memo_concern", f"mc_{candidate_id}"), height=80, key=f"mc_{candidate_id}")
    memo_followup = st.text_area("추가 확인", value=pre_s("memo_followup", f"mf_{candidate_id}"), height=80, key=f"mf_{candidate_id}")
    memo_summary = st.text_area("요약(1~2줄)", value=pre_s("memo_summary", f"msum_{candidate_id}"), height=80, key=f"msum_{candidate_id}")

    _rec_val = pre_s("recommendation", f"rec_{candidate_id}", "미정")
    recommendation = st.selectbox(
        "추천",
        options=["합격","보류","불합","미정"],
        index=["합격","보류","불합","미정"].index(_rec_val) if _rec_val in ["합격","보류","불합","미정"] else 3,
        key=f"rec_{candidate_id}"
    )

    st.divider()
    a,b,b2,c = st.columns([1,1,1,2])

    with a:
        if st.button("💾 저장/업데이트", use_container_width=True, type="primary"):
            row_dict = {
                "timestamp": now_string(),
                "app_version": APP_VERSION,
                "interviewer": interviewer,

                "candidate_id": candidate_id,
                "name": name,
                "student_id": sid,
                "mark": mark,
                "category": cat,
                "level": lvl,

                "score_rules_fit": int(score_rules),
                "score_output_evidence": int(score_output),
                "score_collaboration": int(score_collab),
                "score_self_driven": int(score_self),
                "score_role_skill": int(score_role),
                "score_overall": (int(score_overall_manual) if score_overall_manual > 0 else avg),

                "flag_evidence_risk": str(flag_evidence),
                "flag_schedule_risk": str(flag_schedule),
                "flag_attitude_risk": str(flag_attitude),
                "flag_comm_risk": str(flag_comm),
                "flag_other_risk": str(flag_other),

                "memo_strength": memo_strength,
                "memo_concern": memo_concern,
                "memo_followup": memo_followup,
                "memo_summary": memo_summary,

                "recommendation": recommendation,
            }
            evals = upsert_eval(evals, row_dict)
            save_results(result_path, evals, candidates)
            st.success(f"저장 완료: {result_path}")

    with b:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            evals.to_excel(writer, index=False, sheet_name="Evaluations")
        st.download_button(
            "⬇️ 내 결과 엑셀 다운로드",
            data=buf.getvalue(),
            file_name=os.path.basename(result_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with b2:
        if mask.any():
            if st.button("🗑️ 이 평가 삭제", use_container_width=True, type="secondary"):
                st.session_state[f"confirm_delete_{candidate_id}"] = True

            if st.session_state.get(f"confirm_delete_{candidate_id}", False):
                st.warning(f"**{name}** 평가를 삭제할까요?")
                cd1, cd2 = st.columns(2)
                with cd1:
                    if st.button("✅ 삭제 확인", key=f"del_yes_{candidate_id}", use_container_width=True):
                        evals = evals[~mask].reset_index(drop=True)
                        save_results(result_path, evals, candidates)
                        st.session_state.pop(f"confirm_delete_{candidate_id}", None)
                        # 위젯 키도 정리
                        for wk in [f"sr_{candidate_id}", f"so_{candidate_id}",
                                   f"sc_{candidate_id}", f"ss_{candidate_id}",
                                   f"sro_{candidate_id}", f"som_{candidate_id}",
                                   f"fe_{candidate_id}", f"fs_{candidate_id}",
                                   f"fa_{candidate_id}", f"fc_{candidate_id}", f"fo_{candidate_id}",
                                   f"ms_{candidate_id}", f"mc_{candidate_id}",
                                   f"mf_{candidate_id}", f"msum_{candidate_id}",
                                   f"rec_{candidate_id}"]:
                            st.session_state.pop(wk, None)
                        st.toast("평가가 삭제되었습니다.")
                        st.rerun()
                with cd2:
                    if st.button("❌ 취소", key=f"del_no_{candidate_id}", use_container_width=True):
                        st.session_state.pop(f"confirm_delete_{candidate_id}", None)
                        st.rerun()
        else:
            st.caption("삭제할 평가 없음")

    with c:
        total = len(candidates)
        done = evals[evals["interviewer"] == interviewer]["candidate_id"].nunique() if not evals.empty else 0
        progress_pct = done / total if total > 0 else 0
        st.markdown("#### 📊 진행 현황")
        st.progress(progress_pct)
        st.markdown(f"<div style='text-align:center;'>평가 완료: <b>{done}</b> / {total} ({int(progress_pct*100)}%)</div>", unsafe_allow_html=True)

st.divider()
st.subheader("🔀 면접관 결과 병합 & 종합 대시보드")
st.caption("면접관별 결과 엑셀을 업로드하면 지원자별 평균 점수 순으로 정렬하여 한눈에 비교할 수 있습니다.")

uploads = st.file_uploader("면접 결과 엑셀 업로드(다중 선택)", type=["xlsx"], accept_multiple_files=True)
if uploads:
    merged = []
    failed_files = []
    for f in uploads:
        try:
            dfm = pd.read_excel(f, sheet_name="Evaluations")
            merged.append(dfm)
        except Exception:
            failed_files.append(f.name)
    
    if failed_files:
        st.warning(f"⚠️ 다음 파일은 Evaluations 시트를 읽지 못했습니다: {', '.join(failed_files)}")
    
    if merged:
        merged_df = pd.concat(merged, ignore_index=True)
        for col in EVAL_COLUMNS:
            if col not in merged_df.columns:
                merged_df[col] = ""
        merged_df = merged_df[EVAL_COLUMNS].copy()

        # 점수 컬럼을 숫자로 변환
        score_cols = ["score_rules_fit", "score_output_evidence", "score_collaboration", "score_self_driven", "score_role_skill", "score_overall"]
        for sc in score_cols:
            merged_df[sc] = pd.to_numeric(merged_df[sc], errors="coerce").fillna(0)

        # ---- 상단 요약 메트릭 ----
        n_interviewers = merged_df["interviewer"].nunique()
        n_candidates = merged_df["candidate_id"].nunique()
        n_evals = len(merged_df)

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("📂 업로드 파일", f"{len(uploads)}개")
        with m2:
            st.metric("👥 면접관 수", f"{n_interviewers}명")
        with m3:
            st.metric("🧑‍💼 지원자 수", f"{n_candidates}명")
        with m4:
            st.metric("📝 총 평가 수", f"{n_evals}건")

        st.divider()

        # ---- 지원자별 종합 요약 테이블 ----
        st.markdown("### 📊 지원자별 종합 순위")

        score_labels = {
            "score_rules_fit": "규정적합",
            "score_output_evidence": "증빙의지",
            "score_collaboration": "협업소통",
            "score_self_driven": "자기주도",
            "score_role_skill": "역할역량",
            "score_overall": "종합",
        }

        # 지원자별 집계
        summary_rows = []
        for cid, grp in merged_df.groupby("candidate_id"):
            row_data = {
                "이름": grp["name"].iloc[0],
                "학번": grp["student_id"].iloc[0],
                "분류": grp["category"].iloc[0],
                "레벨": grp["level"].iloc[0],
                "표시": grp["mark"].iloc[0],
                "평가 수": len(grp),
                "면접관": ", ".join(grp["interviewer"].unique()),
            }
            # 각 점수 평균
            for sc in score_cols:
                vals = grp[sc][grp[sc] > 0]
                row_data[score_labels[sc]] = round(vals.mean(), 2) if len(vals) > 0 else 0.0

            # 리스크 플래그 집계
            flag_cols = ["flag_evidence_risk", "flag_schedule_risk", "flag_attitude_risk", "flag_comm_risk", "flag_other_risk"]
            flag_labels = ["증빙", "일정", "태도", "소통", "기타"]
            flagged = []
            for fc, fl in zip(flag_cols, flag_labels):
                if (grp[fc].astype(str) == "True").any():
                    flagged.append(fl)
            row_data["리스크"] = ", ".join(flagged) if flagged else "-"

            # 추천 집계
            recs = grp["recommendation"].value_counts().to_dict()
            rec_parts = []
            for rv in ["합격", "보류", "불합", "미정"]:
                if rv in recs and recs[rv] > 0:
                    rec_parts.append(f"{rv}({recs[rv]})")
            row_data["추천"] = " / ".join(rec_parts) if rec_parts else "미정"

            # 메모 요약 합치기
            summaries = grp["memo_summary"].dropna().astype(str)
            summaries = [s for s in summaries if s.strip() and s.strip() != "nan"]
            row_data["요약"] = " | ".join(summaries) if summaries else ""

            row_data["_sort_score"] = row_data["종합"]
            summary_rows.append(row_data)

        summary_df = pd.DataFrame(summary_rows)

        # 정렬 옵션
        sort_col1, sort_col2, filter_col = st.columns([1, 1, 1])
        with sort_col1:
            sort_by = st.selectbox("정렬 기준", ["종합", "규정적합", "증빙의지", "협업소통", "자기주도", "역할역량", "이름"], index=0)
        with sort_col2:
            sort_order = st.radio("정렬 순서", ["높은 순", "낮은 순"], horizontal=True)
        with filter_col:
            filter_rec = st.multiselect("추천 필터", ["합격", "보류", "불합", "미정"], default=["합격", "보류", "불합", "미정"])

        ascending = sort_order == "낮은 순"
        summary_df = summary_df.sort_values(by=sort_by, ascending=ascending, na_position="last")

        # 추천 필터 적용
        if filter_rec:
            mask_filter = summary_df["추천"].apply(lambda x: any(r in x for r in filter_rec))
            summary_df = summary_df[mask_filter]

        # 순위 추가
        display_df = summary_df.drop(columns=["_sort_score"]).reset_index(drop=True)
        display_df.index = display_df.index + 1
        display_df.index.name = "순위"

        # 표시할 컬럼
        show_cols = ["이름", "학번", "분류", "레벨", "종합", "규정적합", "증빙의지", "협업소통", "자기주도", "역할역량", "평가 수", "면접관", "추천", "리스크", "요약"]
        show_cols = [c for c in show_cols if c in display_df.columns]

        # 점수 컬럼 하이라이트
        def highlight_scores(val):
            try:
                v = float(val)
                if v >= 4.0: return "background-color: #e6f4ea; color: #137333; font-weight: 700;"
                elif v >= 3.0: return "background-color: #fef7e0; color: #b45309; font-weight: 600;"
                elif v > 0: return "background-color: #fce8e6; color: #c5221f; font-weight: 600;"
            except (ValueError, TypeError):
                pass
            return ""

        styled_df = display_df[show_cols].style.applymap(
            highlight_scores,
            subset=[c for c in ["종합", "규정적합", "증빙의지", "협업소통", "자기주도", "역할역량"] if c in show_cols]
        ).format(
            {c: "{:.1f}" for c in ["종합", "규정적합", "증빙의지", "협업소통", "자기주도", "역할역량"] if c in show_cols}
        )

        st.dataframe(styled_df, use_container_width=True, height=min(800, 40 + len(display_df) * 38))

        st.divider()

        # ---- 면접관별 세부 비교 ----
        st.markdown("### 🔍 지원자별 면접관 세부 평가")

        if not summary_df.empty:
            cand_options = summary_df["이름"].tolist()
            selected_cand = st.selectbox("지원자 선택", cand_options, key="merge_cand_select")

            cand_row = summary_df[summary_df["이름"] == selected_cand].iloc[0]
            cand_evals = merged_df[merged_df["name"] == selected_cand]

            if not cand_evals.empty:
                st.markdown(f"""<div class='card-info'>
                    <b style='font-size:1.2em;'>{selected_cand}</b>
                    <span style='color:#666;'>({cand_row.get('학번','')})</span>
                    &nbsp;·&nbsp; 분류: {cand_row.get('분류','')}
                    &nbsp;·&nbsp; 레벨: {cand_row.get('레벨','')}
                    &nbsp;·&nbsp; 종합 평균: <b>{cand_row.get('종합', 0):.1f}</b>
                </div>""", unsafe_allow_html=True)

                for _, ev in cand_evals.iterrows():
                    interviewer_name = safe_str(ev.get("interviewer", ""))
                    rec_val = safe_str(ev.get("recommendation", ""))
                    badge_cls = {"합격": "badge-pass", "보류": "badge-hold", "불합": "badge-fail"}.get(rec_val, "badge-none")

                    scores_html = ""
                    for sc in score_cols:
                        label = score_labels[sc]
                        val = ev.get(sc, 0)
                        try:
                            val = float(val)
                        except (ValueError, TypeError):
                            val = 0.0
                        color = "#137333" if val >= 4 else ("#b45309" if val >= 3 else "#c5221f")
                        scores_html += f"<span style='margin-right:1rem;'>{label}: <b style='color:{color};'>{val:.0f}</b></span>"

                    # 플래그
                    flag_cols_ev = ["flag_evidence_risk", "flag_schedule_risk", "flag_attitude_risk", "flag_comm_risk", "flag_other_risk"]
                    flag_labels_ev = ["증빙", "일정", "태도", "소통", "기타"]
                    flags = [fl for fc, fl in zip(flag_cols_ev, flag_labels_ev) if str(ev.get(fc, "")) == "True"]
                    flag_html = f"<span style='color:#c5221f;'>⚠️ {', '.join(flags)}</span>" if flags else ""

                    memo_parts = []
                    for mk, ml in [("memo_strength", "💪 강점"), ("memo_concern", "⚠️ 우려"), ("memo_followup", "❓ 추가확인"), ("memo_summary", "📝 요약")]:
                        mv = safe_str(ev.get(mk, ""))
                        if mv:
                            memo_parts.append(f"<div style='margin:0.2rem 0;'><b>{ml}:</b> {mv}</div>")
                    memo_html = "".join(memo_parts)

                    card_html = f"""<div class='card' style='margin-bottom:0.8rem;'>
                        <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:0.6rem;'>
                            <span style='font-size:1.05em; font-weight:700;'>🧑‍💼 {interviewer_name}</span>
                            <span class='{badge_cls}'>{rec_val if rec_val else '미정'}</span>
                        </div>
                        <div style='margin-bottom:0.5rem;'>{scores_html}</div>
                        {f"<div style='margin-bottom:0.5rem;'>{flag_html}</div>" if flag_html else ""}
                        <div style='font-size:0.9em; color:#555;'>{memo_html}</div>
                        <div style='font-size:0.8em; color:#999; margin-top:0.4rem;'>🕐 {safe_str(ev.get("timestamp", ""))}</div>
                    </div>"""
                    st.markdown(card_html, unsafe_allow_html=True)

        st.divider()

        # ---- 다운로드 ----
        st.markdown("### 💾 병합 결과 다운로드")
        dl1, dl2 = st.columns(2)
        with dl1:
            # 원본 병합 (면접관별 행)
            out_raw = io.BytesIO()
            with pd.ExcelWriter(out_raw, engine="openpyxl") as writer:
                merged_df.to_excel(writer, index=False, sheet_name="MergedEvaluations")
                # 요약 시트도 추가
                summary_export = summary_df.drop(columns=["_sort_score"], errors="ignore")
                summary_export.to_excel(writer, index=False, sheet_name="Summary")
            st.download_button(
                "⬇️ 전체 병합 엑셀 (원본+요약)",
                data=out_raw.getvalue(),
                file_name="merged_interview_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl2:
            # 요약만
            out_sum = io.BytesIO()
            summary_export2 = summary_df.drop(columns=["_sort_score"], errors="ignore")
            summary_export2.to_excel(out_sum, index=False, sheet_name="Summary")
            st.download_button(
                "⬇️ 요약 순위표만 다운로드",
                data=out_sum.getvalue(),
                file_name="interview_summary_ranking.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.error("업로드한 파일에서 Evaluations 시트를 읽지 못했습니다.")