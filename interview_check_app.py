import os
import io
import time
from datetime import datetime
import pandas as pd
import streamlit as st

APP_VERSION = "1.4.0"

DEFAULT_INPUT_FILE = "Stolio_5기_면접질문.xlsx"
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
st.set_page_config(page_title="Stolio 면접 체크", layout="wide")
st.title("Stolio 면접 체크 프로그램")
st.caption(f"v{APP_VERSION} · 쉬운질문/검색/정렬/점수저장/병합/타이머")

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

# -------- Layout (stable) --------

# [수정 포인트] 데이터를 먼저 확정 짓고, 그 다음에 화면(left/right)을 나눕니다.
# 이렇게 하면 'left' 안에서 변수가 갇히는 문제를 원천 차단합니다.

st.subheader("지원자 선택")
selected_label = st.selectbox("지원자", labels, index=0)
row_idx = label_to_index[selected_label]
r = candidates.loc[row_idx]

# --- 변수 정의 (여기서 미리 다 뽑아둡니다) ---
candidate_id = safe_str(r.get("_candidate_id",""))
name = safe_str(r.get("이름",""))
sid  = safe_str(r.get("학번",""))
mark = safe_str(r.get("학번표시",""))
cat  = safe_str(r.get("분류",""))
lvl  = safe_str(r.get("예상레벨","")) # ★ 이제 이 변수는 전역에서 안전합니다
dup  = safe_str(r.get("중복지원",""))
# ----------------------------------------

left, right = st.columns([1,2], gap="large")

with left:
    # (위에서 이미 데이터를 뽑았으므로 여기선 보여주기만 합니다)
    st.markdown("#### 기본 정보")
    st.write(f"- 표시: **{mark}**")
    st.write(f"- 이름/학번: **{name} ({sid})**")
    if cat: st.write(f"- 분류: {cat}")
    if lvl: st.write(f"- 예상레벨: {lvl}")
    if dup: st.write(f"- 중복지원: {dup}")

    st.divider()

    # -------- Timer (robust, no external deps) --------
    if enable_timer:
        # (타이머 코드는 그대로 유지)
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
        st.write(f"남은 시간: **{int(remaining//60)}:{int(remaining%60):02d}** · 상태: **{'진행중' if st.session_state[k_running] else '일시정지'}**")
        if remaining <= 0:
            st.error("⏰ 면접 시간이 종료되었습니다. (원하면 리셋하세요)")

        if live_timer and st.session_state[k_running] and remaining > 0:
            time.sleep(1)
            st.rerun()

    st.divider()

    # -------- Existing evaluation preview --------
    # (평가 미리보기 코드 그대로 유지)
    mask = (evals["interviewer"] == interviewer) & (evals["candidate_id"] == candidate_id)
    if mask.any():
        last = evals[mask].iloc[-1]
        st.markdown("#### 저장된 평가(이 면접관 기준)")
        st.write(f"- 저장 시각: {safe_str(last.get('timestamp',''))}")
        st.write(f"- 종합: {safe_str(last.get('score_overall',''))}")
        st.write(f"- 추천: {safe_str(last.get('recommendation',''))}")
        summ = safe_str(last.get("memo_summary",""))
        if summ:
            st.write(f"- 요약: {summ}")
    else:
        st.caption("아직 저장된 평가가 없습니다.")

with right:
    st.subheader("지원서 답변 & 면접 질문(쉬운 버전)")

    st.markdown("#### 지원서 답변")
    st.markdown("**지원동기**")
    st.write(safe_str(r.get("지원서답변1(동기)","")))
    st.markdown("**기대/매력**")
    st.write(safe_str(r.get("지원서답변2(기대/매력)","")))
    st.markdown("**관심/경험**")
    st.write(safe_str(r.get("지원서답변3(관심/경험)","")))

    st.markdown("#### 면접 질문")
    st.markdown("**공통Q1**"); st.write(safe_str(r.get("공통Q1","")))
    st.markdown("**공통Q2**"); st.write(safe_str(r.get("공통Q2","")))
    st.markdown("**공통Q3**"); st.write(safe_str(r.get("공통Q3","")))
    st.markdown("**맞춤Q1**"); st.write(safe_str(r.get("맞춤Q1(심화)","")))
    st.markdown("**맞춤Q2**"); st.write(safe_str(r.get("맞춤Q2(규정/운영 연결)","")))
    st.markdown("**맞춤Q3**"); st.write(safe_str(r.get("맞춤Q3(관심/경험 기반)","")))

    st.divider()
    st.subheader("점수 & 메모 입력")

    existing = evals[mask].iloc[-1].to_dict() if mask.any() else {}

    def pre_i(key, default=0):
        v = existing.get(key, default)
        try:
            if v == "" or pd.isna(v): return default
            return int(float(v))
        except Exception:
            return default

    def pre_s(key, default=""):
        v = existing.get(key, default)
        return "" if pd.isna(v) else str(v)

    c1, c2, c3 = st.columns(3)
    with c1:
        score_rules = st.slider("규정 적합도(1~5)", 0, 5, value=pre_i("score_rules_fit"))
        score_output = st.slider("증빙/산출물 의지(1~5)", 0, 5, value=pre_i("score_output_evidence"))
    with c2:
        score_collab = st.slider("협업/소통(1~5)", 0, 5, value=pre_i("score_collaboration"))
        score_self = st.slider("자기주도/문제해결(1~5)", 0, 5, value=pre_i("score_self_driven"))
    with c3:
        score_role = st.slider("역할 적합/역량(1~5)", 0, 5, value=pre_i("score_role_skill"))
        score_overall_manual = st.slider("종합(직접)", 0, 5, value=pre_i("score_overall"), help="0이면 자동 평균이 들어갑니다.")

    avg = auto_avg([score_rules, score_output, score_collab, score_self, score_role])
    st.caption(f"자동 평균(5개): **{avg} / 5.0**")

    st.markdown("#### 리스크 플래그")
    f1,f2,f3,f4,f5 = st.columns(5)
    with f1: flag_evidence = st.checkbox("증빙 리스크", value=(pre_s("flag_evidence_risk")=="True"))
    with f2: flag_schedule = st.checkbox("일정 리스크", value=(pre_s("flag_schedule_risk")=="True"))
    with f3: flag_attitude = st.checkbox("태도 리스크", value=(pre_s("flag_attitude_risk")=="True"))
    with f4: flag_comm = st.checkbox("소통 리스크", value=(pre_s("flag_comm_risk")=="True"))
    with f5: flag_other = st.checkbox("기타", value=(pre_s("flag_other_risk")=="True"))

    memo_strength = st.text_area("강점", value=pre_s("memo_strength"), height=80)
    memo_concern = st.text_area("우려/근거", value=pre_s("memo_concern"), height=80)
    memo_followup = st.text_area("추가 확인", value=pre_s("memo_followup"), height=80)
    memo_summary = st.text_area("요약(1~2줄)", value=pre_s("memo_summary"), height=80)

    recommendation = st.selectbox(
        "추천",
        options=["합격","보류","불합","미정"],
        index=["합격","보류","불합","미정"].index(pre_s("recommendation","미정")) if pre_s("recommendation","미정") in ["합격","보류","불합","미정"] else 3
    )

    st.divider()
    a,b,c = st.columns([1,1,2])

    with a:
        if st.button("💾 저장/업데이트", use_container_width=True):
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

    with c:
        total = len(candidates)
        done = evals[evals["interviewer"] == interviewer]["candidate_id"].nunique() if not evals.empty else 0
        st.markdown("#### 진행 현황")
        st.write(f"- 평가 완료: **{done} / {total}**")

st.divider()
st.subheader("면접관 결과 병합(선택)")
st.caption("면접관별 결과 엑셀을 업로드하면 하나로 합쳐서 다운로드할 수 있습니다.")

uploads = st.file_uploader("면접 결과 엑셀 업로드(다중 선택)", type=["xlsx"], accept_multiple_files=True)
if uploads:
    merged = []
    for f in uploads:
        try:
            dfm = pd.read_excel(f, sheet_name="Evaluations")
            merged.append(dfm)
        except Exception:
            pass
    if merged:
        merged_df = pd.concat(merged, ignore_index=True)
        for col in EVAL_COLUMNS:
            if col not in merged_df.columns:
                merged_df[col] = ""
        merged_df = merged_df[EVAL_COLUMNS]
        st.success(f"병합 완료: {len(merged_df)} rows")
        st.dataframe(merged_df.head(50), use_container_width=True)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="MergedEvaluations")
        st.download_button(
            "⬇️ 병합 엑셀 다운로드",
            data=out.getvalue(),
            file_name="merged_interview_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("업로드한 파일에서 Evaluations 시트를 읽지 못했습니다.")