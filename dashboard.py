"""
[한솔테크닉스 HEVH] LOSSTIME 분석 대시보드 v3.1
- 날짜 파싱 수정 (베트남식 DD.MM)
- DAY/NIGHT 파일명 인식 수정
실행: python -m streamlit run dashboard.py
"""

import re
import io
import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
from openpyxl.chartsheet import Chartsheet

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="HEVH LOSSTIME 대시보드",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.main-header {
    background: linear-gradient(90deg, #1e3a5f, #2563eb);
    padding: 20px; border-radius: 10px; margin-bottom: 20px;
    color: white; text-align: center;
}
.metric-box {
    background: #f8fafc; border: 1px solid #e2e8f0;
    border-radius: 8px; padding: 16px; text-align: center;
    margin-bottom: 8px;
}
.metric-val { font-size: 32px; font-weight: bold; color: #1e3a5f; }
.metric-lbl { font-size: 13px; color: #64748b; margin-top: 4px; }
.pm-p1 { background:#fef2f2; border-left:4px solid #dc2626;
          padding:10px; border-radius:4px; margin:4px 0; }
.pm-p2 { background:#fff7ed; border-left:4px solid #f97316;
          padding:10px; border-radius:4px; margin:4px 0; }
.pm-p3 { background:#fefce8; border-left:4px solid #eab308;
          padding:10px; border-radius:4px; margin:4px 0; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 상수
# ─────────────────────────────────────────────
DAY_SLOTS   = ["A","B","C","D","E","F"]
NIGHT_SLOTS = ["F","G","H","I","J","K"]
DB_PATH     = "losstime_db.csv"

LOSS_TYPE_RULES = [
    ("NEW_OP",       "신규OP교육",         ["new op","đào tạo","op mới"]),
    ("MATERIAL",     "자재부족",           ["thiếu","magazine","stock","chờ liệu","đợi liệu"]),
    ("WAITING_SMT",  "SMT대기",            ["waiting semi","đợi semi","chờ semi"]),
    ("CHANGE_MODEL", "모델교체",           ["change model","đổi model","tách lot"]),
    ("PRINTER",      "Printer불량",        ["printer","priter","lỗi keo","tràn keo"]),
    ("MOUTER",       "Mouter불량",         ["moutor","mounter","stopper","băng tải"]),
    ("INSERT_RH3",   "RH3삽입불량",        ["rh3","rh 3","rhu"]),
    ("INSERT_XGZ",   "XGZ불량",            ["xzg","xgz"]),
    ("INSERT_AXIAL", "Axial불량",          ["axial","radial","cắm rớt","rad lỗi"]),
    ("INSERT_JUMP",  "Jumper불량",         ["jumper","jump "]),
    ("COATING",      "Coating/Reflow불량", ["coating","reflow","flux","nhiệt"]),
    ("AOI_SAOI",     "AOI/S-AOI불량",      ["saoi","s-aoi","aoi"]),
    ("WAVE",         "Wave Solder불량",    ["wave solder","wave"]),
    ("SPI",          "SPI불량",            ["spi"]),
    ("INLOADER",     "Inloader불량",       ["inloader"]),
    ("ICT_ATE",      "ICT/ATE불량",        ["ict","ate","ft lỗi"]),
    ("EQUIP_FAIL",   "설비고장(기타)",     ["hư","lỗi máy","stop line","spare"]),
    ("DONE_PLAN",    "계획완료",           ["done plan","hết plan","kết thúc"]),
    ("SAMPLE",       "샘플/테스트",        ["sample"]),
    ("ETC",          "기타",               []),
]

PROC_COLOR = {"AI":"#ef4444","SMT":"#3b82f6","MI":"#10b981"}

# ─────────────────────────────────────────────
# 로그인
# ─────────────────────────────────────────────
def check_login():
    try:
        return st.secrets["users"]
    except Exception:
        return {
            "hansol":  "hevh2024",
            "vietnam": "hevh2024",
            "korea":   "hevh2024",
        }

def login_page():
    st.markdown("""
    <div style="text-align:center; margin-top:60px;">
        <h2>🏭 한솔테크닉스 HEVH</h2>
        <h4 style="color:#64748b;">LOSSTIME 분석 대시보드</h4>
    </div>
    """, unsafe_allow_html=True)

    col1,col2,col3 = st.columns([1,1.2,1])
    with col2:
        with st.form("login_form"):
            st.markdown("#### 🔐 로그인")
            username = st.text_input("아이디", placeholder="ID 입력")
            password = st.text_input("비밀번호", type="password",
                                     placeholder="PW 입력")
            submitted = st.form_submit_button(
                "로그인", use_container_width=True, type="primary")
            if submitted:
                users = check_login()
                if username in users and users[username] == password:
                    st.session_state["logged_in"] = True
                    st.session_state["username"]  = username
                    st.rerun()
                else:
                    st.error("❌ 아이디 또는 비밀번호가 틀렸습니다.")

# ─────────────────────────────────────────────
# 파싱 엔진
# ─────────────────────────────────────────────
def classify_loss_type(text):
    if not text: return ("ETC","기타")
    t = str(text).lower()
    for code,name,kws in LOSS_TYPE_RULES:
        for kw in kws:
            if kw in t: return (code,name)
    return ("ETC","기타")

def extract_cause_lines(cause_row):
    if not cause_row: return []
    lines = []
    for cell in list(cause_row)[2:9]:
        if not cell: continue
        text = str(cell).strip()
        if not text or text.lower() in ["none","—","-",""]: continue
        sub = [l.strip() for l in re.split(r'[\n\r]+', text) if l.strip()]
        lines.extend(sub)
    seen=set(); unique=[]
    for l in lines:
        if l not in seen: seen.add(l); unique.append(l)
    return unique

def match_cause_to_slots(loss_vals, cause_lines, slots):
    result = {s:"" for s in slots}
    if not cause_lines: return result
    loss_slots = [slots[i] for i,v in enumerate(loss_vals)
                  if i < len(slots) and v > 0]
    if len(cause_lines) == 1:
        for s in slots: result[s] = cause_lines[0]
    else:
        for idx,slot in enumerate(loss_slots):
            if idx < len(cause_lines)-1:
                result[slot] = cause_lines[idx]
            else:
                result[slot] = " | ".join(cause_lines[idx:]); break
    return result

def extract_model_per_slot(model_row, slots):
    result = {s:"" for s in slots}
    if not model_row: return result
    cells = list(model_row)[2:2+len(slots)]
    for idx,cell in enumerate(cells):
        if idx >= len(slots): break
        v = str(cell or "").strip()
        mm = re.search(r'(L\d{2}[A-Z0-9_\-\.]+)', v, re.I)
        if mm: result[slots[idx]] = mm.group(1)
        elif v and v.lower() not in ["none","—","-",""]:
            result[slots[idx]] = v
    last=""
    for s in slots:
        if result[s]: last=result[s]
        elif last: result[s]=last
    return result

def parse_losstime(val):
    if val is None: return 0.0
    m = re.search(r'(\d+\.?\d*)', str(val))
    return float(m.group(1)) if m else 0.0

def is_line_cell(val):
    return bool(re.match(
        r'^(SA\s*\d+|PA\s*\d+|PS\s*\d+|MI\s*\d+)',
        str(val or "").strip(), re.I))

def get_label(row):
    return str(row[1] or "").strip().upper() if len(row)>1 else ""

# ★ 수정된 함수들
def detect_process(fn):
    fu = fn.upper()
    if "AI REPORT" in fu:  return "AI"
    if "SMD REPORT" in fu: return "SMT"
    if "MI TIME" in fu:    return "MI"
    return "UNKNOWN"

def parse_date(text):
    if not text: return None
    t = str(text).strip()

    # DAY/NIGHT 제거
    t = re.sub(r'\b(DAY|NIGHT)\b', '', t, flags=re.I).strip()

    # ★ DD.MM 또는 DD.M (베트남식) → 2026-MM-DD
    m = re.search(r'\b(\d{1,2})\.(\d{1,2})\b', t)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if 1 <= a <= 31 and 1 <= b <= 12:
            return f"2026-{b:02d}-{a:02d}"

    # DD/MM 형식
    m = re.search(r'(\d{1,2})/(\d{1,2})', t)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if 1 <= a <= 31 and 1 <= b <= 12:
            return f"2026-{b:02d}-{a:02d}"

    return None

def detect_shift(sn, fn):
    combined = (sn + fn).upper()
    if "NIGHT" in combined: return "NIGHT"
    return "DAY"

def normalize_line(raw):
    s = str(raw).strip()
    s = re.sub(r':.*$','',s).strip()
    s = re.sub(r'(?i)^PA\s*0*(\d+)', lambda m: f"SA{int(m.group(1)):02d}", s)
    s = re.sub(r'(?i)^SA\s*0*(\d+)', lambda m: f"SA{int(m.group(1)):02d}", s)
    s = re.sub(r'(?i)^PS\s*0*(\d+)', lambda m: f"PS{int(m.group(1)):02d}", s)
    s = re.sub(r'(?i)^MI\s*0*(\d+)', lambda m: f"MI {m.group(1)}", s)
    return s.strip()

def parse_sheet(ws, process, date_str, shift):
    records = []
    rows  = list(ws.iter_rows(values_only=True))
    slots = NIGHT_SLOTS if shift=="NIGHT" else DAY_SLOTS
    i = 0
    while i < len(rows):
        row = rows[i]
        c1  = row[1] if len(row)>1 else None
        if not is_line_cell(c1): i+=1; continue

        line = normalize_line(str(c1))
        model_row=loss_row=cause_row=action_row=None

        for j in range(i+1, min(i+14,len(rows))):
            r = rows[j]; lbl = get_label(r)
            if is_line_cell(r[1] if len(r)>1 else None) and j>i+1: break
            if "MODEL" in lbl and model_row is None: model_row=r
            if "LOSSTIME" in lbl and loss_row is None: loss_row=r
            elif lbl=="CAUSE" and cause_row is None: cause_row=r
            elif lbl=="ACTION" and action_row is None: action_row=r

        if loss_row:
            loss_vals=[]
            for c in range(2, 2+len(slots)):
                try:
                    v = loss_row[c] if c<len(loss_row) else None
                    s = str(v or "")
                    mm = re.search(r'(\d+)\s*min', s, re.I)
                    loss_vals.append(
                        float(mm.group(1)) if mm else parse_losstime(v))
                except: loss_vals.append(0.0)
            while len(loss_vals)<len(slots): loss_vals.append(0.0)

            total     = sum(loss_vals)
            models    = extract_model_per_slot(model_row, slots)
            cause_lns = extract_cause_lines(cause_row)
            slot_cause= match_cause_to_slots(loss_vals, cause_lns, slots)
            cause_all = " | ".join(cause_lns) if cause_lns else ""
            complexity= "복합" if len(cause_lns)>1 else "단일"
            action    = " | ".join(
                str(v) for v in (list(action_row[2:9]) if action_row else [])
                if v and str(v).strip())

            for idx,slot in enumerate(slots):
                lv = loss_vals[idx] if idx<len(loss_vals) else 0.0
                if lv>0:
                    cs = slot_cause.get(slot, cause_all)
                    code,name = classify_loss_type(cs)
                    records.append({
                        "date":date_str,"shift":shift,"process":process,
                        "line":line,"time_slot":slot,
                        "model":models.get(slot,""),
                        "loss_min":lv,"loss_type_code":code,
                        "loss_type_name":name,"complexity":complexity,
                        "loss_detail":cs,"action":action
                    })
            if total>0:
                code_t,name_t = classify_loss_type(cause_all)
                records.append({
                    "date":date_str,"shift":shift,"process":process,
                    "line":line,"time_slot":"TOTAL",
                    "model":models.get(slots[0],""),
                    "loss_min":total,"loss_type_code":code_t,
                    "loss_type_name":name_t,"complexity":complexity,
                    "loss_detail":cause_all,"action":action
                })
        i+=1
    return records

def parse_files(uploaded_files):
    all_records=[]
    prog   = st.progress(0)
    status = st.empty()
    for fi, uf in enumerate(uploaded_files):
        fn      = uf.name
        process = detect_process(fn)
        if process=="UNKNOWN":
            status.warning(f"⏭️ 스킵: {fn}")
            continue
        file_date = parse_date(fn) or "UNKNOWN"
        status.info(f"📂 처리 중: {fn} → {process} / {file_date}")
        try:
            wb = load_workbook(uf, data_only=True)
        except Exception as e:
            status.error(f"열기 실패: {fn} — {e}")
            continue
        for sn in wb.sheetnames:
            ws = wb[sn]
            if isinstance(ws, Chartsheet): continue
            shift    = detect_shift(sn, fn)
            date_str = file_date
            if date_str=="UNKNOWN":
                date_str = parse_date(sn) or "UNKNOWN"
            try:
                recs = parse_sheet(ws, process, date_str, shift)
                all_records.extend(recs)
            except Exception as e:
                st.warning(f"파싱 오류 [{sn}]: {e}")
        prog.progress((fi+1)/len(uploaded_files))
    status.empty(); prog.empty()
    df = pd.DataFrame(all_records)
    if not df.empty and "date" in df.columns:
        df = df[df["date"] != "UNKNOWN"].reset_index(drop=True)
    return df

# ─────────────────────────────────────────────
# DB 관리
# ─────────────────────────────────────────────
def load_db():
    try:
        df = pd.read_csv(DB_PATH, encoding="utf-8-sig")
        if "date" in df.columns:
            df = df[df["date"] != "UNKNOWN"].reset_index(drop=True)
        return df
    except FileNotFoundError:
        return pd.DataFrame()

def save_db(df):
    df.to_csv(DB_PATH, index=False, encoding="utf-8-sig")

def merge_db(existing, new_df):
    if existing.empty: return new_df
    if new_df.empty:   return existing
    combined = pd.concat([existing, new_df], ignore_index=True)
    key_cols = ["date","shift","process","line","time_slot"]
    combined = combined.drop_duplicates(
        subset=[c for c in key_cols if c in combined.columns],
        keep="last")
    return combined.sort_values(
        ["date","shift","process","line","time_slot"]
    ).reset_index(drop=True)

# ─────────────────────────────────────────────
# 엑셀 다운로드
# ─────────────────────────────────────────────
def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        total_df = df[df["time_slot"]=="TOTAL"] \
                   if "time_slot" in df.columns else df

        line_sum = (total_df.groupby(["process","line"])["loss_min"]
                    .sum().reset_index()
                    .sort_values("loss_min",ascending=False)
                    .rename(columns={"loss_min":"누계손실(분)"}))
        line_sum["누계(시간)"] = (line_sum["누계손실(분)"]/60).round(1)
        line_sum.insert(0,"순위",range(1,len(line_sum)+1))
        line_sum.to_excel(writer, sheet_name="라인별누계", index=False)

        type_sum = (total_df.groupby("loss_type_name")["loss_min"]
                    .agg(["sum","count"]).reset_index()
                    .sort_values("sum",ascending=False)
                    .rename(columns={"loss_type_name":"유형",
                                     "sum":"손실(분)","count":"건수"}))
        type_sum.to_excel(writer, sheet_name="손실유형별", index=False)

        date_t = (total_df.groupby(["date","process"])["loss_min"]
                  .sum().unstack(fill_value=0).reset_index())
        date_t.to_excel(writer, sheet_name="날짜별트렌드", index=False)
        df.to_excel(writer, sheet_name="원본데이터", index=False)
    return buf.getvalue()

# ─────────────────────────────────────────────
# 메인 대시보드
# ─────────────────────────────────────────────
def dashboard():
    st.markdown("""
    <div class="main-header">
        <h2>🏭 한솔테크닉스 HEVH — LOSSTIME 분석 대시보드</h2>
        <p style="margin:0;opacity:0.85">AI / SMT / PBA(MI) 공정 | 호치민 법인</p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown(f"👤 **{st.session_state.get('username','')}** 님")
        if st.button("🚪 로그아웃", use_container_width=True):
            st.session_state["logged_in"] = False
            st.rerun()

        st.divider()
        st.header("📂 파일 업로드")
        uploaded = st.file_uploader(
            "xlsx 파일 (여러 개 가능)",
            type=["xlsx"],
            accept_multiple_files=True
        )
        if uploaded:
            if st.button("🚀 분석 시작 / DB 누적",
                         type="primary", use_container_width=True):
                with st.spinner("파싱 중..."):
                    new_df = parse_files(uploaded)
                if not new_df.empty:
                    existing = load_db()
                    merged   = merge_db(existing, new_df)
                    save_db(merged)
                    st.session_state["df"] = merged
                    st.success(
                        f"✅ {len(new_df):,}건 추가 → 누계 {len(merged):,}건")
                else:
                    st.error("파싱된 데이터가 없습니다.")

        st.divider()
        if st.button("💾 저장된 DB 불러오기", use_container_width=True):
            db = load_db()
            if not db.empty:
                st.session_state["df"] = db
                st.success(f"✅ {len(db):,}건 로드")
            else:
                st.warning("저장된 DB가 없습니다.")

        st.divider()
        st.header("🔍 조회 필터")
        df_all = st.session_state.get("df", pd.DataFrame())

        if df_all.empty:
            st.info("데이터를 먼저 불러오세요.")
            return

        dates = sorted([d for d in df_all["date"].dropna().unique()
                        if d != "UNKNOWN"])
        if not dates:
            st.warning("유효한 날짜 없음")
            return

        if len(dates) >= 2:
            date_range = st.select_slider(
                "날짜 범위", options=dates,
                value=(dates[0], dates[-1]))
        else:
            date_range = (dates[0], dates[0])
            st.info(f"날짜: {dates[0]}")

        procs  = st.multiselect("공정", ["AI","SMT","MI"],
                                default=["AI","SMT","MI"])
        shifts = st.multiselect("주야간", ["DAY","NIGHT"],
                                default=["DAY","NIGHT"])
        lines_all = sorted(df_all["line"].dropna().unique())
        sel_lines = st.multiselect("라인 (미선택=전체)", lines_all)
        types_all = sorted(df_all["loss_type_name"].dropna().unique())
        sel_types = st.multiselect("손실유형 (미선택=전체)", types_all)
        view_mode = st.radio(
            "조회 단위",
            ["TOTAL(일계)","타임별(A~K)"],
            horizontal=True)

    # ── 필터 적용 ──
    df = st.session_state.get("df", pd.DataFrame())
    if df.empty:
        st.info("👈 파일을 업로드하거나 DB를 불러오세요.")
        return

    mask = (
        (df["date"] >= date_range[0]) &
        (df["date"] <= date_range[1]) &
        (df["shift"].isin(shifts or ["DAY","NIGHT"])) &
        (df["process"].isin(procs or ["AI","SMT","MI"]))
    )
    if sel_lines: mask &= df["line"].isin(sel_lines)
    if sel_types: mask &= df["loss_type_name"].isin(sel_types)

    if view_mode == "TOTAL(일계)":
        mask &= (df["time_slot"] == "TOTAL")
    else:
        mask &= (df["time_slot"] != "TOTAL")

    fdf      = df[mask].copy()
    total_df = df[mask & (df["time_slot"]=="TOTAL")].copy()

    if fdf.empty:
        st.warning("⚠️ 조건에 맞는 데이터가 없습니다.")
        return

    # ── KPI ──
    total_min = int(total_df["loss_min"].sum()) if not total_df.empty else 0
    total_hr  = total_min // 60
    n_lines   = total_df["line"].nunique() if not total_df.empty else 0
    n_days    = fdf["date"].nunique()

    if not total_df.empty:
        ws_line    = total_df.groupby("line")["loss_min"].sum()
        worst_line = ws_line.idxmax()
        worst_min  = int(ws_line.max())
    else:
        worst_line, worst_min = "—", 0

    c1,c2,c3,c4 = st.columns(4)
    for col, val, lbl in [
        (c1, f"{total_min:,}분", f"총 손실 ({total_hr}시간)"),
        (c2, f"{n_lines}개",     "분석 라인 수"),
        (c3, f"{n_days}일",      "분석 일수"),
        (c4, worst_line,         f"최대 손실 ({worst_min:,}분)"),
    ]:
        col.markdown(f"""
        <div class="metric-box">
            <div class="metric-val">{val}</div>
            <div class="metric-lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    st.divider()

    # ── 탭 ──
    tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs([
        "📊 손실 분석","📈 트렌드","🕐 타임별 분석",
        "🔍 상세 조회","🔧 설비보전 PM","⬇️ 다운로드"
    ])

    # ── TAB1 ──
    with tab1:
        if total_df.empty:
            st.warning("TOTAL 데이터 없음")
        else:
            col_l,col_r = st.columns(2)
            with col_l:
                st.subheader("손실유형별 누계")
                ts = (total_df.groupby("loss_type_name")["loss_min"]
                      .sum().reset_index()
                      .sort_values("loss_min",ascending=True)
                      .rename(columns={"loss_type_name":"유형",
                                       "loss_min":"손실(분)"}))
                fig = px.bar(ts, x="손실(분)", y="유형",
                             orientation="h", color="손실(분)",
                             color_continuous_scale="Reds", height=450)
                fig.update_layout(showlegend=False,
                                  margin=dict(l=0,r=0,t=20,b=0))
                st.plotly_chart(fig, use_container_width=True)

            with col_r:
                st.subheader("라인별 누계 TOP 15")
                ls = (total_df.groupby(["process","line"])["loss_min"]
                      .sum().reset_index()
                      .sort_values("loss_min",ascending=False).head(15))
                fig2 = px.bar(ls, x="line", y="loss_min",
                              color="process",
                              color_discrete_map=PROC_COLOR, height=450,
                              labels={"loss_min":"손실(분)","line":"라인"})
                fig2.update_layout(margin=dict(l=0,r=0,t=20,b=0))
                st.plotly_chart(fig2, use_container_width=True)

            st.subheader("공정별 비중")
            ps = total_df.groupby("process")["loss_min"].sum().reset_index()
            col_pie,_ = st.columns([1,2])
            fig3 = px.pie(ps, values="loss_min", names="process",
                          color="process",
                          color_discrete_map=PROC_COLOR, height=320)
            fig3.update_layout(margin=dict(l=0,r=0,t=20,b=0))
            col_pie.plotly_chart(fig3, use_container_width=True)

    # ── TAB2 ──
    with tab2:
        if total_df.empty:
            st.warning("TOTAL 데이터 없음")
        else:
            st.subheader("날짜별 손실 트렌드")
            dt = (total_df.groupby(["date","process"])["loss_min"]
                  .sum().reset_index())
            fig4 = px.line(dt, x="date", y="loss_min",
                           color="process",
                           color_discrete_map=PROC_COLOR,
                           markers=True, height=380,
                           labels={"loss_min":"손실(분)","date":"날짜"})
            fig4.update_layout(margin=dict(l=0,r=0,t=20,b=0))
            st.plotly_chart(fig4, use_container_width=True)

            st.subheader("DAY vs NIGHT 비교")
            sc = (total_df.groupby(["shift","loss_type_name"])["loss_min"]
                  .sum().reset_index())
            fig5 = px.bar(sc, x="loss_type_name", y="loss_min",
                          color="shift",
                          color_discrete_map={
                              "DAY":"#f59e0b","NIGHT":"#6366f1"},
                          barmode="group", height=380,
                          labels={"loss_min":"손실(분)",
                                  "loss_type_name":"손실유형"})
            fig5.update_layout(margin=dict(l=0,r=0,t=20,b=0),
                               xaxis_tickangle=-30)
            st.plotly_chart(fig5, use_container_width=True)

    # ── TAB3: 타임별 ──
    with tab3:
        st.subheader("🕐 타임별 손실 히트맵")
        slot_df = df[
            (df["date"] >= date_range[0]) &
            (df["date"] <= date_range[1]) &
            (df["shift"].isin(shifts or ["DAY","NIGHT"])) &
            (df["process"].isin(procs or ["AI","SMT","MI"])) &
            (df["time_slot"] != "TOTAL")
        ].copy()
        if sel_lines: slot_df = slot_df[slot_df["line"].isin(sel_lines)]

        if slot_df.empty:
            st.warning("타임별 데이터 없음")
        else:
            slot_order = ["A","B","C","D","E","F","G","H","I","J","K"]
            pivot = (slot_df.groupby(["line","time_slot"])["loss_min"]
                     .sum().reset_index())
            pivot_table = pivot.pivot(
                index="line", columns="time_slot", values="loss_min"
            ).fillna(0)
            cols_sorted = [c for c in slot_order if c in pivot_table.columns]
            pivot_table = pivot_table[cols_sorted]

            fig_heat = px.imshow(
                pivot_table,
                color_continuous_scale="Reds",
                aspect="auto",
                height=max(400, len(pivot_table)*30),
                labels={"x":"시간대","y":"라인","color":"손실(분)"},
                title="라인 × 시간대 손실 히트맵"
            )
            fig_heat.update_layout(margin=dict(l=0,r=0,t=40,b=0))
            st.plotly_chart(fig_heat, use_container_width=True)

            st.subheader("시간대별 손실 합계")
            slot_sum = (slot_df.groupby(["time_slot","process"])["loss_min"]
                        .sum().reset_index())
            slot_sum["time_slot"] = pd.Categorical(
                slot_sum["time_slot"], categories=slot_order, ordered=True)
            slot_sum = slot_sum.sort_values("time_slot")
            fig_slot = px.bar(
                slot_sum, x="time_slot", y="loss_min",
                color="process", color_discrete_map=PROC_COLOR,
                barmode="stack", height=350,
                labels={"loss_min":"손실(분)","time_slot":"시간대"})
            fig_slot.update_layout(margin=dict(l=0,r=0,t=20,b=0))
            st.plotly_chart(fig_slot, use_container_width=True)

            st.subheader("타임별 상세")
            slot_detail = (slot_df.groupby(
                ["date","shift","line","time_slot","loss_type_name"])
                ["loss_min"].sum().reset_index()
                .sort_values(["date","line","time_slot"]))
            st.dataframe(slot_detail, use_container_width=True, height=400)

    # ── TAB4 ──
    with tab4:
        st.subheader("상세 데이터 조회")
        search = st.text_input("🔍 키워드 (라인/원인/모델)")
        show_df = fdf.copy()
        if search:
            m = (
                show_df["line"].astype(str).str.contains(
                    search, case=False, na=False) |
                show_df["loss_detail"].astype(str).str.contains(
                    search, case=False, na=False) |
                show_df["model"].astype(str).str.contains(
                    search, case=False, na=False)
            )
            show_df = show_df[m]
        st.caption(f"총 {len(show_df):,}건")
        disp = ["date","shift","process","line","time_slot",
                "model","loss_min","loss_type_name","complexity","loss_detail"]
        disp = [c for c in disp if c in show_df.columns]
        if not show_df.empty:
            st.dataframe(show_df[disp].reset_index(drop=True),
                         use_container_width=True, height=500)
        else:
            st.info("검색 결과 없음")

    # ── TAB5 ──
    with tab5:
        st.subheader("🔧 설비보전 PM 우선순위")
        if total_df.empty:
            st.warning("데이터 없음")
        else:
            pm_types = [
                "Printer불량","Axial불량","RH3삽입불량",
                "Mouter불량","XGZ불량","설비고장(기타)",
                "Coating/Reflow불량","AOI/S-AOI불량","Wave Solder불량"
            ]
            pm_df = (total_df[total_df["loss_type_name"].isin(pm_types)]
                     .groupby(["process","line","loss_type_name"])["loss_min"]
                     .agg(["sum","count"]).reset_index()
                     .sort_values("sum",ascending=False)
                     .rename(columns={"sum":"누계손실(분)","count":"발생횟수"}))

            def pm_grade(row):
                if row["발생횟수"]>=5 or row["누계손실(분)"]>=500:
                    return "🔴🔴 P1 즉시"
                elif row["발생횟수"]>=3 or row["누계손실(분)"]>=200:
                    return "🔴 P1"
                elif row["발생횟수"]>=2 or row["누계손실(분)"]>=100:
                    return "🟠 P2"
                else: return "🟡 P3"

            if pm_df.empty:
                st.info("설비 불량 데이터 없음")
            else:
                pm_df["PM등급"] = pm_df.apply(pm_grade, axis=1)
                for _,row in pm_df.iterrows():
                    g   = str(row["PM등급"])
                    cls = ("pm-p1" if "P1" in g
                           else "pm-p2" if "P2" in g else "pm-p3")
                    st.markdown(f"""
                    <div class="{cls}">
                        <b>{g}</b> &nbsp;|&nbsp;
                        {row['process']} {row['line']} —
                        {row['loss_type_name']} &nbsp;|&nbsp;
                        누계 <b>{int(row['누계손실(분)']):,}분</b> /
                        {int(row['발생횟수'])}회
                    </div>""", unsafe_allow_html=True)

    # ── TAB6 ──
    with tab6:
        st.subheader("⬇️ 데이터 다운로드")
        col_a,col_b = st.columns(2)
        with col_a:
            st.markdown("#### 📊 엑셀 리포트")
            excel_data = to_excel(fdf)
            st.download_button(
                "📥 엑셀 다운로드", data=excel_data,
                file_name="HEVH_LOSSTIME_리포트.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
                use_container_width=True)
        with col_b:
            st.markdown("#### 📋 CSV 다운로드")
            csv_data = fdf.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                "📥 CSV 다운로드",
                data=csv_data.encode("utf-8-sig"),
                file_name="HEVH_LOSSTIME_데이터.csv",
                mime="text/csv",
                use_container_width=True)

        st.divider()
        st.markdown("#### 🗄️ 누적 DB 현황")
        db_all = load_db()
        if not db_all.empty:
            st.info(
                f"누적 DB: {len(db_all):,}건 | "
                f"날짜: {db_all['date'].nunique()}일 | "
                f"라인: {db_all['line'].nunique()}개")
            if st.button("🗑️ DB 초기화", type="secondary"):
                if os.path.exists(DB_PATH):
                    os.remove(DB_PATH)
                st.session_state.pop("df", None)
                st.success("DB 초기화 완료")
                st.rerun()
        else:
            st.info("저장된 DB 없음")

# ─────────────────────────────────────────────
# 앱 진입점
# ─────────────────────────────────────────────
def main():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if not st.session_state["logged_in"]:
        login_page()
    else:
        dashboard()

if __name__ == "__main__":
    main()
