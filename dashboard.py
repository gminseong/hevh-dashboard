"""
[한솔테크닉스 HEVH] LOSSTIME + SCRAP 분석 대시보드 v3.4
- 손실유형 세분화 (Magazine부족 분리)
- 타임슬롯별 원인 정확 매핑
- 그래프 클릭 상호작용
- GitHub 자동 저장/로드
실행: python -m streamlit run dashboard.py
"""

import re
import io
import os
import base64
import requests
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
    page_title="HEVH 분석 대시보드",
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
    margin-bottom: 8px; cursor: pointer;
}
.metric-box:hover { background: #e8f4fd; border-color: #2563eb; }
.metric-val { font-size: 32px; font-weight: bold; color: #1e3a5f; }
.metric-lbl { font-size: 13px; color: #64748b; margin-top: 4px; }
.pm-p1 { background:#fef2f2; border-left:4px solid #dc2626;
          padding:10px; border-radius:4px; margin:4px 0; }
.pm-p2 { background:#fff7ed; border-left:4px solid #f97316;
          padding:10px; border-radius:4px; margin:4px 0; }
.pm-p3 { background:#fefce8; border-left:4px solid #eab308;
          padding:10px; border-radius:4px; margin:4px 0; }
.detail-box {
    background:#f0f9ff; border:1px solid #bae6fd;
    border-radius:8px; padding:16px; margin-top:8px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 상수
# ─────────────────────────────────────────────
DAY_SLOTS   = ["A","B","C","D","E","F"]
NIGHT_SLOTS = ["F","G","H","I","J","K"]
DB_PATH     = "losstime_db.csv"
SCRAP_DB    = "scrap_db.csv"

# ★ 세분화된 손실유형 규칙
LOSS_TYPE_RULES = [
    ("NEW_OP",       "신규OP교육",    ["new op","đào tạo","op mới"]),
    ("MAGAZINE",     "Magazine부족",  ["magazine","magazin","mag"]),
    ("WAITING_SMT",  "SMT대기",       ["waiting semi","đợi semi","chờ semi"]),
    ("MATERIAL",     "자재부족",      ["thiếu liệu","chờ liệu","đợi liệu",
                                       "thiếu hàng","chờ hàng","thiếu vật tư"]),
    ("CHANGE_MODEL", "모델교체",      ["change model","đổi model","tách lot",
                                       "cover work","đổi ca","chuyển model"]),
    ("PRINTER",      "Printer불량",   ["printer","priter","lỗi keo","tràn keo"]),
    ("MOUTER",       "Mouter불량",    ["moutor","mounter","stopper","băng tải"]),
    ("INSERT_RH3",   "RH3삽입불량",   ["rh3","rh 3","rhu"]),
    ("INSERT_XGZ",   "XGZ불량",       ["xzg","xgz"]),
    ("INSERT_AXIAL", "Axial불량",     ["axial","radial","cắm rớt","rad lỗi"]),
    ("INSERT_JUMP",  "Jumper불량",    ["jumper","jump "]),
    ("COATING",      "Coating/Reflow불량",["coating","reflow","flux","nhiệt"]),
    ("AOI_SAOI",     "AOI/S-AOI불량", ["saoi","s-aoi","aoi"]),
    ("WAVE",         "Wave Solder불량",["wave solder","wave"]),
    ("SPI",          "SPI불량",       ["spi"]),
    ("INLOADER",     "Inloader불량",  ["inloader"]),
    ("ICT_ATE",      "ICT/ATE불량",   ["ict","ate","ft lỗi"]),
    ("EQUIP_FAIL",   "설비고장(기타)",["hư","lỗi máy","stop line","spare"]),
    ("DONE_PLAN",    "계획완료",      ["done plan","hết plan","kết thúc"]),
    ("SAMPLE",       "샘플/테스트",   ["sample"]),
    ("ETC",          "기타",          []),
]

# 스크랩 원인 분류
SCRAP_CAUSE_RULES = [
    ("BROKEN_DROP",  "낙하/파손",     ["rớt","rơi","bể","vỡ","broken","drop","rộng","sập"]),
    ("BROKEN_BI",    "Burn-in파손",   ["burn in","burin","tủ burin","ốc","tuột"]),
    ("SOLDER",       "납불량",        ["tràn keo","solder over","hàn","keo"]),
    ("CRACK",        "크랙/균열",     ["nứt","crack","bể mạch"]),
    ("PAD_DAMAGE",   "Pad손상",       ["tróc pad","pad đồng","tróc"]),
    ("EYELET",       "Eyelet불량",    ["eyelet","thiếu eyelet"]),
    ("SENSOR",       "센서불량",      ["sensor","cảm biến","chinh lai sensor"]),
    ("MAGAZINE",     "Magazine불량",  ["magazine","megazine","mag"]),
    ("NEW_OP",       "신규OP실수",    ["op mới","thời vụ","mới","new op"]),
    ("AUTO",         "자동판정(ATE)", ["atuo_scrap","auto_scrap"]),
    ("ETC",          "기타",          []),
]

OPER_MAP = {"P-AUTO":"AI","P-SMD":"SMT","P-PBA":"MI"}

PROC_COLOR = {"AI":"#ef4444","SMT":"#3b82f6","MI":"#10b981"}
SCRAP_CAUSE_COLOR = {
    "낙하/파손":"#dc2626","Burn-in파손":"#ea580c","납불량":"#d97706",
    "크랙/균열":"#ca8a04","Pad손상":"#65a30d","Eyelet불량":"#0891b2",
    "센서불량":"#7c3aed","Magazine불량":"#db2777","신규OP실수":"#9f1239",
    "자동판정(ATE)":"#6b7280","기타":"#9ca3af"
}

# ─────────────────────────────────────────────
# GitHub DB 관리
# ─────────────────────────────────────────────
def get_github_config():
    try:
        token = st.secrets["github"]["token"]
        repo  = st.secrets["github"]["repo"]
        return token, repo
    except Exception:
        return None, None

def github_load_csv(filename):
    """GitHub에서 CSV 로드"""
    token, repo = get_github_config()
    if not token:
        # 로컬 파일 폴백
        try:
            return pd.read_csv(filename, encoding="utf-8-sig")
        except FileNotFoundError:
            return pd.DataFrame()

    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}",
               "Accept": "application/vnd.github.v3+json"}
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            content = base64.b64decode(r.json()["content"]).decode("utf-8")
            return pd.read_csv(io.StringIO(content))
        return pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def github_save_csv(df, filename, commit_msg=None):
    """GitHub에 CSV 저장"""
    token, repo = get_github_config()
    if not token:
        df.to_csv(filename, index=False, encoding="utf-8-sig")
        return True

    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}",
               "Accept": "application/vnd.github.v3+json"}

    # 기존 SHA 조회
    sha = None
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha = r.json().get("sha")
    except Exception:
        pass

    # 업로드
    csv_str = df.to_csv(index=False, encoding="utf-8-sig")
    content  = base64.b64encode(csv_str.encode("utf-8-sig")).decode()
    msg = commit_msg or f"DB 업데이트: {filename}"
    payload = {"message": msg, "content": content}
    if sha: payload["sha"] = sha

    try:
        r = requests.put(url, headers=headers,
                         json=payload, timeout=15)
        return r.status_code in [200, 201]
    except Exception:
        return False

# ─────────────────────────────────────────────
# 로그인
# ─────────────────────────────────────────────
def check_login():
    try:
        return st.secrets["users"]
    except Exception:
        return {"hansol":"hevh2024","vietnam":"hevh2024","korea":"hevh2024"}

def login_page():
    st.markdown("""
    <div style="text-align:center; margin-top:60px;">
        <h2>🏭 한솔테크닉스 HEVH</h2>
        <h4 style="color:#64748b;">LOSSTIME + SCRAP 분석 대시보드</h4>
    </div>
    """, unsafe_allow_html=True)
    col1,col2,col3 = st.columns([1,1.2,1])
    with col2:
        with st.form("login_form"):
            st.markdown("#### 🔐 로그인")
            username = st.text_input("아이디", placeholder="ID 입력")
            password = st.text_input("비밀번호", type="password",
                                     placeholder="PW 입력")
            if st.form_submit_button("로그인", use_container_width=True,
                                     type="primary"):
                users = check_login()
                if username in users and users[username]==password:
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

def classify_scrap_cause(comment):
    if not comment: return ("AUTO","자동판정(ATE)")
    t = str(comment).lower()
    for code,name,kws in SCRAP_CAUSE_RULES:
        for kw in kws:
            if kw in t: return (code,name)
    return ("ETC","기타")

def normalize_scrap_line(raw):
    s = str(raw or "").strip().upper()
    s = re.sub(r'^PA\s*0*(\d+)', lambda m: f"SA{int(m.group(1)):02d}", s)
    s = re.sub(r'^PS\s*0*(\d+)', lambda m: f"PS{int(m.group(1)):02d}", s)
    s = re.sub(r'^PM\s*0*(\d+)', lambda m: f"MI {int(m.group(1))}", s)
    return s

def extract_slot_causes(cause_row, slots):
    """
    ★ 핵심 개선: 각 타임슬롯 셀을 개별 파싱
    합쳐진 셀(merged)은 동일 원인으로 해당 슬롯 전체 적용
    """
    result = {s: "" for s in slots}
    if not cause_row: return result

    cells = list(cause_row)[2:2+len(slots)]

    # 실제 셀값 추출 (None = 이전 셀과 병합된 것)
    last_val = ""
    for idx, cell in enumerate(cells):
        if idx >= len(slots): break
        slot = slots[idx]
        if cell is not None and str(cell).strip() not in ["","None","—","-"]:
            val = str(cell).strip()
            # 줄바꿈 제거 후 정리
            val = " ".join(val.split())
            result[slot] = val
            last_val = val
        else:
            # 병합 셀 → 이전 값 유지 (단, 비어있으면 공백)
            result[slot] = ""

    return result

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

def extract_model_per_slot(model_row, slots):
    result = {s:"" for s in slots}
    if not model_row: return result
    cells = list(model_row)[2:2+len(slots)]
    last = ""
    for idx,cell in enumerate(cells):
        if idx >= len(slots): break
        v = str(cell or "").strip()
        mm = re.search(r'(L\d{2}[A-Z0-9_\-\.]+)', v, re.I)
        if mm:
            result[slots[idx]] = mm.group(1); last = mm.group(1)
        elif v and v.lower() not in ["none","—","-",""]:
            result[slots[idx]] = v; last = v
        elif last:
            result[slots[idx]] = last
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

def detect_process(fn):
    fu = fn.upper()
    if "AI REPORT" in fu:  return "AI"
    if "SMD REPORT" in fu: return "SMT"
    if "MI TIME" in fu:    return "MI"
    if "SCRAP" in fu or "SCRAB" in fu or "스크랩" in fu: return "SCRAP"
    return "UNKNOWN"

def parse_date(text):
    if not text: return None
    t = re.sub(r'\b(DAY|NIGHT)\b','',str(text),flags=re.I).strip()
    m = re.search(r'\b(\d{1,2})\.(\d{1,2})\b', t)
    if m:
        a,b = int(m.group(1)),int(m.group(2))
        if 1<=a<=31 and 1<=b<=12: return f"2026-{b:02d}-{a:02d}"
    m = re.search(r'(\d{1,2})/(\d{1,2})', t)
    if m:
        a,b = int(m.group(1)),int(m.group(2))
        if 1<=a<=31 and 1<=b<=12: return f"2026-{b:02d}-{a:02d}"
    return None

def detect_shift(sn, fn):
    return "NIGHT" if "NIGHT" in (sn+fn).upper() else "DAY"

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

            total  = sum(loss_vals)
            models = extract_model_per_slot(model_row, slots)
            action = " | ".join(
                str(v) for v in (list(action_row[2:9]) if action_row else [])
                if v and str(v).strip())

            # ★ 타임슬롯별 원인 개별 파싱
            slot_causes = extract_slot_causes(cause_row, slots)
            cause_all   = " | ".join(
                v for v in slot_causes.values() if v)

            # 복합 여부: 슬롯별 원인이 2종류 이상
            unique_causes = set(v for v in slot_causes.values() if v)
            complexity = "복합" if len(unique_causes)>1 else "단일"

            for idx,slot in enumerate(slots):
                lv = loss_vals[idx] if idx<len(loss_vals) else 0.0
                if lv > 0:
                    # ★ 해당 슬롯 원인 사용
                    cs = slot_causes.get(slot,"")
                    # 빈 경우 인접 슬롯 원인 탐색
                    if not cs:
                        for s2 in slots:
                            if slot_causes.get(s2,""):
                                cs = slot_causes[s2]; break
                    code,name = classify_loss_type(cs)
                    records.append({
                        "date":date_str,"shift":shift,"process":process,
                        "line":line,"time_slot":slot,
                        "model":models.get(slot,""),
                        "loss_min":lv,"loss_type_code":code,
                        "loss_type_name":name,"complexity":complexity,
                        "loss_detail":cs,"action":action
                    })

            if total > 0:
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

def parse_scrap_file(uploaded_file):
    records = []
    try:
        fn = uploaded_file.name.lower()
        if fn.endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"스크랩 파일 읽기 실패: {e}"); return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]
    required = ["Work Date","Start Line","Tran Comment"]
    for r in required:
        if r not in df.columns:
            st.warning(f"컬럼 없음: {r}"); return pd.DataFrame()

    for _, row in df.iterrows():
        try:
            work_date   = str(row.get("Work Date","")).strip()[:10]
            start_line  = normalize_scrap_line(row.get("Start Line",""))
            comment     = str(row.get("Tran Comment","")).strip()
            model_desc  = str(row.get("Model Mat Desc","")).strip()
            oper_desc   = str(row.get("Oper Desc","")).strip()
            result_grp  = str(row.get("Result Group","")).strip()
            result_cd   = str(row.get("Result Code","")).strip()
            result_desc = str(row.get("Result Desc","")).strip()
            reason_cd   = str(row.get("Reason Code","")).strip()
            reason_desc = str(row.get("Reason Desc","")).strip()
            qty         = int(row.get("Qty",1) or 1)
            mm = re.search(r'(L\d{2}[A-Z0-9_\-\.]+)', model_desc, re.I)
            model   = mm.group(1) if mm else model_desc
            process = OPER_MAP.get(oper_desc,"기타")
            cause_code,cause_name = classify_scrap_cause(comment)
            is_auto = "Y" if "ATUO_SCRAP" in comment.upper() or \
                             "AUTO_SCRAP" in comment.upper() else "N"
            records.append({
                "work_date":work_date,"process":process,"line":start_line,
                "model":model,"qty":qty,"cause_code":cause_code,
                "cause_name":cause_name,"result_group":result_grp,
                "result_code":result_cd,"result_desc":result_desc,
                "reason_code":reason_cd,"reason_desc":reason_desc,
                "is_auto":is_auto,"comment":comment,
            })
        except Exception: continue
    return pd.DataFrame(records)

def parse_files(uploaded_files):
    loss_records=[]; scrap_list=[]
    prog=st.progress(0); status=st.empty()
    for fi,uf in enumerate(uploaded_files):
        fn=uf.name; process=detect_process(fn)
        if process=="UNKNOWN":
            status.warning(f"⏭️ 스킵: {fn}")
        elif process=="SCRAP":
            status.info(f"📂 스크랩 파싱: {fn}")
            sdf=parse_scrap_file(uf)
            if not sdf.empty:
                scrap_list.append(sdf)
                status.success(f"✅ 스크랩 {len(sdf):,}건")
        else:
            file_date=parse_date(fn) or "UNKNOWN"
            status.info(f"📂 {fn} → {process}/{file_date}")
            try:
                wb=load_workbook(uf,data_only=True)
            except Exception as e:
                status.error(f"열기 실패: {e}")
                prog.progress((fi+1)/len(uploaded_files)); continue
            for sn in wb.sheetnames:
                ws=wb[sn]
                if isinstance(ws,Chartsheet): continue
                shift=detect_shift(sn,fn)
                date_str=file_date
                if date_str=="UNKNOWN": date_str=parse_date(sn) or "UNKNOWN"
                try:
                    recs=parse_sheet(ws,process,date_str,shift)
                    loss_records.extend(recs)
                except Exception as e:
                    st.warning(f"파싱 오류 [{sn}]: {e}")
        prog.progress((fi+1)/len(uploaded_files))
    status.empty(); prog.empty()
    loss_df=pd.DataFrame(loss_records)
    if not loss_df.empty and "date" in loss_df.columns:
        loss_df=loss_df[loss_df["date"]!="UNKNOWN"].reset_index(drop=True)
    scrap_df=pd.concat(scrap_list,ignore_index=True) \
              if scrap_list else pd.DataFrame()
    return loss_df, scrap_df

# ─────────────────────────────────────────────
# DB 관리
# ─────────────────────────────────────────────
def load_db():
    df = github_load_csv(DB_PATH)
    if not df.empty and "date" in df.columns:
        df = df[df["date"]!="UNKNOWN"].reset_index(drop=True)
    return df

def save_db(df):
    github_save_csv(df, DB_PATH, "LOSSTIME DB 업데이트")

def load_scrap_db():
    return github_load_csv(SCRAP_DB)

def save_scrap_db(df):
    github_save_csv(df, SCRAP_DB, "SCRAP DB 업데이트")

def merge_db(existing, new_df):
    if existing.empty: return new_df
    if new_df.empty:   return existing
    combined=pd.concat([existing,new_df],ignore_index=True)
    key_cols=["date","shift","process","line","time_slot"]
    combined=combined.drop_duplicates(
        subset=[c for c in key_cols if c in combined.columns],keep="last")
    return combined.sort_values(
        ["date","shift","process","line","time_slot"]).reset_index(drop=True)

def merge_scrap_db(existing, new_df):
    if existing.empty: return new_df
    if new_df.empty:   return existing
    combined=pd.concat([existing,new_df],ignore_index=True)
    key_cols=["work_date","line","model","comment"]
    combined=combined.drop_duplicates(
        subset=[c for c in key_cols if c in combined.columns],keep="last")
    return combined.sort_values("work_date").reset_index(drop=True)

# ─────────────────────────────────────────────
# 엑셀 다운로드
# ─────────────────────────────────────────────
def to_excel(df):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as writer:
        total_df=df[df["time_slot"]=="TOTAL"] if "time_slot" in df.columns else df
        line_sum=(total_df.groupby(["process","line"])["loss_min"]
                  .sum().reset_index().sort_values("loss_min",ascending=False)
                  .rename(columns={"loss_min":"누계손실(분)"}))
        line_sum["누계(시간)"]=(line_sum["누계손실(분)"]/60).round(1)
        line_sum.insert(0,"순위",range(1,len(line_sum)+1))
        line_sum.to_excel(writer,sheet_name="라인별누계",index=False)
        type_sum=(total_df.groupby("loss_type_name")["loss_min"]
                  .agg(["sum","count"]).reset_index()
                  .sort_values("sum",ascending=False)
                  .rename(columns={"loss_type_name":"유형",
                                   "sum":"손실(분)","count":"건수"}))
        type_sum.to_excel(writer,sheet_name="손실유형별",index=False)
        df.to_excel(writer,sheet_name="원본데이터",index=False)
    return buf.getvalue()

def to_excel_scrap(df):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as writer:
        cause_sum=(df.groupby("cause_name")["qty"].sum().reset_index()
                   .sort_values("qty",ascending=False)
                   .rename(columns={"cause_name":"원인","qty":"수량"}))
        cause_sum.to_excel(writer,sheet_name="원인별",index=False)
        line_sum=(df.groupby(["process","line"])["qty"].sum().reset_index()
                  .sort_values("qty",ascending=False)
                  .rename(columns={"qty":"수량"}))
        line_sum.to_excel(writer,sheet_name="라인별",index=False)
        df.to_excel(writer,sheet_name="원본데이터",index=False)
    return buf.getvalue()

# ─────────────────────────────────────────────
# 메인 대시보드
# ─────────────────────────────────────────────
def dashboard():
    st.markdown("""
    <div class="main-header">
        <h2>🏭 한솔테크닉스 HEVH — LOSSTIME + SCRAP 분석 대시보드</h2>
        <p style="margin:0;opacity:0.85">AI / SMT / PBA(MI) 공정 | 호치민 법인</p>
    </div>
    """, unsafe_allow_html=True)

    # ── 사이드바 ──
    with st.sidebar:
        st.markdown(f"👤 **{st.session_state.get('username','')}** 님")
        if st.button("🚪 로그아웃", use_container_width=True):
            st.session_state["logged_in"]=False; st.rerun()

        st.divider()
        st.header("📂 파일 업로드")
        st.caption("LOSSTIME + SCRAP 동시 업로드 가능")
        uploaded=st.file_uploader(
            "xlsx / xls (여러 개 가능)",
            type=["xlsx","xls"], accept_multiple_files=True)

        if uploaded:
            if st.button("🚀 분석 시작 / DB 누적",
                         type="primary", use_container_width=True):
                with st.spinner("파싱 중..."):
                    new_loss,new_scrap=parse_files(uploaded)
                if not new_loss.empty:
                    ex=load_db()
                    merged=merge_db(ex,new_loss)
                    with st.spinner("GitHub 저장 중..."):
                        ok=save_db(merged)
                    st.session_state["df"]=merged
                    st.success(f"✅ LOSSTIME {len(new_loss):,}건 → GitHub {'저장완료' if ok else '로컬저장'}")
                if not new_scrap.empty:
                    ex_s=load_scrap_db()
                    merged_s=merge_scrap_db(ex_s,new_scrap)
                    with st.spinner("GitHub 저장 중..."):
                        ok_s=save_scrap_db(merged_s)
                    st.session_state["scrap_df"]=merged_s
                    st.success(f"✅ SCRAP {len(new_scrap):,}건 → GitHub {'저장완료' if ok_s else '로컬저장'}")
                if new_loss.empty and new_scrap.empty:
                    st.error("파싱된 데이터 없음")

        st.divider()
        if st.button("💾 DB 불러오기 (GitHub)", use_container_width=True):
            with st.spinner("GitHub에서 로드 중..."):
                db=load_db(); sdb=load_scrap_db()
            if not db.empty:
                st.session_state["df"]=db
                st.success(f"✅ LOSSTIME {len(db):,}건")
            if not sdb.empty:
                st.session_state["scrap_df"]=sdb
                st.success(f"✅ SCRAP {len(sdb):,}건")
            if db.empty and sdb.empty:
                st.warning("저장된 DB 없음")

        st.divider()
        st.header("🔍 조회 필터")
        df_all=st.session_state.get("df",pd.DataFrame())
        if df_all.empty:
            st.info("데이터를 먼저 불러오세요."); return

        dates=sorted([d for d in df_all["date"].dropna().unique()
                      if d!="UNKNOWN"])
        if not dates:
            st.warning("유효한 날짜 없음"); return

        if len(dates)>=2:
            date_range=st.select_slider("날짜 범위",options=dates,
                                        value=(dates[0],dates[-1]))
        else:
            date_range=(dates[0],dates[0]); st.info(f"날짜: {dates[0]}")

        procs=st.multiselect("공정",["AI","SMT","MI"],
                             default=["AI","SMT","MI"])
        shifts=st.multiselect("주야간",["DAY","NIGHT"],
                              default=["DAY","NIGHT"])
        lines_all=sorted(df_all["line"].dropna().unique())
        sel_lines=st.multiselect("라인 (미선택=전체)",lines_all)
        types_all=sorted(df_all["loss_type_name"].dropna().unique())
        sel_types=st.multiselect("손실유형 (미선택=전체)",types_all)
        view_mode=st.radio("조회 단위",["TOTAL(일계)","타임별(A~K)"],
                           horizontal=True)

    # ── 필터 적용 ──
    df=st.session_state.get("df",pd.DataFrame())
    scrap_df=st.session_state.get("scrap_df",pd.DataFrame())
    if df.empty:
        st.info("👈 파일을 업로드하거나 DB를 불러오세요."); return

    mask=(
        (df["date"]>=date_range[0]) & (df["date"]<=date_range[1]) &
        (df["shift"].isin(shifts or ["DAY","NIGHT"])) &
        (df["process"].isin(procs or ["AI","SMT","MI"]))
    )
    if sel_lines: mask &= df["line"].isin(sel_lines)
    if sel_types: mask &= df["loss_type_name"].isin(sel_types)
    if view_mode=="TOTAL(일계)":
        mask &= (df["time_slot"]=="TOTAL")
    else:
        mask &= (df["time_slot"]!="TOTAL")

    fdf=df[mask].copy()
    total_df=df[mask & (df["time_slot"]=="TOTAL")].copy()

    if fdf.empty:
        st.warning("⚠️ 조건에 맞는 데이터가 없습니다."); return

    # ── KPI + 클릭 상호작용 ──
    total_min=int(total_df["loss_min"].sum()) if not total_df.empty else 0
    total_hr=total_min//60
    n_lines=total_df["line"].nunique() if not total_df.empty else 0
    n_days=fdf["date"].nunique()
    scrap_total=int(scrap_df["qty"].sum()) if not scrap_df.empty else 0

    if "kpi_focus" not in st.session_state:
        st.session_state["kpi_focus"] = None

    c1,c2,c3,c4,c5 = st.columns(5)
    kpi_items = [
        (c1, f"{total_min:,}분", f"총 손실 ({total_hr}시간)", "loss"),
        (c2, f"{n_lines}개",     "분석 라인 수",              "line"),
        (c3, f"{n_days}일",      "분석 일수",                  "date"),
        (c4, f"{scrap_total:,}ea","스크랩 누계",              "scrap"),
        (c5, f"{len(scrap_df):,}건" if not scrap_df.empty else "0건",
             "스크랩 이력", "scrap_hist"),
    ]
    for col,val,lbl,key in kpi_items:
        with col:
            if st.button(f"{val}\n{lbl}", key=f"kpi_{key}",
                         use_container_width=True):
                st.session_state["kpi_focus"] = key

    # KPI 클릭 상세
    focus = st.session_state.get("kpi_focus")
    if focus == "loss" and not total_df.empty:
        with st.expander("📊 손실유형별 상세", expanded=True):
            ts=(total_df.groupby("loss_type_name")["loss_min"]
                .sum().reset_index().sort_values("loss_min",ascending=False))
            st.dataframe(ts.rename(columns={
                "loss_type_name":"유형","loss_min":"손실(분)"}),
                use_container_width=True)
    elif focus == "line" and not total_df.empty:
        with st.expander("📋 라인별 누계 상세", expanded=True):
            ls=(total_df.groupby(["process","line"])["loss_min"]
                .sum().reset_index().sort_values("loss_min",ascending=False))
            st.dataframe(ls.rename(columns={"loss_min":"손실(분)"}),
                         use_container_width=True)
    elif focus == "date" and not fdf.empty:
        with st.expander("📅 날짜별 상세", expanded=True):
            ds=(total_df.groupby(["date","process"])["loss_min"]
                .sum().reset_index().sort_values("date"))
            st.dataframe(ds.rename(columns={"loss_min":"손실(분)"}),
                         use_container_width=True)
    elif focus in ["scrap","scrap_hist"] and not scrap_df.empty:
        with st.expander("📛 스크랩 상세", expanded=True):
            st.dataframe(scrap_df, use_container_width=True, height=300)

    st.divider()

    # ── 탭 ──
    tab1,tab2,tab3,tab4,tab5,tab6,tab7 = st.tabs([
        "📊 손실 분석","📈 트렌드","🕐 타임별",
        "📛 스크랩 분석","🔍 상세 조회","🔧 설비보전 PM","⬇️ 다운로드"
    ])

    # ── TAB1: 손실 분석 + 그래프 클릭 ──
    with tab1:
        if total_df.empty:
            st.warning("TOTAL 데이터 없음")
        else:
            col_l,col_r = st.columns(2)
            with col_l:
                st.subheader("손실유형별 누계 (클릭 → 상세)")
                ts=(total_df.groupby("loss_type_name")["loss_min"]
                    .sum().reset_index()
                    .sort_values("loss_min",ascending=True)
                    .rename(columns={"loss_type_name":"유형","loss_min":"손실(분)"}))
                fig=px.bar(ts, x="손실(분)", y="유형", orientation="h",
                           color="유형", height=480)
                fig.update_layout(showlegend=False,
                                  margin=dict(l=0,r=0,t=20,b=0))
                clicked_type = st.plotly_chart(
                    fig, use_container_width=True,
                    on_select="rerun", key="type_chart")

                # ★ 그래프 클릭 상세
                if clicked_type and clicked_type.get("selection",{}).get("points"):
                    sel_pt = clicked_type["selection"]["points"][0]
                    sel_name = sel_pt.get("label") or sel_pt.get("y","")
                    if sel_name:
                        st.markdown(f"""
                        <div class="detail-box">
                        <b>📋 [{sel_name}] 상세 내역</b>
                        </div>""", unsafe_allow_html=True)
                        detail_df = fdf[
                            fdf["loss_type_name"]==sel_name
                        ][["date","shift","line","time_slot","model",
                           "loss_min","loss_detail"]].sort_values(
                            ["date","line"])
                        st.dataframe(detail_df.reset_index(drop=True),
                                     use_container_width=True, height=300)

            with col_r:
                st.subheader("라인별 누계 TOP 15 (클릭 → 상세)")
                ls=(total_df.groupby(["process","line"])["loss_min"]
                    .sum().reset_index()
                    .sort_values("loss_min",ascending=False).head(15))
                fig2=px.bar(ls, x="line", y="loss_min",
                            color="process", color_discrete_map=PROC_COLOR,
                            height=480,
                            labels={"loss_min":"손실(분)","line":"라인"})
                fig2.update_layout(margin=dict(l=0,r=0,t=20,b=0))
                clicked_line = st.plotly_chart(
                    fig2, use_container_width=True,
                    on_select="rerun", key="line_chart")

                if clicked_line and clicked_line.get("selection",{}).get("points"):
                    sel_ln = clicked_line["selection"]["points"][0].get("x","")
                    if sel_ln:
                        st.markdown(f"""
                        <div class="detail-box">
                        <b>📋 [{sel_ln}] 라인 상세</b>
                        </div>""", unsafe_allow_html=True)
                        line_detail = fdf[
                            fdf["line"]==sel_ln
                        ][["date","shift","time_slot","model",
                           "loss_min","loss_type_name","loss_detail"]
                        ].sort_values(["date","time_slot"])
                        st.dataframe(line_detail.reset_index(drop=True),
                                     use_container_width=True, height=300)

            st.subheader("공정별 비중")
            ps=total_df.groupby("process")["loss_min"].sum().reset_index()
            col_pie,_=st.columns([1,2])
            fig3=px.pie(ps, values="loss_min", names="process",
                        color="process", color_discrete_map=PROC_COLOR,
                        height=320)
            fig3.update_layout(margin=dict(l=0,r=0,t=20,b=0))
            col_pie.plotly_chart(fig3, use_container_width=True)

    # ── TAB2: 트렌드 ──
    with tab2:
        if total_df.empty:
            st.warning("TOTAL 데이터 없음")
        else:
            st.subheader("날짜별 손실 트렌드")
            dt=(total_df.groupby(["date","process"])["loss_min"]
                .sum().reset_index())
            fig4=px.line(dt, x="date", y="loss_min",
                         color="process", color_discrete_map=PROC_COLOR,
                         markers=True, height=380,
                         labels={"loss_min":"손실(분)","date":"날짜"})
            fig4.update_layout(margin=dict(l=0,r=0,t=20,b=0))
            st.plotly_chart(fig4, use_container_width=True)

            st.subheader("DAY vs NIGHT 비교")
            sc=(total_df.groupby(["shift","loss_type_name"])["loss_min"]
                .sum().reset_index())
            fig5=px.bar(sc, x="loss_type_name", y="loss_min", color="shift",
                        color_discrete_map={"DAY":"#f59e0b","NIGHT":"#6366f1"},
                        barmode="group", height=380,
                        labels={"loss_min":"손실(분)","loss_type_name":"손실유형"})
            fig5.update_layout(margin=dict(l=0,r=0,t=20,b=0),
                               xaxis_tickangle=-30)
            st.plotly_chart(fig5, use_container_width=True)

    # ── TAB3: 타임별 ──
    with tab3:
        st.subheader("🕐 타임별 손실 히트맵")
        slot_df=df[
            (df["date"]>=date_range[0]) & (df["date"]<=date_range[1]) &
            (df["shift"].isin(shifts or ["DAY","NIGHT"])) &
            (df["process"].isin(procs or ["AI","SMT","MI"])) &
            (df["time_slot"]!="TOTAL")
        ].copy()
        if sel_lines: slot_df=slot_df[slot_df["line"].isin(sel_lines)]

        if slot_df.empty:
            st.warning("타임별 데이터 없음")
        else:
            slot_order=["A","B","C","D","E","F","G","H","I","J","K"]
            pivot=(slot_df.groupby(["line","time_slot"])["loss_min"]
                   .sum().reset_index())
            pivot_table=pivot.pivot(
                index="line",columns="time_slot",values="loss_min").fillna(0)
            cols_sorted=[c for c in slot_order if c in pivot_table.columns]
            pivot_table=pivot_table[cols_sorted]
            fig_heat=px.imshow(
                pivot_table, color_continuous_scale="Reds", aspect="auto",
                height=max(400,len(pivot_table)*30),
                labels={"x":"시간대","y":"라인","color":"손실(분)"},
                title="라인 × 시간대 손실 히트맵")
            fig_heat.update_layout(margin=dict(l=0,r=0,t=40,b=0))
            st.plotly_chart(fig_heat, use_container_width=True)

            st.subheader("시간대별 손실 합계")
            slot_sum=(slot_df.groupby(["time_slot","process"])["loss_min"]
                      .sum().reset_index())
            slot_sum["time_slot"]=pd.Categorical(
                slot_sum["time_slot"],categories=slot_order,ordered=True)
            slot_sum=slot_sum.sort_values("time_slot")
            fig_slot=px.bar(slot_sum, x="time_slot", y="loss_min",
                            color="process", color_discrete_map=PROC_COLOR,
                            barmode="stack", height=350,
                            labels={"loss_min":"손실(분)","time_slot":"시간대"})
            fig_slot.update_layout(margin=dict(l=0,r=0,t=20,b=0))
            st.plotly_chart(fig_slot, use_container_width=True)

            st.subheader("타임별 상세")
            slot_detail=(slot_df.groupby(
                ["date","shift","line","time_slot","loss_type_name",
                 "loss_detail"])["loss_min"]
                .sum().reset_index()
                .sort_values(["date","line","time_slot"]))
            st.dataframe(slot_detail, use_container_width=True, height=400)

    # ── TAB4: 스크랩 ──
    with tab4:
        st.subheader("📛 스크랩 분석")
        if scrap_df.empty:
            st.info("스크랩 데이터 없음. 파일명에 'SCRAP'/'스크랩' 포함 후 업로드.")
        else:
            scrap_dates=sorted(scrap_df["work_date"].dropna().unique())
            if len(scrap_dates)>=2:
                s_range=st.select_slider("스크랩 날짜 범위",
                                         options=scrap_dates,
                                         value=(scrap_dates[0],scrap_dates[-1]),
                                         key="scrap_date")
                sdf=scrap_df[(scrap_df["work_date"]>=s_range[0]) &
                             (scrap_df["work_date"]<=s_range[1])].copy()
            else:
                sdf=scrap_df.copy()

            show_auto=st.checkbox("자동판정(ATE) 포함", value=True)
            if not show_auto: sdf=sdf[sdf["is_auto"]=="N"]

            if sdf.empty:
                st.warning("조건에 맞는 스크랩 데이터 없음")
            else:
                k1,k2,k3,k4=st.columns(4)
                for col,val,lbl in [
                    (k1,f"{int(sdf['qty'].sum()):,}ea","총 스크랩"),
                    (k2,f"{len(sdf):,}건","스크랩 이력"),
                    (k3,f"{sdf['model'].nunique()}종","모델 수"),
                    (k4,f"{sdf['line'].nunique()}개","발생 라인"),
                ]:
                    col.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-val">{val}</div>
                        <div class="metric-lbl">{lbl}</div>
                    </div>""", unsafe_allow_html=True)
                st.divider()

                col_a,col_b=st.columns(2)
                with col_a:
                    st.subheader("원인별 파레토 (클릭 → 상세)")
                    cause_g=(sdf.groupby("cause_name")["qty"]
                             .sum().reset_index()
                             .sort_values("qty",ascending=False))
                    cause_g["누계%"]=(cause_g["qty"].cumsum() /
                                      cause_g["qty"].sum()*100).round(1)
                    fig_p=go.Figure()
                    fig_p.add_bar(x=cause_g["cause_name"],
                                  y=cause_g["qty"], name="수량",
                                  marker_color="#ef4444")
                    fig_p.add_scatter(x=cause_g["cause_name"],
                                      y=cause_g["누계%"],
                                      mode="lines+markers", name="누계%",
                                      yaxis="y2",
                                      line=dict(color="#1e3a5f",width=2))
                    fig_p.update_layout(
                        yaxis2=dict(overlaying="y",side="right",
                                    range=[0,110],ticksuffix="%"),
                        height=400, margin=dict(l=0,r=0,t=20,b=0),
                        xaxis_tickangle=-30)
                    clicked_cause=st.plotly_chart(
                        fig_p, use_container_width=True,
                        on_select="rerun", key="scrap_cause_chart")
                    if clicked_cause and \
                       clicked_cause.get("selection",{}).get("points"):
                        sel_cause=clicked_cause["selection"]["points"][0].get("x","")
                        if sel_cause:
                            st.markdown(f"**📋 [{sel_cause}] 상세**")
                            cd=sdf[sdf["cause_name"]==sel_cause][
                                ["work_date","process","line","model",
                                 "qty","comment"]].sort_values("work_date")
                            st.dataframe(cd.reset_index(drop=True),
                                         use_container_width=True, height=250)

                with col_b:
                    st.subheader("공정별 비중")
                    proc_g=sdf.groupby("process")["qty"].sum().reset_index()
                    fig_pc=px.pie(proc_g, values="qty", names="process",
                                  color="process",
                                  color_discrete_map=PROC_COLOR, height=400)
                    fig_pc.update_layout(margin=dict(l=0,r=0,t=20,b=0))
                    st.plotly_chart(fig_pc, use_container_width=True)

                st.subheader("라인별 TOP 15 (클릭 → 상세)")
                line_g=(sdf.groupby(["process","line"])["qty"]
                        .sum().reset_index()
                        .sort_values("qty",ascending=False).head(15))
                fig_l=px.bar(line_g, x="line", y="qty",
                             color="process", color_discrete_map=PROC_COLOR,
                             height=350,
                             labels={"qty":"수량","line":"라인"})
                fig_l.update_layout(margin=dict(l=0,r=0,t=20,b=0))
                clicked_sl=st.plotly_chart(
                    fig_l, use_container_width=True,
                    on_select="rerun", key="scrap_line_chart")
                if clicked_sl and \
                   clicked_sl.get("selection",{}).get("points"):
                    sel_sl=clicked_sl["selection"]["points"][0].get("x","")
                    if sel_sl:
                        st.markdown(f"**📋 [{sel_sl}] 라인 스크랩 상세**")
                        ld=sdf[sdf["line"]==sel_sl][
                            ["work_date","model","qty","cause_name","comment"]
                        ].sort_values("work_date")
                        st.dataframe(ld.reset_index(drop=True),
                                     use_container_width=True, height=250)

                st.subheader("날짜별 트렌드")
                date_g=(sdf.groupby(["work_date","process"])["qty"]
                        .sum().reset_index())
                fig_dt=px.line(date_g, x="work_date", y="qty",
                               color="process", color_discrete_map=PROC_COLOR,
                               markers=True, height=320,
                               labels={"qty":"수량","work_date":"날짜"})
                fig_dt.update_layout(margin=dict(l=0,r=0,t=20,b=0))
                st.plotly_chart(fig_dt, use_container_width=True)

                st.subheader("🔧 개선 우선순위")
                improve=(sdf[sdf["is_auto"]=="N"]
                         .groupby(["process","line","cause_name"])["qty"]
                         .agg(["sum","count"]).reset_index()
                         .sort_values("sum",ascending=False)
                         .rename(columns={"sum":"수량","count":"발생횟수"}))

                def scrap_grade(row):
                    if row["수량"]>=10 or row["발생횟수"]>=5:
                        return "🔴🔴 즉시개선"
                    elif row["수량"]>=5 or row["발생횟수"]>=3:
                        return "🔴 긴급"
                    elif row["수량"]>=3 or row["발생횟수"]>=2:
                        return "🟠 주의"
                    else: return "🟡 모니터링"

                improve["우선순위"]=improve.apply(scrap_grade,axis=1)
                for _,row in improve.head(10).iterrows():
                    g=str(row["우선순위"])
                    cls=("pm-p1" if "즉시" in g or "긴급" in g
                         else "pm-p2" if "주의" in g else "pm-p3")
                    st.markdown(f"""
                    <div class="{cls}">
                        <b>{g}</b> &nbsp;|&nbsp;
                        {row['process']} {row['line']} —
                        {row['cause_name']} &nbsp;|&nbsp;
                        <b>{int(row['수량']):,}ea</b> / {int(row['발생횟수'])}회
                    </div>""", unsafe_allow_html=True)

                st.divider()
                st.subheader("스크랩 상세 내역")
                search_s=st.text_input("🔍 키워드 (모델/라인/원인/코멘트)")
                show_s=sdf.copy()
                if search_s:
                    ms=(show_s["model"].astype(str).str.contains(
                            search_s,case=False,na=False) |
                        show_s["line"].astype(str).str.contains(
                            search_s,case=False,na=False) |
                        show_s["cause_name"].astype(str).str.contains(
                            search_s,case=False,na=False) |
                        show_s["comment"].astype(str).str.contains(
                            search_s,case=False,na=False))
                    show_s=show_s[ms]
                disp_s=["work_date","process","line","model","qty",
                         "cause_name","is_auto","result_desc","comment"]
                disp_s=[c for c in disp_s if c in show_s.columns]
                st.caption(f"총 {len(show_s):,}건")
                st.dataframe(show_s[disp_s].reset_index(drop=True),
                             use_container_width=True, height=400)

    # ── TAB5: 상세 조회 ──
    with tab5:
        st.subheader("상세 데이터 조회")
        search=st.text_input("🔍 키워드 (라인/원인/모델)")
        show_df=fdf.copy()
        if search:
            m=(show_df["line"].astype(str).str.contains(
                   search,case=False,na=False) |
               show_df["loss_detail"].astype(str).str.contains(
                   search,case=False,na=False) |
               show_df["model"].astype(str).str.contains(
                   search,case=False,na=False))
            show_df=show_df[m]
        st.caption(f"총 {len(show_df):,}건")
        disp=["date","shift","process","line","time_slot",
              "model","loss_min","loss_type_name","complexity","loss_detail"]
        disp=[c for c in disp if c in show_df.columns]
        if not show_df.empty:
            st.dataframe(show_df[disp].reset_index(drop=True),
                         use_container_width=True, height=500)
        else:
            st.info("검색 결과 없음")

    # ── TAB6: 설비보전 PM ──
    with tab6:
        st.subheader("🔧 설비보전 PM 우선순위")
        if total_df.empty:
            st.warning("데이터 없음")
        else:
            pm_types=["Printer불량","Axial불량","RH3삽입불량","Mouter불량",
                      "XGZ불량","설비고장(기타)","Coating/Reflow불량",
                      "AOI/S-AOI불량","Wave Solder불량"]
            pm_df=(total_df[total_df["loss_type_name"].isin(pm_types)]
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
                pm_df["PM등급"]=pm_df.apply(pm_grade,axis=1)
                for _,row in pm_df.iterrows():
                    g=str(row["PM등급"])
                    cls=("pm-p1" if "P1" in g
                         else "pm-p2" if "P2" in g else "pm-p3")
                    st.markdown(f"""
                    <div class="{cls}">
                        <b>{g}</b> &nbsp;|&nbsp;
                        {row['process']} {row['line']} —
                        {row['loss_type_name']} &nbsp;|&nbsp;
                        누계 <b>{int(row['누계손실(분)']):,}분</b> /
                        {int(row['발생횟수'])}회
                    </div>""", unsafe_allow_html=True)

    # ── TAB7: 다운로드 ──
    with tab7:
        st.subheader("⬇️ 데이터 다운로드")
        col_a,col_b,col_c=st.columns(3)
        with col_a:
            st.markdown("#### 📊 LOSSTIME 엑셀")
            st.download_button(
                "📥 LOSSTIME 다운로드", data=to_excel(fdf),
                file_name="HEVH_LOSSTIME_리포트.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
                use_container_width=True)
        with col_b:
            st.markdown("#### 📛 SCRAP 엑셀")
            if not scrap_df.empty:
                st.download_button(
                    "📥 SCRAP 다운로드", data=to_excel_scrap(scrap_df),
                    file_name="HEVH_SCRAP_리포트.xlsx",
                    mime="application/vnd.openxmlformats-officedocument"
                         ".spreadsheetml.sheet",
                    use_container_width=True)
            else:
                st.info("스크랩 데이터 없음")
        with col_c:
            st.markdown("#### 📋 CSV")
            csv_data=fdf.to_csv(index=False,encoding="utf-8-sig")
            st.download_button(
                "📥 CSV 다운로드",
                data=csv_data.encode("utf-8-sig"),
                file_name="HEVH_LOSSTIME_데이터.csv",
                mime="text/csv", use_container_width=True)

        st.divider()
        st.markdown("#### 🗄️ DB 현황")
        col_d1,col_d2=st.columns(2)
        with col_d1:
            db_all=load_db()
            if not db_all.empty:
                st.info(f"LOSSTIME: {len(db_all):,}건 | "
                        f"{db_all['date'].nunique()}일 | "
                        f"{db_all['line'].nunique()}개 라인")
                if st.button("🗑️ LOSSTIME DB 초기화"):
                    github_save_csv(pd.DataFrame(), DB_PATH, "DB 초기화")
                    st.session_state.pop("df",None)
                    st.success("초기화 완료"); st.rerun()
        with col_d2:
            sdb_all=load_scrap_db()
            if not sdb_all.empty:
                st.info(f"SCRAP: {len(sdb_all):,}건 | "
                        f"{int(sdb_all['qty'].sum())}ea | "
                        f"{sdb_all['line'].nunique()}개 라인")
                if st.button("🗑️ SCRAP DB 초기화"):
                    github_save_csv(pd.DataFrame(), SCRAP_DB, "SCRAP DB 초기화")
                    st.session_state.pop("scrap_df",None)
                    st.success("초기화 완료"); st.rerun()

# ─────────────────────────────────────────────
# 앱 진입점
# ─────────────────────────────────────────────
def main():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"]=False
    if not st.session_state["logged_in"]:
        login_page()
    else:
        dashboard()

if __name__=="__main__":
    main()
