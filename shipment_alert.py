"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v21.0
- Shipment Rev 시트 사용
- 일자별 계획(Plan) vs MES 실적 정밀 매칭
- 과거(≤기준일)=실적, 미래(>기준일)=계획
- 캐싱 + HTML 빠른 렌더링
"""
from datetime import datetime
import io
import re
import pandas as pd
import streamlit as st


# ════════════════════════════════════════════════════════════
# 캐싱
# ════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def read_excel_raw(file_bytes, sheet_name):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def read_csv_cached(file_bytes):
    return pd.read_csv(io.BytesIO(file_bytes))


def classify(x):
    x = str(x).strip()
    if x.startswith('018'): return '3IN1'
    if x.startswith('01'): return 'PD'
    return 'OTHER'


def parse_date_from_col(col_name, year=2026):
    """다양한 형태의 컬럼명에서 날짜 추출"""
    try:
        s = str(col_name).lower()
        # plan, actual, cut off 등 prefix 모두 제거
        s = re.sub(r'(plan|actual|cut\s*off|cargo)[\.\s_&]*', '', s).strip()
        # 숫자만 추출
        nums = re.findall(r'\d+', s)
        if len(nums) >= 2:
            month = int(nums[0])
            day = int(nums[1])
            if 1 <= month <= 12 and 1 <= day <= 31:
                return pd.Timestamp(year, month, day)
        elif len(nums) == 1:
            day = int(nums[0])
            if 1 <= day <= 31:
                return pd.Timestamp(year, 6, day)  # W25는 6월
        return None
    except Exception:
        return None


# ════════════════════════════════════════════════════════════
# Shipment Rev 로드
# ════════════════════════════════════════════════════════════
def load_shipment_rev(file_bytes):
    sheet_names = list_sheets(file_bytes)

    # Shipment Rev 우선
    target = None
    for s in sheet_names:
        cl = str(s).strip().lower()
        if 'shipment' in cl and 'rev' in cl:
            target = s
            break
    if target is None:
        for s in sheet_names:
            if str(s).strip().lower() == 'shipment':
                target = s
                break
    if target is None:
        target = sheet_names[0]

    st.success(f"✅ 사용 시트: '{target}'")
    raw = read_excel_raw(file_bytes, target)

    # 헤더 행 탐지
    header_row = None
    for i in range(min(8, len(raw))):
        row_vals = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
        if 'cus' in row_vals:
            header_row = i
            break
    if header_row is None:
        header_row = 2

    # 2줄 헤더 병합
    main = raw.iloc[header_row].values
    sub = raw.iloc[header_row + 1].values if header_row + 1 < len(raw) else [None] * len(main)
    
    merged = []
    for mh, sh in zip(main, sub):
        m_str = str(mh).strip() if pd.notna(mh) else ''
        s_str = str(sh).strip() if pd.notna(sh) else ''
        if m_str in ('nan', 'None'): m_str = ''
        if s_str in ('nan', 'None'): s_str = ''
        
        if m_str and s_str:
            merged.append(f"{m_str}.{s_str}")
        elif m_str:
            merged.append(m_str)
        elif s_str:
            merged.append(s_str)
        else:
            merged.append(f"col_{len(merged)}")

    df = raw.iloc[header_row + 2:].copy()
    df.columns = merged
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)

    # 컬럼명 정규화
    rename_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        cl_clean = cl.replace(' ', '').replace('(', '').replace(')', '')
        
        if cl == 'cus' or 'customer' in cl:
            rename_map[c] = 'Cus'
        elif 'cut off cargo' in cl:
            rename_map[c] = 'Cut off Cargo'
        elif cl == 'hq' or 'hq request' in cl or 'hq.request' in cl:
            rename_map[c] = 'HQ Request'
        elif cl == 'model':
            rename_map[c] = 'model'
        elif cl == 'inch':
            rename_map[c] = 'Inch'
        elif cl == 'bncode' or ('bn' in cl and 'code' in cl):
            rename_map[c] = 'code'
        elif '3in1code' in cl_clean or cl == 'erp':
            rename_map[c] = 'ERP'
        elif 'po remain' in cl or 'po.remain' in cl:
            rename_map[c] = 'PO Remain'
        elif 'ttl ship' in cl or 'ttl.ship' in cl:
            rename_map[c] = 'TTL Ship'
        elif 'ttl plan' in cl or 'ttl.plan' in cl:
            rename_map[c] = 'TTL Plan'
        elif 'ttlstock' in cl_clean:
            rename_map[c] = 'TTLstock'
        elif cl == 'balance':
            rename_map[c] = 'BALANCE'
        elif cl == 'note':
            rename_map[c] = 'Note'
        elif 'o/stock' in cl or 'o.stock' in cl:
            rename_map[c] = 'O/stock'
        elif cl == 'gap':
            rename_map[c] = 'GAP_excel'
    df = df.rename(columns=rename_map)

    # Plan 일자별 컬럼 자동 탐지 (TTL Plan 제외)
    plan_date_cols = []
    for c in df.columns:
        cl = str(c).lower().strip()
        # 'Plan'으로 시작하고 숫자가 포함된 컬럼 (단, TTL Plan은 제외)
        if 'plan' in cl and any(ch.isdigit() for ch in cl):
            if 'ttl' in cl:
                continue
            # plan&actual은 미래 계획+실적 혼합이므로 제외 (선택적)
            if 'actual' in cl:
                continue
            plan_date_cols.append(c)

    # ERP 패턴 매칭
    if 'ERP' not in df.columns:
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(30).tolist()
            erp_count = sum(1 for v in sample 
                            if (v.startswith('013') or v.startswith('018')) 
                            and len(v) >= 10)
            if erp_count >= 3:
                df = df.rename(columns={col: 'ERP'})
                break

    if 'ERP' not in df.columns:
        st.error("❌ ERP 컬럼 매핑 실패")
        return pd.DataFrame(), []

    # 유효 행
    df['ERP'] = df['ERP'].astype(str).str.strip()
    df = df[df['ERP'].str.startswith(('013', '018'))].copy()
    df = df.reset_index(drop=True)

    # 디버그 정보
    with st.expander("🔍 컬럼 매핑 진단", expanded=False):
        st.caption(f"전체 컬럼 ({len(df.columns)}개): {list(df.columns)}")
        st.caption(f"일자별 Plan 컬럼 ({len(plan_date_cols)}개): {plan_date_cols}")
        parsed = {c: str(parse_date_from_col(c)) for c in plan_date_cols}
        st.caption(f"날짜 파싱: {parsed}")

    return df, plan_date_cols


# ════════════════════════════════════════════════════════════
# 분석 (정밀 매칭)
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db, plan_date_cols):
    prod_db = prod_db.copy()
    prod_db['TRAN_WORK_DATE'] = pd.to_datetime(prod_db['TRAN_WORK_DATE'], errors='coerce')
    today = prod_db['TRAN_WORK_DATE'].max()
    
    if pd.isna(today):
        st.error("❌ 생산실적 날짜 파싱 실패")
        return pd.DataFrame()

    st.info(f"📅 기준일: **{today.strftime('%Y-%m-%d')}** (MES 최신 일자) | 이 일자까지는 실적, 이후는 계획 사용")

    # MES 실적 일자별 집계
    prod_db['ERP'] = prod_db['FINAL_MAT_ID'].astype(str).str.strip()
    prod_db['TYPE'] = prod_db['ERP'].apply(classify)
    
    valid = prod_db[
        ((prod_db['TYPE']=='PD') & (prod_db['OPER_DESC']=='P-ATE')) |
        ((prod_db['TYPE']=='3IN1') & (prod_db['OPER_DESC']=='ASSY'))
    ].copy()

    # ERP × 날짜
    daily = valid.groupby(['ERP', valid['TRAN_WORK_DATE'].dt.normalize()])['QTY'].sum().reset_index()
    daily.columns = ['ERP', 'DATE', 'QTY']
    daily_dict = {(r['ERP'], r['DATE']): r['QTY'] for _, r in daily.iterrows()}
    
    # ERP별 총 누적
    total_actual = valid.groupby('ERP')['QTY'].sum().to_dict()

    # Shipment Rev
    m = ship_db.copy()
    m['MODEL_TYPE'] = m['ERP'].apply(classify)
    m['MES_누적실적'] = m['ERP'].map(total_actual).fillna(0).astype(int)

    # 핵심 컬럼 숫자화
    for col in ['PO Remain', 'TTL Ship', 'TTL Plan', 'TTLstock', 'BALANCE', 'O/stock'] + plan_date_cols:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors='coerce').fillna(0)

    if 'O/stock' not in m.columns:
        m['O/stock'] = 0
    m['O/stock'] = m['O/stock'].astype(int)

    # 일자별 매칭으로 조정 TTL 계산
    plan_date_map = {col: parse_date_from_col(col) for col in plan_date_cols}
    valid_plan_cols = {col: dt for col, dt in plan_date_map.items() if dt is not None}

    today_norm = today.normalize()

    def calc_adjusted_ttl(row):
        erp = row['ERP']
        past_actual = 0  # 과거 일자의 MES 실적 합
        future_plan = 0  # 미래 일자의 계획 합
        
        for col, dt in valid_plan_cols.items():
            dt_norm = dt.normalize()
            if dt_norm <= today_norm:
                # 과거 또는 오늘: MES 실적 사용
                past_actual += daily_dict.get((erp, dt_norm), 0)
            else:
                # 미래: 계획 사용
                future_plan += row.get(col, 0)
        
        return int(past_actual + future_plan), int(past_actual), int(future_plan)

    if valid_plan_cols:
        results = m.apply(calc_adjusted_ttl, axis=1, result_type='expand')
        m['조정_TTL'] = results[0]
        m['과거_실적합'] = results[1]
        m['미래_계획합'] = results[2]
    else:
        m['조정_TTL'] = m['MES_누적실적']
        m['과거_실적합'] = m['MES_누적실적']
        m['미래_계획합'] = 0
        st.warning("⚠️ 일자별 Plan 컬럼이 없어 누적실적만 사용")

    # 조정 BALANCE
    m['조정_BALANCE'] = m['O/stock'] + m['조정_TTL'] - m['TTL Ship'].astype(int)
    
    # GAP = 조정 - 계획
    if 'BALANCE' in m.columns:
        m['GAP'] = m['조정_BALANCE'] - m['BALANCE'].astype(int)
    else:
        m['GAP'] = 0

    # 알람
    def get_alert(row):
        adj_bal = row['조정_BALANCE']
        plan_bal = row.get('BALANCE', 0)
        if adj_bal < -5000:
            return "🔴 긴급"
        elif adj_bal < 0:
            return "🟠 부족"
        elif plan_bal >= 0 and adj_bal < plan_bal - 1000:
            return "🟡 악화"
        else:
            return "✅ 정상"
    m['알람'] = m.apply(get_alert, axis=1)

    # 정수 변환
    for col in ['PO Remain', 'TTL Ship', 'TTL Plan', 'TTLstock', 'BALANCE']:
        if col in m.columns:
            m[col] = m[col].astype(int)

    return m


# ════════════════════════════════════════════════════════════
# HTML 테이블
# ════════════════════════════════════════════════════════════
def render_html_table(df, height=500):
    numeric_cols = set()
    for c in df.columns:
        if df[c].dtype.kind in 'iuf':
            numeric_cols.add(c)

    def fmt(val, is_num):
        if pd.isna(val) or val == '':
            return '-'
        if is_num:
            try:
                n = int(float(val))
                if n < 0:
                    return f'<span style="color:#DC2626;font-weight:700;">{n:,}</span>'
                if n == 0:
                    return '<span style="color:#9CA3AF;">0</span>'
                return f'{n:,}'
            except Exception:
                pass
        return str(val)

    def bg(alert):
        a = str(alert)
        if '긴급' in a: return '#FEE2E2'
        if '부족' in a: return '#FFEDD5'
        if '악화' in a: return '#FEF3C7'
        return '#FFFFFF'

    parts = [
        f'<div style="max-height:{height}px;overflow:auto;border:1px solid #E5E7EB;border-radius:6px;">',
        '<table style="width:100%;border-collapse:collapse;font-family:-apple-system,sans-serif;font-size:12px;">',
        '<thead style="position:sticky;top:0;z-index:10;"><tr style="background-color:#374151;color:white;">'
    ]
    
    for col in df.columns:
        parts.append(f'<th style="padding:8px 6px;text-align:center;border:1px solid #4B5563;font-weight:700;white-space:nowrap;">{col}</th>')
    parts.append('</tr></thead><tbody>')

    for _, row in df.iterrows():
        row_bg = bg(row.get('알람', ''))
        parts.append(f'<tr style="background-color:{row_bg};">')
        for col in df.columns:
            is_num = col in numeric_cols
            cell = fmt(row[col], is_num)
            align = 'right' if is_num else 'center'
            parts.append(f'<td style="padding:6px;text-align:{align};border-bottom:1px solid #E5E7EB;white-space:nowrap;">{cell}</td>')
        parts.append('</tr>')
    
    parts.append('</tbody></table></div>')
    st.markdown(''.join(parts), unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 일자별 Plan vs MES 실적 정밀 매칭 | 과거=실적, 미래=계획")

    ship_db = st.session_state.get('ship_db', pd.DataFrame())
    plan_cols = st.session_state.get('plan_date_cols', [])
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_t = st.session_state.get('ship_updated', '-')
    prod_t = st.session_state.get('prod_updated', '-')

    s1, s2 = st.columns(2)
    if not ship_db.empty:
        s1.success(f"📁 출하계획: **{len(ship_db)}건** | {ship_t}")
    else:
        s1.warning("📁 출하계획: 없음")
    if not prod_db.empty:
