"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v20.0 (진짜 최종)
- Shipment Rev 시트 사용
- 일자별 계획(6-13~6-20) vs MES 실적 정밀 매칭
- 과거: 실적 / 미래: 계획 → 조정 BALANCE
- 캐싱 + 빠른 HTML 렌더링
"""
from datetime import datetime
import io
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


# ════════════════════════════════════════════════════════════
# Shipment Rev 로드 (2줄 헤더 병합)
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

    # 헤더 행 자동 탐지 (Cus 키워드)
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
    plan_date_cols = []  # 일자별 계획 컬럼들
    
    for c in df.columns:
        cl = str(c).lower().strip()
        cl_clean = cl.replace(' ', '').replace('(', '').replace(')', '')
        
        if cl == 'cus' or 'customer' in cl:
            rename_map[c] = 'Cus'
        elif 'cut off cargo' in cl:
            rename_map[c] = 'Cut off Cargo'
        elif cl == 'hq' or 'hq request' in cl:
            rename_map[c] = 'HQ Request'
        elif cl == 'model':
            rename_map[c] = 'model'
        elif cl == 'inch':
            rename_map[c] = 'Inch'
        elif cl == 'bncode' or ('bn' in cl and 'code' in cl):
            rename_map[c] = 'code'
        elif '3in1code' in cl_clean or cl == 'erp':
            rename_map[c] = 'ERP'
        elif 'po remain' in cl:
            rename_map[c] = 'PO Remain'
        elif 'ttl ship' in cl:
            rename_map[c] = 'TTL Ship'
        elif 'ttl plan' in cl:
            rename_map[c] = 'TTL Plan'
        elif 'ttlstock' in cl_clean:
            rename_map[c] = 'TTLstock'
        elif cl == 'balance':
            rename_map[c] = 'BALANCE'
        elif cl == 'note':
            rename_map[c] = 'Note'
        elif 'o/stock' in cl or cl == 'gap':
            rename_map[c] = c  # 그대로
        # 일자별 계획 컬럼 (Plan.6-13, Plan.6/13, plan.6-15 등)
        elif 'plan' in cl and any(d in cl for d in ['6-', '6/', '6.', '-13', '-14', '-15', '-16', '-17', '-18', '-19', '-20']):
            plan_date_cols.append(c)
    
    df = df.rename(columns=rename_map)
    
    # plan_date_cols 다시 찾기 (rename 후)
    plan_date_cols = []
    for c in df.columns:
        cl = str(c).lower()
        # "Plan.6-13" 같은 패턴
        if cl.startswith('plan.') or cl.startswith('plan '):
            # 날짜 부분 추출
            date_part = cl.replace('plan.', '').replace('plan ', '').strip()
            # 6-13, 6/13, 6.13 등 형태 검사
            if any(d in date_part for d in ['-', '/', '.']):
                plan_date_cols.append(c)

    # ERP 패턴 매칭 (없으면)
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

    st.caption(f"📌 일자별 계획 컬럼: {plan_date_cols}")
    return df, plan_date_cols


# ════════════════════════════════════════════════════════════
# 일자 파싱 헬퍼
# ════════════════════════════════════════════════════════════
def parse_date_from_col(col_name, year=2026):
    """'Plan.6-13', 'Plan.6/13' → datetime(2026, 6, 13)"""
    try:
        s = str(col_name).lower().replace('plan.', '').replace('plan ', '').strip()
        # 6-13, 6/13, 6.13
        for sep in ['-', '/', '.']:
            if sep in s:
                parts = s.split(sep)
                if len(parts) >= 2:
                    month = int(parts[0])
                    day = int(parts[1])
                    return pd.Timestamp(year, month, day)
        return None
    except Exception:
        return None


# ════════════════════════════════════════════════════════════
# 분석 (정밀 매칭)
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db, plan_date_cols):
    # 1. 오늘 날짜 = MES 데이터의 max
    prod_db = prod_db.copy()
    prod_db['TRAN_WORK_DATE'] = pd.to_datetime(prod_db['TRAN_WORK_DATE'], errors='coerce')
    today = prod_db['TRAN_WORK_DATE'].max()
    
    if pd.isna(today):
        st.error("❌ 생산실적 날짜 파싱 실패")
        return pd.DataFrame()

    st.info(f"📅 기준일: {today.strftime('%Y-%m-%d')} (MES 최신 일자)")

    # 2. MES 실적 일자별 집계
    prod_db['ERP'] = prod_db['FINAL_MAT_ID'].astype(str).str.strip()
    prod_db['TYPE'] = prod_db['ERP'].apply(classify)
    
    pd_data = prod_db[(prod_db['TYPE']=='PD') & (prod_db['OPER_DESC']=='P-ATE')]
    in1_data = prod_db[(prod_db['TYPE']=='3IN1') & (prod_db['OPER_DESC']=='ASSY')]
    valid = pd.concat([pd_data, in1_data])

    # ERP × 날짜 일자별 실적
    daily_actual = valid.groupby(['ERP', 'TRAN_WORK_DATE'])['QTY'].sum().reset_index()
    daily_actual_dict = {(r['ERP'], r['TRAN_WORK_DATE'].normalize()): r['QTY'] 
                         for _, r in daily_actual.iterrows()}
    
    # ERP별 총 누적 실적
    total_actual = valid.groupby('ERP')['QTY'].sum().to_dict()

    # 3. Shipment Rev 처리
    m = ship_db.copy()
    m['MODEL_TYPE'] = m['ERP'].apply(classify)
    m['MES_누적실적'] = m['ERP'].map(total_actual).fillna(0).astype(int)

    # 핵심 컬럼 숫자화
    for col in ['PO Remain', 'TTL Ship', 'TTL Plan', 'TTLstock', 'BALANCE'] + plan_date_cols:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors='coerce').fillna(0)
    
    # O/stock 컬럼 찾기
    o_stock_col = None
    for c in m.columns:
        if 'o/stock' in str(c).lower():
            o_stock_col = c
            break
    if o_stock_col:
        m['O/stock'] = pd.to_numeric(m[o_stock_col], errors='coerce').fillna(0).astype(int)
    else:
        m['O/stock'] = 0

    # 4. 일자별 매칭으로 조정 TTL 계산
    plan_date_map = {col: parse_date_from_col(col) for col in plan_date_cols}
    valid_plan_cols = {col: dt for col, dt in plan_date_map.items() if dt is not None}

    def calc_adjusted_ttl(row):
        erp = row['ERP']
        adjusted = 0
        for col, dt in valid_plan_cols.items():
            if dt <= today:
                # 과거 또는 오늘: MES 실적 사용
                actual = daily_actual_dict.get((erp, dt.normalize()), 0)
                adjusted += actual
            else:
                # 미래: 계획 사용
                adjusted += row[col] if col in row else 0
        return int(adjusted)

    if valid_plan_cols:
        m['조정_TTL'] = m.apply(calc_adjusted_ttl, axis=1)
    else:
        # 일자별 계획 없으면 MES 누적실적 사용
        m['조정_TTL'] = m['MES_누적실적']
        st.warning("⚠️ 일자별 계획 컬럼이 없어 단순 누적실적 사용")

    # 5. 조정 BALANCE
    m['조정_BALANCE'] = m['O/stock'] + m['조정_TTL'] - m['TTL Ship']
    
    # 6. GAP = 조정 BAL - 계획 BAL
    if 'BALANCE' in m.columns:
        m['GAP'] = m['조정_BALANCE'] - m['BALANCE'].astype(int)
    else:
        m['GAP'] = 0

    # 7. 알람
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

    return m


# ════════════════════════════════════════════════════════════
# HTML 테이블 (빠름, 음수 빨강)
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

    parts = [f'<div style="max-height:{height}px;overflow:auto;border:1px solid #E5E7EB;border-radius:6px;">',
             '<table style="width:100%;border-collapse:collapse;font-family:-apple-system,sans-serif;font-size:12px;">',
             '<thead style="position:sticky;top:0;z-index:10;"><tr style="background-color:#374151;color:white;">']
    
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
    st.caption("📌 일자별 계획 vs MES 실적 정밀 매칭 | 과거=실적, 미래=계획")

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
        s2.success(f"📊 생산실적: **{len(prod_db)}건** | {prod_t}")
    else:
        s2.warning("📊 생산실적: 없음")

    st.markdown("##### 📤 파일 업로드")
    files = st.file_uploader(" ", type=["xlsx","csv"], accept_multiple_files=True,
                              key="ship_up_v20", label_visibility="collapsed")

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v20"):
            with st.spinner("처리 중..."):
                for f in files:
                    fname = f.name.lower()
                    f.seek(0)
                    file_bytes = f.read()
                    if fname.endswith('.xlsx'):
                        df, p_cols = load_shipment_rev(file_bytes)
                        if not df.empty:
                            st.session_state['ship_db'] = df
                            st.session_state['plan_date_cols'] = p_cols
                            st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                            st.success(f"✅ 출하계획 ({len(df)}건)")
                    elif fname.endswith('.csv'):
                        try:
                            df = read_csv_cached(file_bytes)
                            st.session_state['prod_db'] = df
                            st.session_state['prod_updated'] = datetime.now().strftime('%m-%d %H:%M')
                            st.success(f"✅ 생산실적 ({len(df)}건)")
                        except Exception as e:
                            st.error(f"❌ {f.name}: {e}")
            st.rerun()

    st.markdown("---")

    if ship_db.empty or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    with st.spinner("분석 중..."):
        try:
            m = analyze(ship_db, prod_db, plan_cols)
        except Exception as e:
            st.error(f"❌ 분석 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
            return

    if m.empty:
        return

    # KPI
    st.markdown("### 🚨 분석 결과")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(m)}건")
    k2.metric("🔴 긴급", int((m['알람']=='🔴 긴급').sum()))
    k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
    k4.metric("🟡 악화", int((m['알람']=='🟡 악화').sum()))
    k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    # 표시 컬럼
    show = ['알람','Cus','Cut off Cargo','HQ Request','model','code','ERP','MODEL_TYPE',
            'PO Remain','TTL Ship','O/stock','TTL Plan','TTLstock',
            'MES_누적실적','조정_TTL','BALANCE','조정_BALANCE','GAP','Note']
    show = [c for c in show if c in m.columns]

    # 긴급/부족
    urgent = m[m['알람'].isin(['🔴 긴급','🟠 부족','🟡 악화'])].copy()
    urgent = urgent.sort_values('조정_BALANCE')

    if not urgent.empty:
        st.markdown("---")
        st.markdown(f"#### 🚨 즉시 조치 필요 ({len(urgent)}건)")
