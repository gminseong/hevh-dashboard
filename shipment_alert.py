"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v23.0
- 목적: 계획 차질 감지 + Cut off 전 추가 생산 필요량 알람
- 핵심 컬럼: 생산실적 차이, 조정_BALANCE, 필요 수량 변경
- 4단계 알람: 출하불가/부족/차질/정상
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
    """컬럼명에서 날짜 추출"""
    try:
        s = str(col_name).lower()
        s = re.sub(r'(plan|actual|cut\s*off|cargo)[\.\s_&]*', '', s).strip()
        nums = re.findall(r'\d+', s)
        if len(nums) >= 2:
            month, day = int(nums[0]), int(nums[1])
            if 1 <= month <= 12 and 1 <= day <= 31:
                return pd.Timestamp(year, month, day)
        elif len(nums) == 1:
            day = int(nums[0])
            if 1 <= day <= 31:
                return pd.Timestamp(year, 6, day)
        return None
    except Exception:
        return None


def merge_two_row_header(raw, header_row):
    """2줄 헤더 병합"""
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
    return merged


# ════════════════════════════════════════════════════════════
# Sheet1 Note 추출
# ════════════════════════════════════════════════════════════
def load_sheet1_notes(file_bytes):
    try:
        sheet_names = list_sheets(file_bytes)
        target = None
        for s in sheet_names:
            if str(s).strip().lower() == 'sheet1':
                target = s
                break
        if target is None:
            return {}

        raw = read_excel_raw(file_bytes, target)
        header_row = None
        for i in range(min(8, len(raw))):
            row_vals = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
            if 'model' in row_vals and 'erp' in row_vals:
                header_row = i
                break
        if header_row is None:
            return {}

        merged = merge_two_row_header(raw, header_row)
        df = raw.iloc[header_row + 2:].copy()
        df.columns = merged
        df.columns = [str(c).strip() for c in df.columns]
        df = df.reset_index(drop=True)

        # Note 컬럼 찾기
        if 'Note' not in df.columns:
            for c in df.columns:
                if str(c).lower().strip() == 'note':
                    df = df.rename(columns={c: 'Note'})
                    break

        if 'ERP' not in df.columns or 'Note' not in df.columns:
            return {}

        df['ERP'] = df['ERP'].astype(str).str.strip()
        df = df[df['ERP'].str.startswith(('013', '018'))]
        return dict(zip(df['ERP'], df['Note'].fillna('')))
    except Exception:
        return {}


# ════════════════════════════════════════════════════════════
# Shipment Rev 로드
# ════════════════════════════════════════════════════════════
def load_shipment_rev(file_bytes):
    sheet_names = list_sheets(file_bytes)

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

    header_row = None
    for i in range(min(8, len(raw))):
        row_vals = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
        if 'cus' in row_vals:
            header_row = i
            break
    if header_row is None:
        header_row = 2

    merged = merge_two_row_header(raw, header_row)
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
            rename_map[c] = 'PO'
        elif 'ttl ship' in cl or 'ttl.ship' in cl:
            rename_map[c] = '_TTLShip'  # 임시 (PO 없을 때 백업)
        elif 'ttl plan' in cl or 'ttl.plan' in cl:
            rename_map[c] = '이번주 계획'
        elif 'o/stock' in cl or 'o.stock' in cl:
            rename_map[c] = '현재재고'
    df = df.rename(columns=rename_map)

    # Plan 일자별 컬럼
    plan_date_cols = []
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'plan' in cl and any(ch.isdigit() for ch in cl):
            if 'ttl' in cl or 'actual' in cl:
                continue
            if cl == '이번주 계획'.lower():
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

    df['ERP'] = df['ERP'].astype(str).str.strip()
    df = df[df['ERP'].str.startswith(('013', '018'))].copy()
    df = df.reset_index(drop=True)

    with st.expander("🔍 컬럼 매핑 진단", expanded=False):
        st.caption(f"전체 컬럼 ({len(df.columns)}개): {list(df.columns)}")
        st.caption(f"일자별 Plan 컬럼: {plan_date_cols}")

    return df, plan_date_cols


# ════════════════════════════════════════════════════════════
# 분석
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db, plan_date_cols, note_dict):
    prod_db = prod_db.copy()
    prod_db['TRAN_WORK_DATE'] = pd.to_datetime(prod_db['TRAN_WORK_DATE'], errors='coerce')
    today = prod_db['TRAN_WORK_DATE'].max()
    
    if pd.isna(today):
        st.error("❌ 생산실적 날짜 파싱 실패")
        return pd.DataFrame()

    st.info(f"📅 기준일: **{today.strftime('%Y-%m-%d')}** (MES 최신) | 이 일자까지=실적, 이후=계획 적용")

    # MES 실적 일자별
    prod_db['ERP'] = prod_db['FINAL_MAT_ID'].astype(str).str.strip()
    prod_db['TYPE'] = prod_db['ERP'].apply(classify)
    
    valid = prod_db[
        ((prod_db['TYPE']=='PD') & (prod_db['OPER_DESC']=='P-ATE')) |
        ((prod_db['TYPE']=='3IN1') & (prod_db['OPER_DESC']=='ASSY'))
    ].copy()

    daily = valid.groupby(['ERP', valid['TRAN_WORK_DATE'].dt.normalize()])['QTY'].sum().reset_index()
    daily.columns = ['ERP', 'DATE', 'QTY']
    daily_dict = {(r['ERP'], r['DATE']): r['QTY'] for _, r in daily.iterrows()}
    total_actual = valid.groupby('ERP')['QTY'].sum().to_dict()

    # Shipment Rev
    m = ship_db.copy()
    m['MODEL_TYPE'] = m['ERP'].apply(classify)
    m['현재실적'] = m['ERP'].map(total_actual).fillna(0).astype(int)
    m['Note'] = m['ERP'].map(note_dict).fillna('')

    # 숫자화
    for col in ['PO', '이번주 계획', '현재재고', '_TTLShip'] + plan_date_cols:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors='coerce').fillna(0)

    # PO 없으면 _TTLShip 사용
    if 'PO' not in m.columns and '_TTLShip' in m.columns:
        m['PO'] = m['_TTLShip']
    elif 'PO' in m.columns and '_TTLShip' in m.columns:
        m['PO'] = m.apply(lambda r: r['PO'] if r['PO'] > 0 else r['_TTLShip'], axis=1)

    if '현재재고' not in m.columns:
        m['현재재고'] = 0
    m['현재재고'] = m['현재재고'].astype(int)

    # 일자별 계획 매칭
    plan_date_map = {col: parse_date_from_col(col) for col in plan_date_cols}
    valid_plan_cols = {col: dt for col, dt in plan_date_map.items() if dt is not None}
    today_norm = today.normalize()

    def calc_metrics(row):
        erp = row['ERP']
        past_actual = 0
        past_plan = 0
        future_plan = 0
        
        for col, dt in valid_plan_cols.items():
            dt_norm = dt.normalize()
            plan_val = row.get(col, 0)
            if dt_norm <= today_norm:
                past_actual += daily_dict.get((erp, dt_norm), 0)
                past_plan += plan_val
            else:
                future_plan += plan_val
        
        # 생산실적 차이 = 과거 실적 - 과거 계획 (음수=계획 대비 부진)
        prod_gap = past_actual - past_plan
        # 조정 TTL = 과거 실적 + 미래 계획
        adjusted_ttl = past_actual + future_plan
        
        return int(prod_gap), int(adjusted_ttl)

    if valid_plan_cols:
        results = m.apply(calc_metrics, axis=1, result_type='expand')
        m['생산실적 차이'] = results[0]
        m['_조정TTL'] = results[1]
    else:
        m['생산실적 차이'] = 0
        m['_조정TTL'] = m['현재실적']
        st.warning("⚠️ 일자별 Plan 없음. 단순 누적실적 사용")

    # 조정_BALANCE = 현재재고 + 조정TTL - PO
    m['조정_BALANCE'] = (m['현재재고'] + m['_조정TTL'] - m['PO'].astype(int)).astype(int)
    
    # 필요 수량 변경 = 조정_BALANCE 0 만들기 위한 추가 생산
    m['필요 수량 변경'] = m['조정_BALANCE'].apply(lambda x: -x if x < 0 else 0).astype(int)

    # 알람 (4단계)
    def get_alert(row):
        adj_bal = row['조정_BALANCE']
        prod_gap = row['생산실적 차이']
        if adj_bal < -5000:
            return "🔴 출하불가"
        elif adj_bal < 0:
            return "🟠 부족"
        elif prod_gap < -1000:
            return "🟡 차질"
        else:
            return "✅ 정상"
    m['알람'] = m.apply(get_alert, axis=1)

    # 정수화
    for col in ['PO', '이번주 계획']:
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
        if '출하불가' in a: return '#FEE2E2'
        if '부족' in a: return '#FFEDD5'
        if '차질' in a: return '#FEF3C7'
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
    st.caption("📌 계획 차질 감지 + Cut off까지 추가 생산 필요량 자동 알람")

    ship_db = st.session_state.get('ship_db', pd.DataFrame())
    plan_cols = st.session_state.get('plan_date_cols', [])
    note_dict = st.session_state.get('note_dict', {})
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_t = st.session_state.get('ship_updated', '-')
    prod_t = st.session_state.get('prod_updated', '-')

    s1, s2 = st.columns(2)
    if not ship_db.empty:
        s1.success(f"📁 출하계획: **{len(ship_db)}건** | Note {len(note_dict)}건 | {ship_t}")
    else:
        s1.warning("📁 출하계획: 없음")
    if not prod_db.empty:
        s2.success(f"📊 생산실적: **{len(prod_db)}건** | {prod_t}")
    else:
        s2.warning("📊 생산실적: 없음")

    st.markdown("##### 📤 파일 업로드")
    files = st.file_uploader(" ", type=["xlsx","csv"], accept_multiple_files=True,
                              key="ship_up_v23", label_visibility="collapsed")

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v23"):
            with st.spinner("처리 중..."):
                for f in files:
                    fname = f.name.lower()
                    f.seek(0)
                    file_bytes = f.read()
                    if fname.endswith('.xlsx'):
                        df, p_cols = load_shipment_rev(file_bytes)
                        notes = load_sheet1_notes(file_bytes)
                        if not df.empty:
                            st.session_state['ship_db'] = df
                            st.session_state['plan_date_cols'] = p_cols
                            st.session_state['note_dict'] = notes
                            st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                            st.success(f"✅ 출하계획 ({len(df)}건) + Note ({len(notes)}건)")
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
            m = analyze(ship_db, prod_db, plan_cols, note_dict)
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
    k2.metric("🔴 출하불가", int((m['알람']=='🔴 출하불가').sum()))
    k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
    k4.metric("🟡 차질", int((m['알람']=='🟡 차질').sum()))
    k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    # 표시 컬럼 (목적 명확 버전)
    show = ['알람','Cus','Cut off Cargo','HQ Request','model','code','ERP','MODEL_TYPE',
            'PO','현재재고','이번주 계획','현재실적',
            '생산실적 차이','조정_BALANCE','필요 수량 변경','Note']
    show = [c for c in show if c in m.columns]

    # 긴급 (출하불가/부족/차질)
    urgent = m[m['알람'].isin(['🔴 출하불가','🟠 부족','🟡 차질'])].copy()
    urgent = urgent.sort_values(['알람', '조정_BALANCE'])

    if not urgent.empty:
        st.markdown("---")
        st.markdown(f"#### 🚨 즉시 조치 필요 ({len(urgent)}건)")
        st.caption("⚠️ 조정_BALANCE 음수 → 부족분 발생 | 필요 수량 변경 = Cut off까지 추가 생산 필요량")
        render_html_table(urgent[show], height=400)
    else:
        st.markdown("---")
        st.success("✅ 차질/부족 없음 - 모든 모델 정상")

    # 전체 상세
    st.markdown("---")
    st.markdown("#### 📋 전체 상세")
    
    f1, f2, f3 = st.columns(3)
    a_sel = f1.multiselect("알람", ['🔴 출하불가','🟠 부족','🟡 차질','✅ 정상'], key="a_v23")
    t_sel = f2.multiselect("모델", ['PD','3IN1','OTHER'], key="t_v23")
    cus_opt = sorted(m['Cus'].dropna().unique()) if 'Cus' in m.columns else []
    c_sel = f3.multiselect("거래선", cus_opt, key="c_v23") if cus_opt else []
    
    v = m.copy()
    if a_sel: v = v[v['알람'].isin(a_sel)]
    if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
    if c_sel: v = v[v['Cus'].isin(c_sel)]

    if not v.empty:
        render_html_table(v[show], height=550)
        csv = v[show].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 CSV 다운로드", csv,
                           f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           "text/csv", key="dl_v23")
    else:
        st.info("선택한 필터에 해당하는 데이터가 없습니다.")

    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v23"):
        for k in ['ship_db','ship_updated','plan_date_cols','note_dict']:
            st.session_state.pop(k, None)
        st.cache_data.clear()
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v23"):
        for k in ['prod_db','prod_updated']:
            st.session_state.pop(k, None)
        st.cache_data.clear()
        st.rerun()
