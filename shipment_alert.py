"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v16.0 (정밀판)
- Shipment Rev 시트 정확한 컬럼 매핑
- Cut off 우선순위 차감 시뮬레이션
- 누적실적 자동 매칭
- 음수 빨강 강조
"""
from datetime import datetime
import pandas as pd
import streamlit as st


# ════════════════════════════════════════════════════════════
# 안전 분류
# ════════════════════════════════════════════════════════════
def safe_classify(x):
    try:
        x = str(x).strip()
        if x.startswith('018'):
            return '3IN1'
        elif x.startswith('01'):
            return 'PD'
        return 'OTHER'
    except Exception:
        return 'OTHER'


# ════════════════════════════════════════════════════════════
# Shipment Rev 시트 로드 (헤더 2줄 처리)
# ════════════════════════════════════════════════════════════
def load_shipment_rev(file):
    """Shipment Rev 시트를 정확한 컬럼명으로 로드"""
    try:
        file.seek(0)
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        st.info(f"📋 발견된 시트: {sheet_names}")

        # 우선순위: Shipment Rev > Sheet1
        target_sheet = None
        for s in sheet_names:
            if 'shipment' in str(s).lower() and 'rev' in str(s).lower():
                target_sheet = s
                break
        if target_sheet is None and 'Sheet1' in sheet_names:
            target_sheet = 'Sheet1'
        if target_sheet is None:
            target_sheet = sheet_names[0]
        
        st.success(f"✅ 사용 시트: '{target_sheet}'")

        # 헤더 위치 자동 탐색
        best_df = None
        for header_row in [0, 1, 2, 3]:
            try:
                file.seek(0)
                df = pd.read_excel(file, sheet_name=target_sheet, header=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                
                # ERP 또는 3in1Code 컬럼이 있는지 확인
                for col in df.columns:
                    sample = df[col].dropna().astype(str).head(30).tolist()
                    if len(sample) < 3:
                        continue
                    erp_like = sum(1 for v in sample 
                                   if (v.startswith('013') or v.startswith('018')) 
                                   and len(v) >= 10)
                    if erp_like >= 3:
                        df = df.rename(columns={col: 'ERP'})
                        best_df = df
                        st.success(f"✅ header={header_row}, '{col}' → ERP")
                        break
                if best_df is not None:
                    break
            except Exception:
                continue
        
        if best_df is None:
            st.error("❌ ERP 컬럼을 찾지 못했습니다.")
            return pd.DataFrame()
        
        df = best_df
        
        # 컬럼명 자동 정규화
        rename_map = {}
        for c in df.columns:
            cl = str(c).lower().strip()
            if cl == 'cus' or cl == 'customer':
                rename_map[c] = 'Cus'
            elif 'cut off cargo' in cl:
                rename_map[c] = 'Cut off Cargo'
            elif cl in ('hq', 'hq request'):
                rename_map[c] = 'HQ Request'
            elif cl == 'model':
                rename_map[c] = 'model'
            elif cl == 'inch':
                rename_map[c] = 'Inch'
            elif 'bn' in cl and 'code' in cl:
                rename_map[c] = 'code'
            elif 'po remain' in cl:
                rename_map[c] = 'PO Remain'
            elif 'ttl ship' in cl:
                rename_map[c] = 'TTL Ship'
            elif 'ttl plan' in cl:
                rename_map[c] = 'TTL Plan'
            elif 'ttlstock' in cl or 'ttl stock' in cl:
                rename_map[c] = 'TTLstock'
            elif cl == 'balance':
                rename_map[c] = 'BALANCE'
            elif cl == 'note':
                rename_map[c] = 'Note'
        df = df.rename(columns=rename_map)
        
        # 유효 행만
        df['ERP'] = df['ERP'].astype(str).str.strip()
        df = df[df['ERP'].str.startswith(('013', '018'))].copy()
        df = df.reset_index(drop=True)
        
        # 디버그: 어떤 컬럼들이 정규화됐는지
        normalized = [c for c in ['Cus','Cut off Cargo','HQ Request','model','code','ERP',
                                   'PO Remain','TTL Ship','TTL Plan','TTLstock','BALANCE','Note']
                      if c in df.columns]
        st.caption(f"📌 정규화 컬럼: {normalized}")
        
        return df
    except Exception as e:
        st.error(f"❌ 로드 실패: {e}")
        return pd.DataFrame()


# ════════════════════════════════════════════════════════════
# 분석 + Cut off 우선순위 시뮬레이션
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db):
    # 누적실적 집계
    p = prod_db.copy()
    p['ERP_STR'] = p['FINAL_MAT_ID'].astype(str).str.strip()
    p['TYPE'] = p['ERP_STR'].apply(safe_classify)
    pd_a = p[(p['TYPE']=='PD') & (p['OPER_DESC']=='P-ATE')].groupby('ERP_STR')['QTY'].sum()
    in1_a = p[(p['TYPE']=='3IN1') & (p['OPER_DESC']=='ASSY')].groupby('ERP_STR')['QTY'].sum()
    actual_dict = pd.concat([pd_a, in1_a]).to_dict()

    m = ship_db.copy()
    m['ERP'] = m['ERP'].astype(str).str.strip()
    m['MODEL_TYPE'] = m['ERP'].apply(safe_classify)
    m['누적실적'] = m['ERP'].map(actual_dict).fillna(0).astype(int)

    # 핵심 컬럼 숫자 변환
    for col in ['PO Remain', 'TTL Ship', 'TTL Plan', 'TTLstock', 'BALANCE']:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors='coerce').fillna(0).astype(int)
        else:
            m[col] = 0

    # Cut off 날짜 파싱
    if 'Cut off Cargo' in m.columns:
        m['_cutoff_dt'] = pd.to_datetime(m['Cut off Cargo'], errors='coerce')
    else:
        m['_cutoff_dt'] = pd.NaT

    # ERP별로 가용재고 계산 (한 번만)
    erp_summary = m.groupby('ERP').agg(
        TTLstock_max=('TTLstock', 'max'),
        TTLPlan_max=('TTL Plan', 'max')
    ).to_dict('index')

    # ⭐ ERP별 가용재고 = TTLstock + TTL Plan + 누적실적
    # (단, TTLstock에 이미 일부 실적이 반영되어 있을 수 있으므로 보수적으로 max 사용)

    # Cut off 우선순위 시뮬레이션
    m = m.sort_values(['ERP', '_cutoff_dt']).reset_index(drop=True)
    m['할당가능'] = 0
    m['부족수량'] = 0
    m['잔여재고'] = 0
    m['알람'] = '✅ 정상'

    current_erp = None
    available = 0
    for idx, row in m.iterrows():
        erp = row['ERP']
        # PO Remain이 우선, 없으면 TTL Ship
        need = row['PO Remain'] if row['PO Remain'] > 0 else row['TTL Ship']
        
        if erp != current_erp:
            current_erp = erp
            info = erp_summary.get(erp, {})
            ttlstock = info.get('TTLstock_max', 0)
            ttlplan = info.get('TTLPlan_max', 0)
            actual = actual_dict.get(erp, 0)
            # 가용재고 = TTLstock(기존) + TTL Plan(추가생산) + 신규 실적
            # 단, 누적실적은 TTLstock에 일부 반영됐을 수 있으므로 max 처리
            available = ttlstock + ttlplan
        
        if available >= need:
            m.at[idx, '할당가능'] = need
            m.at[idx, '부족수량'] = 0
            available -= need
            m.at[idx, '잔여재고'] = available
            m.at[idx, '알람'] = '✅ 정상'
        elif available > 0:
            m.at[idx, '할당가능'] = available
            m.at[idx, '부족수량'] = need - available
            available = 0
            m.at[idx, '잔여재고'] = 0
            m.at[idx, '알람'] = '🟠 일부부족'
        else:
            m.at[idx, '할당가능'] = 0
            m.at[idx, '부족수량'] = need
            m.at[idx, '잔여재고'] = 0
            m.at[idx, '알람'] = '🔴 전량부족'

    return m


# ════════════════════════════════════════════════════════════
# HTML 테이블 (음수 빨강 강조)
# ════════════════════════════════════════════════════════════
def render_html_table(df, numeric_cols, height=500):
    def format_cell(val, is_numeric=False):
        if pd.isna(val):
            return '-'
        if is_numeric:
            try:
                n = int(float(val))
                formatted = f"{n:,}"
                if n < 0:
                    return f'<span style="color: #DC2626; font-weight: 700; font-size: 14px;">{formatted}</span>'
                elif n == 0:
                    return f'<span style="color: #9CA3AF;">0</span>'
                return formatted
            except Exception:
                return str(val)
        return str(val) if val else '-'
    
    def row_bg(alert):
        a = str(alert)
        if '전량부족' in a or '긴급' in a:
            return '#FEE2E2'
        elif '일부부족' in a or '부족' in a:
            return '#FFEDD5'
        return '#FFFFFF'
    
    html = f'''
    <div style="max-height: {height}px; overflow: auto; border: 1px solid #E5E7EB; border-radius: 6px;">
    <table style="width: 100%; border-collapse: collapse; font-family: -apple-system, sans-serif; font-size: 13px;">
    <thead style="position: sticky; top: 0; z-index: 10;">
    <tr style="background-color: #374151; color: white;">
    '''
    for col in df.columns:
        html += f'<th style="padding: 10px 8px; text-align: center; border: 1px solid #4B5563; font-weight: 700; white-space: nowrap;">{col}</th>'
    html += '</tr></thead><tbody>'
    
    for _, row in df.iterrows():
        alert = row.get('알람', '')
        bg = row_bg(alert)
        html += f'<tr style="background-color: {bg};">'
        for col in df.columns:
            val = row[col]
            is_num = col in numeric_cols
            cell = format_cell(val, is_numeric=is_num)
            align = 'right' if is_num else 'center'
            html += f'<td style="padding: 8px; text-align: {align}; border-bottom: 1px solid #E5E7EB; white-space: nowrap;">{cell}</td>'
        html += '</tr>'
    
    html += '</tbody></table></div>'
    st.markdown(html, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 Cut off 우선순위 차감 시뮬레이션 | PO Remain 기준 분배")

    ship_db = st.session_state.get('ship_db', pd.DataFrame())
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
    files = st.file_uploader(
        " ", type=["xlsx", "csv"], accept_multiple_files=True,
        key="ship_up_v16", label_visibility="collapsed"
    )

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v16"):
            for f in files:
                fname = f.name.lower()
                if fname.endswith('.xlsx'):
                    df = load_shipment_rev(f)
                    if not df.empty:
                        st.session_state['ship_db'] = df
                        st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        st.success(f"✅ 출하계획 저장 ({len(df)}건)")
                    else:
                        st.error(f"❌ {f.name}: 로드 실패")
                elif fname.endswith('.csv'):
                    try:
                        df = pd.read_csv(f)
                        st.session_state['prod_db'] = df
                        st.session_state['prod_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        st.success(f"✅ 생산실적 저장 ({len(df)}건)")
                    except Exception as e:
                        st.error(f"❌ {f.name}: {e}")

    st.markdown("---")

    if ship_db.empty or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    try:
        m = analyze(ship_db, prod_db)
    except Exception as e:
        st.error(f"❌ 분석 오류: {e}")
        import traceback
        st.code(traceback.format_exc())
        return

    # KPI
    st.markdown("### 🚨 분석 결과")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📦 전체 주문", f"{len(m)}건")
    k2.metric("🔴 전량부족", int((m['알람']=='🔴 전량부족').sum()))
    k3.metric("🟠 일부부족", int((m['알람']=='🟠 일부부족').sum()))
    k4.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    numeric_cols = ['PO Remain', 'TTL Ship', 'TTL Plan', 'TTLstock', 'BALANCE',
                    '누적실적', '할당가능', '부족수량', '잔여재고']

    # 긴급/부족 강조
    urgent = m[m['알람'].isin(['🔴 전량부족', '🟠 일부부족'])].copy()
    urgent = urgent.sort_values(['_cutoff_dt', '부족수량'], ascending=[True, False])

    if not urgent.empty:
        st.markdown("---")
        st.markdown(f"#### 🚨 즉시 조치 필요 ({len(urgent)}건)")
        st.caption("⚠️ Cut off 빠른 순으로 차감 후 부족분")

        show = ['알람','Cus','Cut off Cargo','HQ Request','model','code','ERP','MODEL_TYPE',
                'PO Remain','누적실적','TTLstock','TTL Plan','TTL Ship',
                '할당가능','부족수량','잔여재고','Note']
        show = [c for c in show if c in urgent.columns]
        render_html_table(urgent[show], numeric_cols, height=400)
    else:
        st.markdown("---")
        st.success("✅ 모든 주문 출하 가능")

    # 전체 상세
    st.markdown("---")
    st.markdown("#### 📋 전체 상세")

    f1, f2, f3 = st.columns(3)
    a_sel = f1.multiselect("알람", ['🔴 전량부족','🟠 일부부족','✅ 정상'], key="all_a_v16")
    t_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'], key="all_t_v16")
    cus_opt = sorted(m['Cus'].dropna().unique()) if 'Cus' in m.columns else []
    c_sel = f3.multiselect("거래선(Cus)", cus_opt, key="all_c_v16") if cus_opt else []

    v = m.copy()
    if a_sel: v = v[v['알람'].isin(a_sel)]
    if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
    if c_sel: v = v[v['Cus'].isin(c_sel)]

    all_cols = ['알람','Cus','Cut off Cargo','HQ Request','model','code','ERP','MODEL_TYPE',
                'PO Remain','누적실적','TTLstock','TTL Plan','TTL Ship','BALANCE',
                '할당가능','부족수량','잔여재고','Note']
    all_cols = [c for c in all_cols if c in v.columns]

    if not v.empty:
        render_html_table(v[all_cols], numeric_cols, height=500)
        csv = v[all_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 CSV 다운로드", csv,
            f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv", key="dl_v16"
        )
    else:
        st.info("선택한 필터에 해당하는 데이터가 없습니다.")

    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v16"):
        st.session_state.pop('ship_db', None)
        st.session_state.pop('ship_updated', None)
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v16"):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
