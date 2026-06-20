"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v13.0
- v11 베이스 (검증된 Sheet1 로드)
- 새 디자인: KPI → 긴급/부족 강조 → 전체 상세
- 그레이 테마 + 음수 빨강
"""
from datetime import datetime
import pandas as pd
import streamlit as st


# ════════════════════════════════════════════════════════════
# Sheet1 로드 (v11 검증 방식)
# ════════════════════════════════════════════════════════════
def load_sheet1(file):
    try:
        file.seek(0)
        df = pd.read_excel(file, sheet_name='Sheet1')
        df.columns = [str(c).strip() for c in df.columns]
        if 'model' in df.columns:
            df = df[df['model'].notna()].copy()
        if 'ERP' in df.columns:
            df['ERP'] = df['ERP'].astype(str).str.strip()
        return df
    except Exception:
        return pd.DataFrame()


# ════════════════════════════════════════════════════════════
# 분석
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db):
    # 누적실적 집계
    p = prod_db.copy()
    p['ERP_STR'] = p['FINAL_MAT_ID'].astype(str).str.strip()
    p['TYPE'] = p['ERP_STR'].apply(
        lambda x: '3IN1' if x.startswith('018') else ('PD' if x.startswith('01') else 'OTHER')
    )
    pd_a = p[(p['TYPE']=='PD') & (p['OPER_DESC']=='P-ATE')].groupby('ERP_STR')['QTY'].sum()
    in1_a = p[(p['TYPE']=='3IN1') & (p['OPER_DESC']=='ASSY')].groupby('ERP_STR')['QTY'].sum()
    actual_dict = pd.concat([pd_a, in1_a]).to_dict()

    # 머지
    m = ship_db.copy()
    m['ERP'] = m['ERP'].astype(str).str.strip()
    m['MODEL_TYPE'] = m['ERP'].apply(
        lambda x: '3IN1' if x.startswith('018') else ('PD' if x.startswith('01') else 'OTHER')
    )
    m['누적실적'] = m['ERP'].map(actual_dict).fillna(0).astype(int)

    # 기초재고
    stock_col = None
    for c in m.columns:
        if 'base stock' in str(c).lower() or 'stock' in str(c).lower():
            stock_col = c
            break
    if stock_col:
        m['기초재고'] = pd.to_numeric(m[stock_col], errors='coerce').fillna(0).astype(int)
    else:
        m['기초재고'] = 0

    # TTLstock = 기초재고 + 누적실적
    m['TTLstock'] = m['기초재고'] + m['누적실적']

    # SO, Plan
    m['SO'] = pd.to_numeric(m.get('SO', 0), errors='coerce').fillna(0).astype(int)
    
    plan_col = None
    for c in m.columns:
        if 'plan' in str(c).lower():
            plan_col = c
            break
    if plan_col:
        m['Plan.W25'] = pd.to_numeric(m[plan_col], errors='coerce').fillna(0).astype(int)
    else:
        m['Plan.W25'] = 0

    # NEW_BALANCE = TTLstock + Plan - SO
    m['NEW_BALANCE'] = m['TTLstock'] + m['Plan.W25'] - m['SO']

    # 달성률
    rate = m['누적실적'] / m['Plan.W25'].where(m['Plan.W25'] != 0)
    rate = pd.to_numeric(rate, errors='coerce').replace([float('inf'), -float('inf')], 0).fillna(0)
    m['달성률(%)'] = (rate * 100).round(1)

    # 알람
    def get_alert(row):
        bal = row['NEW_BALANCE']
        rate = row['달성률(%)']
        if bal < -5000:
            return "🔴 긴급"
        elif bal < 0:
            return "🟠 부족"
        elif rate < 70 and row['Plan.W25'] > 0:
            return "🟡 지연"
        else:
            return "✅ 정상"
    m['알람'] = m.apply(get_alert, axis=1)
    
    return m


# ════════════════════════════════════════════════════════════
# 스타일
# ════════════════════════════════════════════════════════════
def make_styled_df(df, numeric_cols):
    def hl_neg(val):
        try:
            if float(val) < 0:
                return 'color: #DC2626; font-weight: bold;'
        except Exception:
            pass
        return ''

    def hl_row(row):
        if '알람' not in row.index:
            return [''] * len(row)
        a = str(row['알람'])
        if '긴급' in a:
            bg = '#FEE2E2'
        elif '부족' in a:
            bg = '#FFEDD5'
        elif '지연' in a:
            bg = '#FEF3C7'
        else:
            bg = '#F9FAFB'
        return [f'background-color: {bg};'] * len(row)

    try:
        styled = df.style.apply(hl_row, axis=1)
        valid = [c for c in numeric_cols if c in df.columns]
        if valid:
            styled = styled.applymap(hl_neg, subset=valid)
        styled = styled.set_table_styles([
            {'selector': 'thead th',
             'props': [('background-color', '#374151'),
                       ('color', 'white'),
                       ('font-weight', 'bold'),
                       ('text-align', 'center'),
                       ('padding', '10px 8px')]},
            {'selector': 'tbody td',
             'props': [('padding', '8px'),
                       ('border-bottom', '1px solid #E5E7EB'),
                       ('text-align', 'center')]},
        ])
        return styled
    except Exception:
        return df


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD: P-ATE 기준 | 3in1: ASSY 기준 | NEW_BALANCE = TTLstock + Plan - SO")

    # 데이터 현황
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

    # 업로드
    st.markdown("##### 📤 파일 업로드")
    files = st.file_uploader(
        " ", type=["xlsx", "csv"], accept_multiple_files=True,
        key="ship_up_v13", label_visibility="collapsed"
    )
    
    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True):
            for f in files:
                fname = f.name.lower()
                if fname.endswith('.xlsx'):
                    df = load_sheet1(f)
                    if df.empty:
                        st.error(f"❌ {f.name}: Sheet1 로드 실패 (시트가 없거나 빈 데이터)")
                    elif 'ERP' not in df.columns:
                        st.error(f"❌ {f.name}: ERP 컬럼 없음. 보유 컬럼: {list(df.columns)}")
                    else:
                        st.session_state['ship_db'] = df
                        st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        st.success(f"✅ 출하계획 저장 ({len(df)}건)")
                elif fname.endswith('.csv'):
                    try:
                        df = pd.read_csv(f)
                        st.session_state['prod_db'] = df
                        st.session_state['prod_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        st.success(f"✅ 생산실적 저장 ({len(df)}건)")
                    except Exception as e:
                        st.error(f"❌ {f.name}: {e}")
            # rerun 제거 - 메시지 보이게
            # st.rerun()

    st.markdown("---")
    
    if ship_db.empty or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    # 분석
    try:
        m = analyze(ship_db, prod_db)
    except Exception as e:
        st.error(f"❌ 분석 오류: {e}")
        import traceback
        st.code(traceback.format_exc())
        return

    # KPI
    st.markdown("### 🚨 분석 결과")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(m)}건")
    k2.metric("🔴 긴급", int((m['알람']=='🔴 긴급').sum()))
    k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
    k4.metric("🟡 지연", int((m['알람']=='🟡 지연').sum()))
    k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    numeric_cols = ['SO', '기초재고', '누적실적', 'TTLstock', 'Plan.W25',
                    '달성률(%)', 'NEW_BALANCE']

    # 긴급/부족 강조
    urgent_short = m[m['알람'].isin(['🔴 긴급', '🟠 부족'])].copy()
    urgent_short = urgent_short.sort_values('NEW_BALANCE')

    if not urgent_short.empty:
        st.markdown("---")
        st.markdown(f"#### 🚨 즉시 조치 필요 ({len(urgent_short)}건)")
        st.caption("⚠️ NEW_BALANCE 음수 = 출하 부족분 발생 예상")
        
        show = ['알람','model','code','ERP','MODEL_TYPE',
                'SO','TTLstock','Plan.W25','NEW_BALANCE','Note']
        show = [c for c in show if c in urgent_short.columns]
        
        try:
            st.dataframe(
                make_styled_df(urgent_short[show], numeric_cols),
                use_container_width=True,
                height=min(50 + len(urgent_short) * 38, 450)
            )
        except Exception:
            st.dataframe(urgent_short[show], use_container_width=True)
    else:
        st.markdown("---")
        st.success("✅ 긴급/부족 모델 없음 — 모든 모델이 정상 상태입니다.")

    # 전체 상세
    st.markdown("---")
    st.markdown("#### 📋 전체 상세")
    
    f1, f2 = st.columns(2)
    a_sel = f1.multiselect("알람 등급", 
                            ['🔴 긴급','🟠 부족','🟡 지연','✅ 정상'],
                            key="all_a_v13")
    t_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'], key="all_t_v13")
    
    v = m.copy()
    if a_sel: v = v[v['알람'].isin(a_sel)]
    if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
    
    all_cols = ['알람','model','code','ERP','MODEL_TYPE',
                'SO','기초재고','누적실적','TTLstock','Plan.W25',
                '달성률(%)','NEW_BALANCE','Note']
    all_cols = [c for c in all_cols if c in v.columns]
    
    if not v.empty:
        try:
            st.dataframe(
                make_styled_df(v[all_cols], numeric_cols),
                use_container_width=True, height=500
            )
        except Exception:
            st.dataframe(v[all_cols], use_container_width=True, height=500)
        
        csv = v[all_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 CSV 다운로드", csv,
            f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv"
        )
    else:
        st.info("선택한 필터에 해당하는 데이터가 없습니다.")

    # 초기화
    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v13"):
        st.session_state.pop('ship_db', None)
        st.session_state.pop('ship_updated', None)
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v13"):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
