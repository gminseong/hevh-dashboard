"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v15.0
- 2가지 BALANCE (계획 vs 실적)
- 누적실적/Plan 모두 표시
- 음수 빨강 강화 (HTML 직접 렌더링)
- 모든 시트 자동 탐색
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
# 출하계획 로드
# ════════════════════════════════════════════════════════════
def load_sheet1(file):
    """모든 시트/헤더 탐색하여 ERP 컬럼 자동 발견"""
    try:
        file.seek(0)
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        st.info(f"📋 발견된 시트: {sheet_names}")

        # 1차: Sheet1 직접 시도
        if 'Sheet1' in sheet_names:
            file.seek(0)
            df = pd.read_excel(file, sheet_name='Sheet1')
            df.columns = [str(c).strip() for c in df.columns]
            if 'ERP' in df.columns:
                sample = df['ERP'].dropna().astype(str).head(20).tolist()
                erp_like = sum(1 for v in sample 
                               if (v.startswith('013') or v.startswith('018')) 
                               and len(v) >= 10)
                if erp_like >= 3:
                    st.success("✅ Sheet1 사용 (깔끔한 구조)")
                    if 'model' in df.columns:
                        df = df[df['model'].notna()].copy()
                    df['ERP'] = df['ERP'].astype(str).str.strip()
                    df = df[df['ERP'].str.startswith(('013', '018'))].copy()
                    return df

        # 2차: 모든 시트 × 헤더 탐색
        best_df = pd.DataFrame()
        best_info = None
        best_score = 0

        for sheet_name in sheet_names:
            for header_row in [0, 1, 2, 3]:
                try:
                    file.seek(0)
                    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
                    df.columns = [str(c).strip() for c in df.columns]

                    if 'ERP' in df.columns:
                        sample = df['ERP'].dropna().astype(str).head(20).tolist()
                        erp_like = sum(1 for v in sample 
                                       if (v.startswith('013') or v.startswith('018')) 
                                       and len(v) >= 10)
                        if erp_like >= 3:
                            st.success(f"✅ '{sheet_name}' (header={header_row}): ERP 직접 발견")
                            if 'model' in df.columns:
                                df = df[df['model'].notna()].copy()
                            df['ERP'] = df['ERP'].astype(str).str.strip()
                            df = df[df['ERP'].str.startswith(('013', '018'))].copy()
                            return df

                    for col in df.columns:
                        sample = df[col].dropna().astype(str).head(30).tolist()
                        if len(sample) < 3:
                            continue
                        erp_like = sum(1 for v in sample 
                                       if (str(v).startswith('013') or str(v).startswith('018')) 
                                       and len(str(v)) >= 10)
                        if erp_like > best_score:
                            best_score = erp_like
                            best_info = (sheet_name, header_row, col)
                            best_df = df.copy()
                except Exception:
                    continue

        if best_score >= 3 and best_info:
            sheet_name, header_row, erp_col = best_info
            st.success(f"✅ 자동 매핑: '{sheet_name}' (header={header_row}), '{erp_col}' → 'ERP'")
            best_df = best_df.rename(columns={erp_col: 'ERP'})
            best_df['ERP'] = best_df['ERP'].astype(str).str.strip()

            for c in list(best_df.columns):
                if str(c).lower().strip() == 'model':
                    if c != 'model':
                        best_df = best_df.rename(columns={c: 'model'})
                    break

            if 'model' in best_df.columns:
                best_df = best_df[best_df['model'].notna()].copy()
            best_df = best_df[best_df['ERP'].str.startswith(('013', '018'))].copy()
            return best_df

        st.error("❌ ERP 컬럼을 찾지 못했습니다.")
        return pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Excel 로드 실패: {e}")
        return pd.DataFrame()


# ════════════════════════════════════════════════════════════
# 컬럼 자동 매핑 헬퍼
# ════════════════════════════════════════════════════════════
def find_column(df, patterns):
    """patterns 리스트의 키워드를 모두 포함하는 컬럼 찾기"""
    for c in df.columns:
        cl = str(c).lower()
        if all(p in cl for p in patterns):
            return c
    return None


# ════════════════════════════════════════════════════════════
# 분석
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

    # 기초재고 (base stock 우선, 없으면 stock 키워드)
    stock_col = find_column(m, ['base', 'stock'])
    if stock_col is None:
        stock_col = find_column(m, ['ttlstock'])
    if stock_col is None:
        for c in m.columns:
            if 'stock' in str(c).lower():
                stock_col = c
                break
    if stock_col:
        m['기초재고'] = pd.to_numeric(m[stock_col], errors='coerce').fillna(0).astype(int)
    else:
        m['기초재고'] = 0

    # TTLstock = 기초재고 + 누적실적
    m['TTLstock'] = m['기초재고'] + m['누적실적']

    # SO
    so_col = None
    for c in m.columns:
        cl = str(c).lower().strip()
        if cl == 'so' or 'ttl ship' in cl or cl == 'ship':
            so_col = c
            break
    if so_col:
        m['SO'] = pd.to_numeric(m[so_col], errors='coerce').fillna(0).astype(int)
    else:
        m['SO'] = 0

    # Plan (W25 우선)
    plan_col = find_column(m, ['plan', 'w'])
    if plan_col is None:
        plan_col = find_column(m, ['ttl', 'plan'])
    if plan_col is None:
        for c in m.columns:
            if 'plan' in str(c).lower():
                plan_col = c
                break
    if plan_col:
        m['Plan'] = pd.to_numeric(m[plan_col], errors='coerce').fillna(0).astype(int)
        st.caption(f"📌 Plan 컬럼: `{plan_col}`")
    else:
        m['Plan'] = 0
        st.warning("⚠️ Plan 컬럼을 찾지 못했습니다. 0으로 처리됩니다.")

    # ⭐ 2가지 BALANCE
    # 1) BAL_계획: 계획까지 다 채웠을 때
    m['BAL_계획'] = m['TTLstock'] + m['Plan'] - m['SO']
    # 2) BAL_실적: 현재 실적만 (지금 당장 출하 가능 여부)
    m['BAL_실적'] = m['TTLstock'] - m['SO']

    # 달성률
    rate = m['누적실적'] / m['Plan'].where(m['Plan'] != 0)
    rate = pd.to_numeric(rate, errors='coerce').replace([float('inf'), -float('inf')], 0).fillna(0)
    m['달성률(%)'] = (rate * 100).round(1)

    # 알람 (BAL_실적 우선 + BAL_계획 보조)
    def get_alert(row):
        bal_real = row['BAL_실적']
        bal_plan = row['BAL_계획']
        # 계획대로 해도 부족 → 긴급
        if bal_plan < -5000:
            return "🔴 긴급"
        elif bal_plan < 0:
            return "🟠 부족"
        # 실적 기준 부족 (계획은 OK)
        elif bal_real < 0:
            return "🟡 실적부족"
        else:
            return "✅ 정상"
    m['알람'] = m.apply(get_alert, axis=1)

    # code, Note 자동 매핑
    if 'code' not in m.columns:
        for c in m.columns:
            cl = str(c).lower()
            if 'bn' in cl and 'code' in cl:
                m = m.rename(columns={c: 'code'})
                break
    if 'Note' not in m.columns:
        for c in m.columns:
            if str(c).lower() == 'note':
                m = m.rename(columns={c: 'Note'})
                break

    return m


# ════════════════════════════════════════════════════════════
# HTML 테이블 직접 렌더링 (강한 음수 색상)
# ════════════════════════════════════════════════════════════
def render_html_table(df, numeric_cols, height=500):
    """HTML 테이블로 직접 렌더링 (Streamlit Styler 한계 우회)"""
    
    def format_cell(val, is_numeric=False, is_negative=False):
        if pd.isna(val):
            return ''
        if is_numeric:
            try:
                n = int(float(val))
                formatted = f"{n:,}"
                if n < 0:
                    return f'<span style="color: #DC2626; font-weight: 700; font-size: 14px;">{formatted}</span>'
                return formatted
            except Exception:
                return str(val)
        return str(val)
    
    def row_bg(alert):
        if '긴급' in str(alert):
            return '#FEE2E2'
        elif '부족' in str(alert) and '실적' not in str(alert):
            return '#FFEDD5'
        elif '실적부족' in str(alert):
            return '#FEF3C7'
        return '#FFFFFF'
    
    # HTML 빌드
    html = f'''
    <div style="max-height: {height}px; overflow-y: auto; border: 1px solid #E5E7EB; border-radius: 6px;">
    <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px;">
    <thead style="position: sticky; top: 0; z-index: 10;">
    <tr style="background-color: #374151; color: white;">
    '''
    
    for col in df.columns:
        html += f'<th style="padding: 10px 8px; text-align: center; border: 1px solid #4B5563; font-weight: 700;">{col}</th>'
    html += '</tr></thead><tbody>'
    
    for _, row in df.iterrows():
        alert = row.get('알람', '')
        bg = row_bg(alert)
        html += f'<tr style="background-color: {bg};">'
        for col in df.columns:
            val = row[col]
            is_num = col in numeric_cols
            cell_html = format_cell(val, is_numeric=is_num)
            align = 'right' if is_num else 'center'
            html += f'<td style="padding: 8px; text-align: {align}; border-bottom: 1px solid #E5E7EB;">{cell_html}</td>'
        html += '</tr>'
    
    html += '</tbody></table></div>'
    
    st.markdown(html, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD: P-ATE | 3in1: ASSY | BAL_계획 = TTLstock + Plan - SO | BAL_실적 = TTLstock - SO")

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
        key="ship_up_v15", label_visibility="collapsed"
    )

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v15"):
            for f in files:
                fname = f.name.lower()
                if fname.endswith('.xlsx'):
                    df = load_sheet1(f)
                    if df.empty:
                        st.error(f"❌ {f.name}: 로드 실패")
                    elif 'ERP' not in df.columns:
                        st.error(f"❌ {f.name}: ERP 컬럼 없음")
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
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(m)}건")
    k2.metric("🔴 긴급", int((m['알람']=='🔴 긴급').sum()))
    k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
    k4.metric("🟡 실적부족", int((m['알람']=='🟡 실적부족').sum()))
    k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    numeric_cols = ['SO', '기초재고', '누적실적', 'TTLstock', 'Plan',
                    '달성률(%)', 'BAL_계획', 'BAL_실적']

    # 긴급/부족 강조
    urgent = m[m['알람'].isin(['🔴 긴급', '🟠 부족', '🟡 실적부족'])].copy()
    urgent = urgent.sort_values('BAL_계획')

    if not urgent.empty:
        st.markdown("---")
        st.markdown(f"#### 🚨 즉시 조치 필요 ({len(urgent)}건)")
        st.caption("⚠️ BAL_계획 = 계획까지 다 채웠을 때 | BAL_실적 = 현재 실적만으로")

        show = ['알람','model','code','ERP','MODEL_TYPE',
                'SO','기초재고','누적실적','TTLstock','Plan',
                '달성률(%)','BAL_계획','BAL_실적','Note']
        show = [c for c in show if c in urgent.columns]
        render_html_table(urgent[show], numeric_cols, height=400)
    else:
        st.markdown("---")
        st.success("✅ 긴급/부족 모델 없음")

    # 전체 상세
    st.markdown("---")
    st.markdown("#### 📋 전체 상세")

    f1, f2 = st.columns(2)
    a_sel = f1.multiselect("알람 등급",
                           ['🔴 긴급','🟠 부족','🟡 실적부족','✅ 정상'],
                           key="all_a_v15")
    t_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'], key="all_t_v15")

    v = m.copy()
    if a_sel: v = v[v['알람'].isin(a_sel)]
    if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]

    all_cols = ['알람','model','code','ERP','MODEL_TYPE',
                'SO','기초재고','누적실적','TTLstock','Plan',
                '달성률(%)','BAL_계획','BAL_실적','Note']
    all_cols = [c for c in all_cols if c in v.columns]

    if not v.empty:
        render_html_table(v[all_cols], numeric_cols, height=500)

        csv = v[all_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 CSV 다운로드", csv,
            f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv", key="dl_v15"
        )
    else:
        st.info("선택한 필터에 해당하는 데이터가 없습니다.")

    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v15"):
        st.session_state.pop('ship_db', None)
        st.session_state.pop('ship_updated', None)
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v15"):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
