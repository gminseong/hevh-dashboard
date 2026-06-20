"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v18.0 (확정판)
- Sheet1 우선 (header=2 = 3번째 행)
- 정확한 컬럼 매핑: model, code, ERP, SO, base stock, Plan.W25
- 음수 빨강 + 그레이 테마
- 누적실적 자동 매칭
"""
from datetime import datetime
import pandas as pd
import streamlit as st


def safe_classify(x):
    try:
        x = str(x).strip()
        if x.startswith('018'): return '3IN1'
        if x.startswith('01'): return 'PD'
        return 'OTHER'
    except Exception:
        return 'OTHER'


# ════════════════════════════════════════════════════════════
# Sheet1 로드 (header=2 고정, 2줄 헤더 처리)
# ════════════════════════════════════════════════════════════
def load_shipment(file):
    try:
        file.seek(0)
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names

        # Sheet1 우선
        target = None
        for s in sheet_names:
            if str(s).strip().lower() == 'sheet1':
                target = s
                break
        if target is None:
            for s in sheet_names:
                cl = str(s).strip().lower()
                if 'shipment' in cl and 'rev' in cl:
                    target = s
                    break
        if target is None:
            target = sheet_names[0]

        st.success(f"✅ 사용 시트: '{target}'")

        # 2줄 헤더 자동 병합 (행 1, 2 → 단일 헤더)
        file.seek(0)
        raw = pd.read_excel(file, sheet_name=target, header=None)
        
        # 헤더 위치 찾기: 'model'과 'ERP'가 같은 행에 있는지
        header_row = None
        for i in range(min(8, len(raw))):
            row_values = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
            if 'model' in row_values and 'erp' in row_values:
                header_row = i
                break
        
        if header_row is None:
            # Shipment Rev의 경우 cus 키워드로 시도
            for i in range(min(8, len(raw))):
                row_values = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
                if 'cus' in row_values:
                    header_row = i
                    break
        
        if header_row is None:
            st.error("❌ 헤더 행을 찾을 수 없습니다.")
            return pd.DataFrame()

        st.caption(f"🎯 헤더 행: {header_row}번째")

        # 헤더와 서브헤더 병합
        main_header = raw.iloc[header_row].values
        sub_header = raw.iloc[header_row + 1].values if header_row + 1 < len(raw) else [None] * len(main_header)
        
        merged_cols = []
        for mh, sh in zip(main_header, sub_header):
            mh_str = str(mh).strip() if pd.notna(mh) else ''
            sh_str = str(sh).strip() if pd.notna(sh) else ''
            
            if mh_str and sh_str and mh_str != 'nan' and sh_str != 'nan':
                # 둘 다 있으면 합치기 (예: 'Cut off' + '6/16' → 'Cut off.6/16')
                merged_cols.append(f"{mh_str}.{sh_str}")
            elif mh_str and mh_str != 'nan':
                merged_cols.append(mh_str)
            elif sh_str and sh_str != 'nan':
                merged_cols.append(sh_str)
            else:
                merged_cols.append(f"col_{len(merged_cols)}")

        # 데이터는 header_row + 2부터
        df = raw.iloc[header_row + 2:].copy()
        df.columns = merged_cols
        df = df.reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        
        st.caption(f"📋 컬럼 (앞 15개): {list(df.columns)[:15]}")

        # 컬럼명 정규화
        rename_map = {}
        for c in df.columns:
            cl = str(c).lower().strip()
            cl_clean = cl.replace(' ', '').replace('(', '').replace(')', '')
            
            if cl == 'model':
                rename_map[c] = 'model'
            elif cl == 'code':
                rename_map[c] = 'code'
            elif cl == 'erp' or '3in1code' in cl_clean:
                rename_map[c] = 'ERP'
            elif cl == 'so':
                rename_map[c] = 'SO'
            elif 'base stock' in cl or 'basestock' in cl_clean:
                rename_map[c] = 'base stock'
            elif cl in ('plan.w25', 'plan w25', 'planw25') or (cl.startswith('plan') and 'w' in cl and 'actual' not in cl):
                rename_map[c] = 'Plan.W25'
            elif cl == 'balance':
                rename_map[c] = 'balance'
            elif cl == 'note':
                rename_map[c] = 'Note'
            elif cl == 'cus':
                rename_map[c] = 'Cus'
            elif 'cut off cargo' in cl:
                rename_map[c] = 'Cut off Cargo'
            elif cl == 'hq' or 'hq request' in cl:
                rename_map[c] = 'HQ Request'
            elif 'po remain' in cl:
                rename_map[c] = 'PO Remain'
            elif 'ttl ship' in cl or 'ttlship' in cl_clean:
                rename_map[c] = 'TTL Ship'
            elif 'ttl plan' in cl or 'ttlplan' in cl_clean:
                rename_map[c] = 'TTL Plan'
            elif 'ttlstock' in cl_clean:
                rename_map[c] = 'TTLstock'
        df = df.rename(columns=rename_map)

        # ERP 없으면 패턴 매칭
        if 'ERP' not in df.columns:
            for col in df.columns:
                sample = df[col].dropna().astype(str).head(30).tolist()
                if len(sample) < 3:
                    continue
                erp_count = sum(1 for v in sample 
                                if (v.startswith('013') or v.startswith('018')) 
                                and len(v) >= 10)
                if erp_count >= 3:
                    df = df.rename(columns={col: 'ERP'})
                    break

        if 'ERP' not in df.columns:
            st.error(f"❌ ERP 컬럼 매핑 실패. 컬럼: {list(df.columns)}")
            return pd.DataFrame()

        # 유효 행
        df['ERP'] = df['ERP'].astype(str).str.strip()
        df = df[df['ERP'].str.startswith(('013', '018'))].copy()
        df = df.reset_index(drop=True)

        normalized = [c for c in ['model','code','ERP','SO','base stock','Plan.W25','balance','Note',
                                  'Cus','Cut off Cargo','HQ Request','PO Remain',
                                  'TTL Ship','TTL Plan','TTLstock']
                      if c in df.columns]
        st.success(f"📌 정규화 완료 ({len(normalized)}개): {normalized}")

        return df

    except Exception as e:
        st.error(f"❌ Excel 로드 실패: {e}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()


# ════════════════════════════════════════════════════════════
# 분석
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db):
    # 누적실적
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

    # 숫자 변환
    for col in ['SO', 'base stock', 'Plan.W25', 'balance']:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors='coerce').fillna(0).astype(int)
        else:
            m[col] = 0

    # TTLstock = base stock + 누적실적
    m['TTLstock'] = m['base stock'] + m['누적실적']
    
    # BAL_계획 = TTLstock + Plan.W25 - SO  (이번주 다 만들면)
    m['BAL_계획'] = m['TTLstock'] + m['Plan.W25'] - m['SO']
    # BAL_실적 = TTLstock - SO  (지금 당장)
    m['BAL_실적'] = m['TTLstock'] - m['SO']
    
    # 달성률
    rate = m['누적실적'] / m['Plan.W25'].where(m['Plan.W25'] != 0)
    rate = pd.to_numeric(rate, errors='coerce').replace([float('inf'), -float('inf')], 0).fillna(0)
    m['달성률(%)'] = (rate * 100).round(1)
    
    # 알람
    def get_alert(row):
        bal_plan = row['BAL_계획']
        bal_real = row['BAL_실적']
        if bal_plan < -5000:
            return "🔴 긴급"
        elif bal_plan < 0:
            return "🟠 부족"
        elif bal_real < 0:
            return "🟡 실적부족"
        else:
            return "✅ 정상"
    m['알람'] = m.apply(get_alert, axis=1)
    
    return m


# ════════════════════════════════════════════════════════════
# HTML 테이블
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
        if '긴급' in a: return '#FEE2E2'
        if '부족' in a and '실적' not in a: return '#FFEDD5'
        if '실적부족' in a: return '#FEF3C7'
        return '#FFFFFF'

    html = f'<div style="max-height:{height}px;overflow:auto;border:1px solid #E5E7EB;border-radius:6px;">'
    html += '<table style="width:100%;border-collapse:collapse;font-family:-apple-system,sans-serif;font-size:13px;">'
    html += '<thead style="position:sticky;top:0;z-index:10;"><tr style="background-color:#374151;color:white;">'
    for col in df.columns:
        html += f'<th style="padding:10px 8px;text-align:center;border:1px solid #4B5563;font-weight:700;white-space:nowrap;">{col}</th>'
    html += '</tr></thead><tbody>'
    
    for _, row in df.iterrows():
        bg = row_bg(row.get('알람', ''))
        html += f'<tr style="background-color:{bg};">'
        for col in df.columns:
            val = row[col]
            is_num = col in numeric_cols
            cell = format_cell(val, is_num)
            align = 'right' if is_num else 'center'
            html += f'<td style="padding:8px;text-align:{align};border-bottom:1px solid #E5E7EB;white-space:nowrap;">{cell}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    st.markdown(html, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 BAL_계획 = TTLstock + Plan - SO | BAL_실적 = TTLstock - SO")

    ship_db = st.session_state.get('ship_db', pd.DataFrame())
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_t = st.session_state.get('ship_updated', '-')
    prod_t = st.session_state.get('prod_updated', '-')

    s1, s2 = st.columns(2)
    s1.success(f"📁 출하계획: **{len(ship_db)}건** | {ship_t}") if not ship_db.empty else s1.warning("📁 출하계획: 없음")
    s2.success(f"📊 생산실적: **{len(prod_db)}건** | {prod_t}") if not prod_db.empty else s2.warning("📊 생산실적: 없음")

    st.markdown("##### 📤 파일 업로드")
    files = st.file_uploader(" ", type=["xlsx","csv"], accept_multiple_files=True,
                              key="ship_up_v18", label_visibility="collapsed")

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v18"):
            for f in files:
                fname = f.name.lower()
                if fname.endswith('.xlsx'):
                    df = load_shipment(f)
                    if not df.empty:
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

    numeric_cols = ['SO','base stock','누적실적','TTLstock','Plan.W25',
                    '달성률(%)','BAL_계획','BAL_실적','balance']

    # 긴급/부족
    urgent = m[m['알람'].isin(['🔴 긴급','🟠 부족','🟡 실적부족'])].sort_values('BAL_계획')
    if not urgent.empty:
        st.markdown("---")
        st.markdown(f"#### 🚨 즉시 조치 필요 ({len(urgent)}건)")
        show = [c for c in ['알람','model','code','ERP','MODEL_TYPE',
                            'SO','base stock','누적실적','TTLstock','Plan.W25',
                            '달성률(%)','BAL_계획','BAL_실적','Note'] if c in urgent.columns]
        render_html_table(urgent[show], numeric_cols, height=400)
    else:
        st.markdown("---")
        st.success("✅ 긴급/부족 모델 없음")

    # 전체 상세
    st.markdown("---")
    st.markdown("#### 📋 전체 상세")
    f1, f2 = st.columns(2)
    a_sel = f1.multiselect("알람", ['🔴 긴급','🟠 부족','🟡 실적부족','✅ 정상'], key="all_a_v18")
    t_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'], key="all_t_v18")
    
    v = m.copy()
    if a_sel: v = v[v['알람'].isin(a_sel)]
    if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
    
    all_cols = [c for c in ['알람','model','code','ERP','MODEL_TYPE',
                            'SO','base stock','누적실적','TTLstock','Plan.W25',
                            '달성률(%)','BAL_계획','BAL_실적','balance','Note'] if c in v.columns]
    
    if not v.empty:
        render_html_table(v[all_cols], numeric_cols, height=500)
        csv = v[all_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 CSV 다운로드", csv,
                           f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           "text/csv", key="dl_v18")

    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v18"):
        st.session_state.pop('ship_db', None)
        st.session_state.pop('ship_updated', None)
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v18"):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
