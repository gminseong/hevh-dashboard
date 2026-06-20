"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v11.0 (완성판)
- Sheet1 + Shipment Rev 자동 통합
- 출하 우선순위 시뮬레이션 (Cut off FIFO)
- 그레이 테마 + 음수 빨간색 강조
"""
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px


# ════════════════════════════════════════════════════════════
# 시트 자동 로드
# ════════════════════════════════════════════════════════════
def load_sheet1(file):
    """Sheet1 로드 (ERP, SO, base stock, Plan.W25 등 포함)"""
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


def load_shipment_rev(file):
    """Shipment Rev 로드 (Cus, Cut off Cargo, PO Remain 등)"""
    try:
        file.seek(0)
        raw = pd.read_excel(file, sheet_name='Shipment Rev', header=None)
        
        # 헤더 행 찾기
        header_row = None
        for i in range(min(5, len(raw))):
            row_str = ' '.join(str(v) for v in raw.iloc[i].values if pd.notna(v))
            if 'Cus' in row_str:
                header_row = i
                break
        if header_row is None:
            return pd.DataFrame()
        
        file.seek(0)
        df = pd.read_excel(file, sheet_name='Shipment Rev', header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 컬럼 자동 매핑
        rename_map = {}
        for c in df.columns:
            cl = str(c).lower()
            if cl == 'cus' or 'customer' in cl:
                rename_map[c] = 'Cus'
            elif 'cut off cargo' in cl:
                rename_map[c] = 'Cut off Cargo'
            elif cl in ('hq', 'hq request'):
                rename_map[c] = 'HQ Request'
            elif cl == 'model':
                rename_map[c] = 'model'
            elif 'bn' in cl and 'code' in cl:
                rename_map[c] = 'code'
            elif '3in1code' in cl or cl == 'erp':
                rename_map[c] = 'ERP'
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
        df = df.rename(columns=rename_map)
        
        # ERP가 자동 매핑 안 됐으면 패턴으로 찾기
        if 'ERP' not in df.columns:
            for c in df.columns:
                sample = df[c].dropna().astype(str).head(20).tolist()
                if sample:
                    erp_like = sum(1 for v in sample 
                                   if (v.startswith('013') or v.startswith('018')) 
                                   and len(v) >= 10)
                    if erp_like >= len(sample) * 0.5:
                        df = df.rename(columns={c: 'ERP'})
                        break
        
        # 필수: Cus + ERP
        if 'Cus' not in df.columns or 'ERP' not in df.columns:
            return pd.DataFrame()
        
        # 빈 행 제거
        df = df[df['ERP'].notna() & df['Cus'].notna()].copy()
        df['ERP'] = df['ERP'].astype(str).str.strip()
        df['Cus'] = df['Cus'].astype(str).str.strip()
        
        # Cut off 날짜 정규화
        if 'Cut off Cargo' in df.columns:
            df['Cut_off_date'] = pd.to_datetime(df['Cut off Cargo'], errors='coerce')
        
        return df
    except Exception:
        return pd.DataFrame()


def load_shipment_excel(file):
    """출하계획 엑셀에서 Sheet1 + Shipment Rev 통합 로드"""
    sheet1 = load_sheet1(file)
    ship_rev = load_shipment_rev(file)
    return sheet1, ship_rev


# ════════════════════════════════════════════════════════════
# 출하 우선순위 시뮬레이션
# ════════════════════════════════════════════════════════════
def simulate_allocation(ship_rev, erp_stock_dict):
    """
    Cut off 빠른 순으로 차감 시뮬레이션
    
    erp_stock_dict: {erp: 가용재고(TTLstock+Plan+누적실적)}
    """
    if ship_rev.empty:
        return ship_rev
    
    df = ship_rev.copy()
    
    # PO Remain을 숫자로
    if 'PO Remain' in df.columns:
        df['필요량'] = pd.to_numeric(df['PO Remain'], errors='coerce').fillna(0)
    elif 'TTL Ship' in df.columns:
        df['필요량'] = pd.to_numeric(df['TTL Ship'], errors='coerce').fillna(0)
    else:
        df['필요량'] = 0
    
    # 정렬: ERP → Cut off 빠른순
    if 'Cut_off_date' in df.columns:
        df = df.sort_values(['ERP', 'Cut_off_date']).reset_index(drop=True)
    else:
        df = df.sort_values('ERP').reset_index(drop=True)
    
    # 시뮬레이션
    df['할당가능'] = 0
    df['부족수량'] = 0
    df['잔여재고'] = 0
    df['알람'] = ''
    
    current_erp = None
    available = 0
    
    for idx, row in df.iterrows():
        erp = row['ERP']
        need = row['필요량']
        
        # ERP 바뀌면 재고 초기화
        if erp != current_erp:
            current_erp = erp
            available = erp_stock_dict.get(erp, 0)
        
        # 차감 시뮬레이션
        if available >= need:
            df.at[idx, '할당가능'] = need
            df.at[idx, '부족수량'] = 0
            available -= need
            df.at[idx, '잔여재고'] = available
            df.at[idx, '알람'] = '✅ 정상'
        elif available > 0:
            df.at[idx, '할당가능'] = available
            df.at[idx, '부족수량'] = need - available
            available = 0
            df.at[idx, '잔여재고'] = 0
            df.at[idx, '알람'] = '🟠 일부부족'
        else:
            df.at[idx, '할당가능'] = 0
            df.at[idx, '부족수량'] = need
            df.at[idx, '잔여재고'] = 0
            df.at[idx, '알람'] = '🔴 전량부족'
    
    return df


# ════════════════════════════════════════════════════════════
# 스타일 적용 (그레이 테마 + 음수 강조)
# ════════════════════════════════════════════════════════════
def style_alert_df(df):
    """그레이 테마 + 음수 빨간색 + 알람 행 배경"""
    def highlight_negative(val):
        """음수면 빨간색 굵게"""
        try:
            v = float(val)
            if v < 0:
                return 'color: #DC2626; font-weight: bold;'
            elif v > 0 and v >= 1000:
                return 'color: #059669;'
        except Exception:
            pass
        return ''
    
    def highlight_alert_row(row):
        """알람 등급별 행 배경"""
        if '알람' not in row.index:
            return [''] * len(row)
        alert = str(row['알람'])
        if '긴급' in alert or '전량부족' in alert:
            bg = '#FEE2E2'
        elif '부족' in alert:
            bg = '#FFEDD5'
        elif '지연' in alert:
            bg = '#FEF3C7'
        elif '정상' in alert:
            bg = '#F0FDF4'
        else:
            bg = ''
        return [f'background-color: {bg};'] * len(row)
    
    # 숫자 컬럼만 음수 강조 대상
    num_cols = [c for c in df.columns 
                if c in ['SO', 'TTLstock', 'Plan', 'NEW_BALANCE', 
                        '필요량', '할당가능', '부족수량', '잔여재고',
                        'PO Remain', 'TTL Ship', 'TTL Plan',
                        '기초재고', '누적실적', 'BALANCE']]
    
    try:
        styled = df.style.apply(highlight_alert_row, axis=1)
        if num_cols:
            styled = styled.applymap(highlight_negative, subset=num_cols)
        # 헤더 스타일
        styled = styled.set_table_styles([
            {'selector': 'thead th',
             'props': [('background-color', '#374151'),
                       ('color', 'white'),
                       ('font-weight', 'bold'),
                       ('text-align', 'center'),
                       ('padding', '8px')]},
            {'selector': 'tbody td',
             'props': [('padding', '6px 10px'),
                       ('border-bottom', '1px solid #E5E7EB')]},
        ])
        return styled
    except Exception:
        return df


# ════════════════════════════════════════════════════════════
# 메인 렌더링
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD: P-ATE 기준 | 3in1: ASSY 기준 | 출하 우선순위 시뮬레이션")

    # ───── 데이터 현황 ─────
    sheet1_db = st.session_state.get('sheet1_db', pd.DataFrame())
    shiprev_db = st.session_state.get('shiprev_db', pd.DataFrame())
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_t = st.session_state.get('ship_updated', '-')
    prod_t = st.session_state.get('prod_updated', '-')

    s1, s2 = st.columns(2)
    if not sheet1_db.empty or not shiprev_db.empty:
        srcs = []
        if not sheet1_db.empty: srcs.append(f"Sheet1 {len(sheet1_db)}건")
        if not shiprev_db.empty: srcs.append(f"ShipRev {len(shiprev_db)}건")
        s1.success(f"📁 출하계획: {' / '.join(srcs)} | {ship_t}")
    else:
        s1.warning("📁 출하계획: 데이터 없음")
    if not prod_db.empty:
        s2.success(f"📊 생산실적: **{len(prod_db)}건** | {prod_t}")
    else:
        s2.warning("📊 생산실적: 데이터 없음")

    # ───── 파일 업로드 ─────
    st.markdown("##### 📤 파일 업로드 (.xlsx + .csv 동시 가능)")
    files = st.file_uploader(
        " ", type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="ship_up", label_visibility="collapsed"
    )

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True):
            for f in files:
                if f.name.lower().endswith('.xlsx'):
                    try:
                        sheet1, shiprev = load_shipment_excel(f)
                        if not sheet1.empty:
                            st.session_state['sheet1_db'] = sheet1
                            st.success(f"✅ Sheet1 저장 ({len(sheet1)}건)")
                        if not shiprev.empty:
                            st.session_state['shiprev_db'] = shiprev
                            st.success(f"✅ Shipment Rev 저장 ({len(shiprev)}건)")
                        if sheet1.empty and shiprev.empty:
                            st.error("❌ Sheet1, Shipment Rev 모두 인식 실패")
                        st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                    except Exception as e:
                        st.error(f"❌ 출하계획 오류: {e}")
                elif f.name.lower().endswith('.csv'):
                    try:
                        df = pd.read_csv(f)
                        st.session_state['prod_db'] = df
                        st.session_state['prod_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        st.success(f"✅ 생산실적 저장 ({len(df)}건)")
                    except Exception as e:
                        st.error(f"❌ 생산실적 오류: {e}")
            st.rerun()

    st.markdown("---")

    # 필수: 출하계획 + 생산실적
    if (sheet1_db.empty and shiprev_db.empty) or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    # ───── 누적실적 집계 ─────
    try:
        p = prod_db.copy()
        p['ERP_STR'] = p['FINAL_MAT_ID'].astype(str).str.strip()
        p['TYPE'] = p['ERP_STR'].apply(
            lambda x: '3IN1' if x.startswith('018') else ('PD' if x.startswith('01') else 'OTHER')
        )
        pd_a = p[(p['TYPE']=='PD') & (p['OPER_DESC']=='P-ATE')].groupby('ERP_STR')['QTY'].sum()
        in1_a = p[(p['TYPE']=='3IN1') & (p['OPER_DESC']=='ASSY')].groupby('ERP_STR')['QTY'].sum()
        actual_dict = pd.concat([pd_a, in1_a]).to_dict()
    except Exception as e:
        st.error(f"❌ 실적 집계 오류: {e}")
        return

    # ───── ERP별 가용재고 계산 (Sheet1 기준) ─────
    erp_stock_dict = {}
    if not sheet1_db.empty:
        for _, row in sheet1_db.iterrows():
            erp = str(row.get('ERP', '')).strip()
            if not erp:
                continue
            # base stock
            stock = 0
            for c in sheet1_db.columns:
                if 'base stock' in str(c).lower() or 'stock' in str(c).lower():
                    stock = pd.to_numeric(row.get(c, 0), errors='coerce')
                    stock = stock if pd.notna(stock) else 0
                    break
            # Plan.W25
            plan = pd.to_numeric(row.get('Plan.W25', 0), errors='coerce')
            plan = plan if pd.notna(plan) else 0
            # 누적실적
            actual = actual_dict.get(erp, 0)
            erp_stock_dict[erp] = float(stock) + float(plan) + float(actual)

    # ═════════════════════════════════════════════════════════
    # 분석 결과 표시
    # ═════════════════════════════════════════════════════════
    st.markdown("### 🚨 분석 결과")
    
    tab1, tab2 = st.tabs(["📋 ERP별 (마스터)", "🚛 Cus별 (출하 시뮬레이션)"])
    
    # ───── 탭1: ERP별 마스터 ─────
    with tab1:
        if sheet1_db.empty:
            st.warning("Sheet1 데이터가 없어 ERP별 분석을 할 수 없습니다.")
        else:
            m = sheet1_db.copy()
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
            m['기초재고'] = pd.to_numeric(m[stock_col], errors='coerce').fillna(0).astype(int) if stock_col else 0
            
            m['TTLstock'] = m['기초재고'] + m['누적실적']
            m['SO_num'] = pd.to_numeric(m.get('SO', 0), errors='coerce').fillna(0).astype(int)
            m['Plan_num'] = pd.to_numeric(m.get('Plan.W25', 0), errors='coerce').fillna(0).astype(int)
            m['NEW_BALANCE'] = m['TTLstock'] + m['Plan_num'] - m['SO_num']
            
            # 달성률
            rate = m['누적실적'] / m['Plan_num'].where(m['Plan_num'] != 0)
            rate = pd.to_numeric(rate, errors='coerce').replace([float('inf'), -float('inf')], 0).fillna(0)
            m['달성률(%)'] = (rate * 100).round(1)
            
            # 알람
            def alert_master(row):
                bal = row['NEW_BALANCE']
                if bal < -5000:
                    return "🔴 긴급"
                elif bal < 0:
                    return "🟠 부족"
                elif row['달성률(%)'] < 70:
                    return "🟡 지연"
                else:
                    return "✅ 정상"
            m['알람'] = m.apply(alert_master, axis=1)
            
            # KPI
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("전체", f"{len(m)}건")
            k2.metric("🔴 긴급", int((m['알람']=='🔴 긴급').sum()))
            k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
            k4.metric("🟡 지연", int((m['알람']=='🟡 지연').sum()))
            k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))
            
            # 필터
            f1, f2 = st.columns(2)
            a_sel = f1.multiselect("알람", ['🔴 긴급','🟠 부족','🟡 지연','✅ 정상'], key="m_a")
            t_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'], key="m_t")
            v = m.copy()
            if a_sel: v = v[v['알람'].isin(a_sel)]
            if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
            
            # 표시
            show = ['알람','model','code','ERP','MODEL_TYPE',
                    'SO_num','기초재고','누적실적','TTLstock','Plan_num',
                    '달성률(%)','NEW_BALANCE','Note']
            show = [c for c in show if c in v.columns]
            if not v.empty:
                disp = v[show].rename(columns={'SO_num':'SO','Plan_num':'Plan.W25'})
                try:
                    st.dataframe(style_alert_df(disp), use_container_width=True, height=500)
                except Exception:
                    st.dataframe(disp, use_container_width=True, height=500)
    
    # ───── 탭2: Cus별 시뮬레이션 ─────
    with tab2:
        if shiprev_db.empty:
            st.warning("Shipment Rev 데이터가 없어 거래선별 분석을 할 수 없습니다.")
        else:
            sim = simulate_allocation(shiprev_db, erp_stock_dict)
            
            # KPI
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("전체 주문", f"{len(sim)}건")
            k2.metric("🔴 전량부족", int((sim['알람']=='🔴 전량부족').sum()))
            k3.metric("🟠 일부부족", int((sim['알람']=='🟠 일부부족').sum()))
            k4.metric("✅ 정상", int((sim['알람']=='✅ 정상').sum()))
            
            # 필터
            f1, f2 = st.columns(2)
            a_sel2 = f1.multiselect("알람", ['🔴 전량부족','🟠 일부부족','✅ 정상'], key="s_a")
            cus_opt = sorted(sim['Cus'].dropna().unique()) if 'Cus' in sim.columns else []
            c_sel = f2.multiselect("거래선(Cus)", cus_opt, key="s_c")
            v = sim.copy()
            if a_sel2: v = v[v['알람'].isin(a_sel2)]
            if c_sel: v = v[v['Cus'].isin(c_sel)]
            
            # 표시
            show = ['알람','Cus','Cut off Cargo','HQ Request',
                    'model','code','ERP','필요량','할당가능','부족수량','잔여재고','Note']
            show = [c for c in show if c in v.columns]
            if not v.empty:
                try:
                    st.dataframe(style_alert_df(v[show]), use_container_width=True, height=500)
                except Exception:
                    st.dataframe(v[show], use_container_width=True, height=500)
            
            # 거래선별 차트
            st.markdown("##### 📊 거래선별 부족 현황")
            if not v.empty and 'Cus' in v.columns:
                cus_sum = v[v['부족수량'] > 0].groupby('Cus')['부족수량'].sum().reset_index()
                if not cus_sum.empty:
                    fig = px.bar(cus_sum.sort_values('부족수량', ascending=True),
                                 x='부족수량', y='Cus', orientation='h',
                                 color='부족수량', color_continuous_scale='Reds',
                                 height=300)
                    fig.update_layout(margin=dict(l=0,r=0,t=10,b=0))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("✅ 부족 항목 없음")

    # ───── 초기화 ─────
    st.markdown("---")
    r1, r2, r3 = st.columns(3)
    if r1.button("🗑️ Sheet1 초기화", use_container_width=True):
        st.session_state.pop('sheet1_db', None)
        st.session_state.pop('ship_updated', None)
        st.rerun()
    if r2.button("🗑️ ShipRev 초기화", use_container_width=True):
        st.session_state.pop('shiprev_db', None)
        st.rerun()
    if r3.button("🗑️ 생산실적 초기화", use_container_width=True):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
