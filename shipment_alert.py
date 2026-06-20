"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v10.0
- Sheet1 우선 → Shipment Rev 자동 폴백
- 어떤 시트 구조든 자동 처리
- style 제거로 KeyError 차단
"""
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px


# ════════════════════════════════════════════════════════════
# 시트 자동 탐지 및 로드
# ════════════════════════════════════════════════════════════
def load_shipment_excel(file):
    """
    출하계획 엑셀 파일을 자동으로 로드.
    Sheet1 우선, 없으면 Shipment Rev, 마지막엔 첫 시트.
    
    반환: (DataFrame, 사용된 시트명, 데이터 소스 유형)
    """
    file.seek(0)
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    
    # ───── 우선순위 1: Sheet1 (깔끔한 구조) ─────
    if 'Sheet1' in sheet_names:
        file.seek(0)
        df = pd.read_excel(file, sheet_name='Sheet1')
        # 컬럼 검증: ERP, SO가 있어야 함
        if 'ERP' in df.columns and 'SO' in df.columns:
            if 'model' in df.columns:
                df = df[df['model'].notna()].copy()
            return df, 'Sheet1', 'clean'
    
    # ───── 우선순위 2: Shipment Rev (헤더 2줄 처리 필요) ─────
    if 'Shipment Rev' in sheet_names:
        file.seek(0)
        raw = pd.read_excel(file, sheet_name='Shipment Rev', header=None)
        
        # 헤더 자동 탐지: ERP/3in1Code/Cus 가 있는 행 찾기
        header_row = None
        for i in range(min(5, len(raw))):
            row_str = ' '.join(str(v) for v in raw.iloc[i].values if pd.notna(v))
            if 'Cus' in row_str or 'ERP' in row_str or '3in1Code' in row_str:
                header_row = i
                break
        
        if header_row is not None:
            file.seek(0)
            df = pd.read_excel(file, sheet_name='Shipment Rev', header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Shipment Rev 컬럼명 → Sheet1 스타일로 정규화
            df = normalize_shipment_rev_columns(df)
            return df, 'Shipment Rev', 'rev'
    
    # ───── 우선순위 3: 첫 시트 (마지막 수단) ─────
    file.seek(0)
    df = pd.read_excel(file, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df, sheet_names[0], 'fallback'


def normalize_shipment_rev_columns(df):
    """
    Shipment Rev의 모호한 컬럼명을 표준 컬럼명으로 매핑.
    """
    # 컬럼 패턴 추정 (대소문자 무시)
    rename_map = {}
    cols_list = list(df.columns)
    
    for col in cols_list:
        col_lower = str(col).lower().strip()
        # 패턴 매칭
        if col_lower in ('cus', 'customer'):
            rename_map[col] = 'Cus'
        elif 'cut off cargo' in col_lower:
            rename_map[col] = 'Cut off Cargo'
        elif col_lower in ('hq', 'hq request'):
            rename_map[col] = 'HQ Request'
        elif col_lower == 'model':
            rename_map[col] = 'model'
        elif col_lower == 'code':
            rename_map[col] = 'code'
        elif col_lower in ('erp', '3in1code', '3in1code(fg)'):
            rename_map[col] = 'ERP'
        elif col_lower == 'note':
            rename_map[col] = 'Note'
    
    df = df.rename(columns=rename_map)
    
    # ERP 컬럼이 없으면 패턴으로 찾기 (0137xxx, 0187xxx 형태)
    if 'ERP' not in df.columns:
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(20).tolist()
            if sample:
                erp_like = sum(1 for v in sample 
                               if (v.startswith('013') or v.startswith('018')) 
                               and len(v) >= 10)
                if erp_like >= len(sample) * 0.5:  # 절반 이상이 ERP 패턴
                    df = df.rename(columns={col: 'ERP'})
                    break
    
    # SO 컬럼 추정 (model 다음 다음 다음 정도의 숫자 컬럼)
    if 'SO' not in df.columns and 'ERP' in df.columns:
        erp_idx = list(df.columns).index('ERP')
        # ERP 바로 뒤의 첫 번째 숫자 컬럼을 SO로 추정
        for col in df.columns[erp_idx+1:]:
            sample = pd.to_numeric(df[col], errors='coerce').dropna()
            if len(sample) > 5 and sample.max() > 100:
                df = df.rename(columns={col: 'SO'})
                break
    
    # 빈 행 제거
    if 'model' in df.columns:
        df = df[df['model'].notna()].copy()
    elif 'ERP' in df.columns:
        df = df[df['ERP'].notna()].copy()
    
    return df


# ════════════════════════════════════════════════════════════
# 메인 렌더링
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD: P-ATE 기준 | 3in1: ASSY 기준 | 자동 시트 탐지")

    # ───── 데이터 현황 ─────
    ship_db = st.session_state.get('ship_db', pd.DataFrame())
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_t = st.session_state.get('ship_updated', '-')
    prod_t = st.session_state.get('prod_updated', '-')
    ship_src = st.session_state.get('ship_source', '-')

    s1, s2 = st.columns(2)
    if not ship_db.empty:
        s1.success(f"📁 출하계획: **{len(ship_db)}건** | {ship_t} | 시트: {ship_src}")
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
                        df, sheet_used, src_type = load_shipment_excel(f)
                        
                        # 필수 컬럼 검증
                        if 'ERP' not in df.columns:
                            st.error(f"❌ '{sheet_used}' 시트에서 ERP 컬럼을 찾지 못했습니다.")
                            st.caption(f"보유 컬럼: {list(df.columns)[:10]}...")
                            continue
                        if 'SO' not in df.columns:
                            st.warning(f"⚠️ SO 컬럼 없음. 기본값 0으로 처리됩니다.")
                            df['SO'] = 0
                        
                        st.session_state['ship_db'] = df
                        st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        st.session_state['ship_source'] = f"{sheet_used}"
                        st.success(f"✅ 출하계획 저장: {sheet_used} 시트 ({len(df)}건)")
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

    if ship_db.empty or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    # ───── 분석 ─────
    try:
        # 1. 실적 집계
        p = prod_db.copy()
        p['ERP_STR'] = p['FINAL_MAT_ID'].astype(str).str.strip()
        p['TYPE'] = p['ERP_STR'].apply(
            lambda x: '3IN1' if x.startswith('018') else ('PD' if x.startswith('01') else 'OTHER')
        )
        
        pd_a = p[(p['TYPE']=='PD') & (p['OPER_DESC']=='P-ATE')].groupby('ERP_STR')['QTY'].sum()
        in1_a = p[(p['TYPE']=='3IN1') & (p['OPER_DESC']=='ASSY')].groupby('ERP_STR')['QTY'].sum()
        actual = pd.concat([pd_a, in1_a]).reset_index()
        actual.columns = ['ERP', '누적실적']

        # 2. 머지
        m = ship_db.copy()
        m['ERP'] = m['ERP'].astype(str).str.strip()
        actual['ERP'] = actual['ERP'].astype(str).str.strip()
        m = m.merge(actual, on='ERP', how='left')
        m['누적실적'] = pd.to_numeric(m['누적실적'], errors='coerce').fillna(0)
        m['MODEL_TYPE'] = m['ERP'].apply(
            lambda x: '3IN1' if str(x).startswith('018') else ('PD' if str(x).startswith('01') else 'OTHER')
        )

        # 3. 숫자 변환
        m['SO_숫자'] = pd.to_numeric(m.get('SO', 0), errors='coerce').fillna(0)
        
        # base stock 찾기
        stock_col = None
        for c in m.columns:
            if 'base stock' in str(c).lower() or 'stock' in str(c).lower():
                stock_col = c
                break
        if stock_col:
            m['기초재고'] = pd.to_numeric(m[stock_col], errors='coerce').fillna(0)
        else:
            m['기초재고'] = 0

        # 4. NEW_GAP
        m['NEW_GAP'] = m['기초재고'] + m['누적실적'] - m['SO_숫자']

        # 5. 달성률
        plan_col = None
        for c in m.columns:
            if 'plan' in str(c).lower() and 'w' in str(c).lower():
                plan_col = c
                break
        if plan_col is None:
            for c in m.columns:
                if 'plan' in str(c).lower():
                    plan_col = c
                    break
        
        if plan_col:
            plan = pd.to_numeric(m[plan_col], errors='coerce')
            rate = m['누적실적'] / plan.where(plan != 0)
            rate = pd.to_numeric(rate, errors='coerce').replace([float('inf'), -float('inf')], 0).fillna(0)
            m['달성률(%)'] = (rate * 100).round(1)
        else:
            m['달성률(%)'] = 0.0

        # 6. 알람 등급
        def get_alert(row):
            gap = row['NEW_GAP']
            rate = row['달성률(%)']
            if gap < -5000:
                return "🔴 긴급"
            elif gap < 0:
                return "🟠 부족"
            elif rate < 70:
                return "🟡 지연"
            else:
                return "✅ 정상"
        m['알람'] = m.apply(get_alert, axis=1)

    except Exception as e:
        st.error(f"❌ 분석 오류: {e}")
        return

    # ───── KPI ─────
    st.markdown("### 🚨 분석 결과")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("전체", f"{len(m)}건")
    k2.metric("🔴 긴급", int((m['알람']=='🔴 긴급').sum()))
    k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
    k4.metric("🟡 지연", int((m['알람']=='🟡 지연').sum()))
    k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    # ───── 필터 ─────
    f1, f2, f3 = st.columns(3)
    alert_sel = f1.multiselect("알람 등급", ['🔴 긴급','🟠 부족','🟡 지연','✅ 정상'])
    type_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'])
    cus_options = sorted(m['Cus'].dropna().unique()) if 'Cus' in m.columns else []
    cus_sel = f3.multiselect("거래선", cus_options) if cus_options else []
    
    v = m.copy()
    if alert_sel:
        v = v[v['알람'].isin(alert_sel)]
    if type_sel:
        v = v[v['MODEL_TYPE'].isin(type_sel)]
    if cus_sel:
        v = v[v['Cus'].isin(cus_sel)]

    # ───── 테이블 ─────
    st.markdown("##### 📋 알람 상세")
    show_cols = []
    for c in ['알람', 'Cus', 'model', 'code', 'ERP', 'MODEL_TYPE',
              'Cut off Cargo', 'HQ Request',
              'SO', '기초재고', '누적실적', '달성률(%)', 'NEW_GAP', 'Note']:
        if c in v.columns:
            show_cols.append(c)
    
    if not v.empty and show_cols:
        st.dataframe(v[show_cols], use_container_width=True, height=450)
    else:
        st.info("표시할 데이터가 없습니다.")

    # ───── 차트 ─────
    st.markdown("##### 📈 알람 분포")
    if not v.empty:
        cnt = v['알람'].value_counts().reset_index()
        cnt.columns = ['알람', '건수']
        cmap = {'🔴 긴급':'#ef4444','🟠 부족':'#f97316','🟡 지연':'#eab308','✅ 정상':'#22c55e'}
        fig = px.bar(cnt, x='알람', y='건수', color='알람',
                     color_discrete_map=cmap, height=300)
        fig.update_layout(showlegend=False, margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fig, use_container_width=True)

    # ───── 초기화 ─────
    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True):
        st.session_state.pop('ship_db', None)
        st.session_state.pop('ship_updated', None)
        st.session_state.pop('ship_source', None)
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
