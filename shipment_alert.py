"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v14.0 (완전 안정판)
- 모든 시트 자동 탐색 (Sheet1, Shipment Rev 등)
- 모든 헤더 위치 자동 시도 (0~3행)
- ERP 패턴 자동 매핑 (Unnamed 컬럼도 대응)
- 모든 타입 안전 처리 (float, NaN 등)
- 그레이 테마 + 음수 빨강
- 긴급/부족 상단 강조 + 전체 상세
"""
from datetime import datetime
import pandas as pd
import streamlit as st


# ════════════════════════════════════════════════════════════
# 안전한 타입 분류
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
# 출하계획 로드 (모든 시트 + 헤더 자동 탐색)
# ════════════════════════════════════════════════════════════
def load_sheet1(file):
    """모든 시트/헤더를 탐색하여 ERP 컬럼이 있는 시트 찾기"""
    try:
        file.seek(0)
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        st.info(f"📋 발견된 시트: {sheet_names}")

        best_df = pd.DataFrame()
        best_sheet_info = None
        best_score = 0

        # 모든 시트 × 헤더 위치 탐색
        for sheet_name in sheet_names:
            for header_row in [0, 1, 2, 3]:
                try:
                    file.seek(0)
                    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
                    df.columns = [str(c).strip() for c in df.columns]

                    # 1) ERP 컬럼이 직접 있으면 최고 우선순위
                    if 'ERP' in df.columns:
                        sample = df['ERP'].dropna().astype(str).head(20).tolist()
                        erp_like = sum(1 for v in sample 
                                       if (v.startswith('013') or v.startswith('018')) 
                                       and len(v) >= 10)
                        if erp_like >= 3:
                            st.success(f"✅ 시트 '{sheet_name}' (header={header_row}): ERP 컬럼 직접 발견")
                            if 'model' in df.columns:
                                df = df[df['model'].notna()].copy()
                            df['ERP'] = df['ERP'].astype(str).str.strip()
                            return df

                    # 2) ERP 패턴으로 자동 탐지 (best score 갱신)
                    for col in df.columns:
                        sample = df[col].dropna().astype(str).head(30).tolist()
                        if len(sample) < 3:
                            continue
                        erp_like = sum(1 for v in sample 
                                       if (str(v).startswith('013') or str(v).startswith('018')) 
                                       and len(str(v)) >= 10)
                        if erp_like > best_score:
                            best_score = erp_like
                            best_sheet_info = (sheet_name, header_row, col)
                            best_df = df.copy()
                except Exception:
                    continue

        # 자동 매핑 결과
        if best_score >= 3 and best_sheet_info:
            sheet_name, header_row, erp_col = best_sheet_info
            st.success(f"✅ 자동 매핑: 시트 '{sheet_name}' (header={header_row}), '{erp_col}' → 'ERP'")
            best_df = best_df.rename(columns={erp_col: 'ERP'})
            best_df['ERP'] = best_df['ERP'].astype(str).str.strip()

            # model 컬럼 자동 인식
            for c in list(best_df.columns):
                cl = str(c).lower().strip()
                if cl == 'model':
                    if c != 'model':
                        best_df = best_df.rename(columns={c: 'model'})
                    break

            # 빈 행 제거
            if 'model' in best_df.columns:
                best_df = best_df[best_df['model'].notna()].copy()
            else:
                best_df = best_df[best_df['ERP'].notna()].copy()

            # ERP가 유효한 값만 (013 또는 018 시작)
            best_df = best_df[
                best_df['ERP'].astype(str).str.startswith(('013', '018'))
            ].copy()

            return best_df

        st.error("❌ 모든 시트에서 ERP 컬럼을 찾지 못했습니다.")
        return pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Excel 로드 실패: {e}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()


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

    # 머지
    m = ship_db.copy()
    m['ERP'] = m['ERP'].astype(str).str.strip()
    m['MODEL_TYPE'] = m['ERP'].apply(safe_classify)
    m['누적실적'] = m['ERP'].map(actual_dict).fillna(0).astype(int)

    # 기초재고 컬럼 찾기
    stock_col = None
    for c in m.columns:
        cl = str(c).lower()
        if 'base stock' in cl or 'ttlstock' in cl or 'stock' in cl:
            stock_col = c
            break
    if stock_col:
        m['기초재고'] = pd.to_numeric(m[stock_col], errors='coerce').fillna(0).astype(int)
    else:
        m['기초재고'] = 0

    # TTLstock = 기초재고 + 누적실적
    m['TTLstock'] = m['기초재고'] + m['누적실적']

    # SO 찾기
    so_col = None
    for c in m.columns:
        cl = str(c).lower().strip()
        if cl == 'so' or 'ttl ship' in cl:
            so_col = c
            break
    if so_col:
        m['SO'] = pd.to_numeric(m[so_col], errors='coerce').fillna(0).astype(int)
    else:
        m['SO'] = 0

    # Plan 찾기
    plan_col = None
    for c in m.columns:
        cl = str(c).lower()
        if 'plan' in cl and 'w' in cl:
            plan_col = c
            break
    if plan_col is None:
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
        r = row['달성률(%)']
        if bal < -5000:
            return "🔴 긴급"
        elif bal < 0:
            return "🟠 부족"
        elif r < 70 and row['Plan.W25'] > 0:
            return "🟡 지연"
        else:
            return "✅ 정상"
    m['알람'] = m.apply(get_alert, axis=1)

    # code 컬럼 자동 인식
    if 'code' not in m.columns:
        for c in m.columns:
            cl = str(c).lower()
            if 'bn' in cl and 'code' in cl:
                m = m.rename(columns={c: 'code'})
                break

    # Note 컬럼 자동 인식
    if 'Note' not in m.columns:
        for c in m.columns:
            if str(c).lower() == 'note':
                m = m.rename(columns={c: 'Note'})
                break

    return m


# ════════════════════════════════════════════════════════════
# 스타일 (그레이 테마 + 음수 빨강)
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
# 메인 렌더링
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
        key="ship_up_v14", label_visibility="collapsed"
    )

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v14"):
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
            # st.rerun() 제거 — 메시지 사라지지 않게

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

    # ═══════════════════════════════════════════════════
    # KPI 카드
    # ═══════════════════════════════════════════════════
    st.markdown("### 🚨 분석 결과")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(m)}건")
    k2.metric("🔴 긴급", int((m['알람']=='🔴 긴급').sum()))
    k3.metric("🟠 부족", int((m['알람']=='🟠 부족').sum()))
    k4.metric("🟡 지연", int((m['알람']=='🟡 지연').sum()))
    k5.metric("✅ 정상", int((m['알람']=='✅ 정상').sum()))

    numeric_cols = ['SO', '기초재고', '누적실적', 'TTLstock', 'Plan.W25',
                    '달성률(%)', 'NEW_BALANCE']

    # ═══════════════════════════════════════════════════
    # 긴급 & 부족 모델 (상단 강조)
    # ═══════════════════════════════════════════════════
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

    # ═══════════════════════════════════════════════════
    # 전체 상세 (필터 + 표)
    # ═══════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("#### 📋 전체 상세")

    f1, f2 = st.columns(2)
    a_sel = f1.multiselect("알람 등급",
                           ['🔴 긴급','🟠 부족','🟡 지연','✅ 정상'],
                           key="all_a_v14")
    t_sel = f2.multiselect("모델 타입", ['PD','3IN1','OTHER'], key="all_t_v14")

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
            "text/csv",
            key="dl_v14"
        )
    else:
        st.info("선택한 필터에 해당하는 데이터가 없습니다.")

    # 초기화
    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v14"):
        st.session_state.pop('ship_db', None)
        st.session_state.pop('ship_updated', None)
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v14"):
        st.session_state.pop('prod_db', None)
        st.session_state.pop('prod_updated', None)
        st.rerun()
