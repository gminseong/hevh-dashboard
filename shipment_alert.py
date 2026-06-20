"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 모듈 v7.0 (안정판)
- Sheet1 헤더 2줄 자동 병합
- 통합 파일 업로드 (자동 인식)
- session_state 기반 안전 저장
- PD 모델: P-ATE 완성 기준
- 3in1 모델: ASSY 완성 기준
"""
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px


# ════════════════════════════════════════════════════════════
# 모델 타입 판별
# ════════════════════════════════════════════════════════════
def classify_model_type(code):
    if pd.isna(code):
        return "UNKNOWN"
    code = str(code).strip()
    if code.startswith("018"):
        return "3IN1"
    elif code.startswith("01"):
        return "PD"
    return "OTHER"


# ════════════════════════════════════════════════════════════
# 유연 컬럼 매칭
# ════════════════════════════════════════════════════════════
def find_column(df, candidates):
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    for cand in candidates:
        key = cand.strip().lower()
        for col_key, col_orig in cols_lower.items():
            if key in col_key:
                return col_orig
    return None


# ════════════════════════════════════════════════════════════
# Sheet1 헤더 2줄 자동 병합
# ════════════════════════════════════════════════════════════
def smart_load_sheet1(file):
    """Sheet1의 2줄 헤더를 병합하여 깨끗한 DataFrame 반환"""
    file.seek(0)
    raw = pd.read_excel(file, sheet_name="Sheet1", header=None)

    # 0행과 1행을 합쳐서 헤더 생성
    new_headers = []
    for col_idx in range(len(raw.columns)):
        h0 = raw.iloc[0, col_idx] if len(raw) > 0 else None
        h1 = raw.iloc[1, col_idx] if len(raw) > 1 else None

        h0_str = str(h0).strip() if pd.notna(h0) else ''
        h1_str = str(h1).strip() if pd.notna(h1) else ''
        if h0_str in ['nan', 'None']:
            h0_str = ''
        if h1_str in ['nan', 'None']:
            h1_str = ''

        if h0_str and h1_str:
            new_headers.append(f"{h0_str}.{h1_str}")
        elif h0_str:
            new_headers.append(h0_str)
        elif h1_str:
            new_headers.append(h1_str)
        else:
            new_headers.append(f"col_{col_idx}")

    # 2행부터 실제 데이터
    df = raw.iloc[2:].copy()
    df.columns = new_headers
    df.reset_index(drop=True, inplace=True)

    # model 컬럼 기준 빈 행 제거
    model_col = find_column(df, ['model'])
    if model_col:
        df = df[df[model_col].notna()].copy()

    return df


# ════════════════════════════════════════════════════════════
# 실적 집계
# ════════════════════════════════════════════════════════════
def aggregate_actual(prod_df):
    df = prod_df.copy()
    df['MODEL_TYPE'] = df['FINAL_MAT_ID'].apply(classify_model_type)

    pd_actual = (df[(df['MODEL_TYPE'] == 'PD') & (df['OPER_DESC'] == 'P-ATE')]
                 .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())

    in1_actual = (df[(df['MODEL_TYPE'] == '3IN1') & (df['OPER_DESC'] == 'ASSY')]
                  .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())

    actual = pd.concat([pd_actual, in1_actual], ignore_index=True)
    actual = actual.rename(columns={'FINAL_MAT_ID': '_ERP_KEY', 'QTY': '누적실적'})
    return actual


# ════════════════════════════════════════════════════════════
# 알람 등급
# ════════════════════════════════════════════════════════════
def classify_alert(row, today):
    cutoff_cols = []
    for c in row.index:
        if 'cut off' in str(c).lower():
            val = row[c]
            try:
                if pd.notna(val) and float(val) != 0:
                    cutoff_cols.append(c)
            except Exception:
                continue

    days_left = 999
    for c in cutoff_cols:
        try:
            col_str = str(c)
            for sep in ['Cut off.', 'cut off.']:
                if sep in col_str:
                    date_str = col_str.split(sep)[-1].strip().split(' ')[0]
                    d = pd.to_datetime(date_str, errors='coerce')
                    if pd.notna(d):
                        diff = (d.date() - today).days
                        if diff >= 0:
                            days_left = min(days_left, diff)
                    break
        except Exception:
            continue

    gap = row.get('NEW_GAP', 0)
    rate = row.get('달성률(%)', 0)

    try:
        gap = float(gap) if pd.notna(gap) else 0
    except Exception:
        gap = 0
    try:
        rate = float(rate) if pd.notna(rate) else 0
    except Exception:
        rate = 0

    if gap < 0 and days_left <= 1:
        return "🔴 긴급"
    elif gap < 0:
        return "🟠 부족"
    elif rate < 70 and days_left <= 2:
        return "🟡 지연"
    else:
        return "✅ 정상"


# ════════════════════════════════════════════════════════════
# 머지 + 계산
# ════════════════════════════════════════════════════════════
def build_alert_table(ship_df, prod_df):
    actual = aggregate_actual(prod_df)

    # ERP 컬럼 찾기
    erp_col = find_column(ship_df, ['ERP', '3in1Code(FG)', 'FG Code'])
    if erp_col is None:
        st.error("❌ 출하계획에서 ERP 컬럼을 찾지 못했습니다.")
        return pd.DataFrame()

    # ERP 키 통일
    ship_work = ship_df.copy()
    ship_work['_ERP_KEY'] = ship_work[erp_col].astype(str).str.strip()
    actual['_ERP_KEY'] = actual['_ERP_KEY'].astype(str).str.strip()

    merged = ship_work.merge(actual, on='_ERP_KEY', how='left')
    merged['누적실적'] = pd.to_numeric(merged['누적실적'], errors='coerce').fillna(0)
    merged['MODEL_TYPE'] = merged['_ERP_KEY'].apply(classify_model_type)

    # SO (출하 필요량)
    so_col = find_column(merged, ['SO', 'TTL Ship'])
    if so_col:
        merged['SO'] = pd.to_numeric(merged[so_col], errors='coerce').fillna(0)
    else:
        merged['SO'] = 0.0

    # 기초재고
    stock_col = find_column(merged, ['base stock', 'O/stock', 'stock'])
    if stock_col:
        merged['기초재고'] = pd.to_numeric(merged[stock_col], errors='coerce').fillna(0)
    else:
        merged['기초재고'] = 0.0

    # NEW GAP
    merged['NEW_GAP'] = merged['기초재고'] + merged['누적실적'] - merged['SO']

    # 달성률 (Plan 기준)
    plan_col = find_column(merged, ['Plan.W25', 'Plan', 'plan'])
    if plan_col:
        plan = pd.to_numeric(merged[plan_col], errors='coerce')
        rate = merged['누적실적'] / plan.where(plan != 0)
        rate = pd.to_numeric(rate, errors='coerce').replace(
            [float('inf'), -float('inf')], 0).fillna(0)
        merged['달성률(%)'] = (rate * 100).round(1)
    else:
        merged['달성률(%)'] = 0.0

    today = datetime.today().date()
    merged['알람'] = merged.apply(lambda r: classify_alert(r, today), axis=1)

    return merged


# ════════════════════════════════════════════════════════════
# 메인 렌더링
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD 모델: P-ATE 완성 기준  |  3in1 모델: ASSY 완성 기준  |  💾 세션 저장")

    # ───── 1. DB 상태 ─────
    ship_db = st.session_state.get('ship_db', pd.DataFrame())
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_updated = st.session_state.get('ship_updated', '미상')
    prod_updated = st.session_state.get('prod_updated', '미상')

    st.markdown("##### 📊 데이터 현황")
    s1, s2 = st.columns(2)
    with s1:
        if not ship_db.empty:
            st.success(f"✅ **출하계획**: {len(ship_db):,}건  |  업데이트: {ship_updated}")
        else:
            st.warning("⚠️ **출하계획**: 데이터 없음")
    with s2:
        if not prod_db.empty:
            st.success(f"✅ **생산실적**: {len(prod_db):,}건  |  업데이트: {prod_updated}")
        else:
            st.warning("⚠️ **생산실적**: 데이터 없음")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 2. 파일 업로드 (통합형) ─────
    st.markdown("##### 📤 파일 업로드")
    st.caption("📌 .xlsx(출하계획)와 .csv(생산실적) 동시/개별 업로드 가능 — 자동 인식")

    uploaded_files = st.file_uploader(
        " ",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="ship_alert_multi",
        label_visibility="collapsed"
    )

    if uploaded_files:
        ship_uploaded = None
        prod_uploaded = None

        for f in uploaded_files:
            if f.name.lower().endswith('.xlsx'):
                ship_uploaded = f
            elif f.name.lower().endswith('.csv'):
                prod_uploaded = f

        pc1, pc2 = st.columns(2)
        with pc1:
            if ship_uploaded:
                st.success(f"📁 출하계획: **{ship_uploaded.name}** ({ship_uploaded.size/1024:.1f}KB)")
            else:
                st.caption("📁 출하계획(.xlsx) 미선택")
        with pc2:
            if prod_uploaded:
                st.success(f"📊 생산실적: **{prod_uploaded.name}** ({prod_uploaded.size/1024:.1f}KB)")
            else:
                st.caption("📊 생산실적(.csv) 미선택")

        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="ship_apply"):
            saved = []

            if ship_uploaded:
                try:
                    new_ship = smart_load_sheet1(ship_uploaded)
                    st.session_state['ship_db'] = new_ship
                    st.session_state['ship_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                    saved.append(f"출하계획 {len(new_ship)}건")
                except Exception as e:
                    st.error(f"❌ 출하계획 저장 실패: {e}")
                    import traceback
                    with st.expander("상세 에러"):
                        st.code(traceback.format_exc())

            if prod_uploaded:
                try:
                    new_prod = pd.read_csv(prod_uploaded)
                    st.session_state['prod_db'] = new_prod
                    st.session_state['prod_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                    saved.append(f"생산실적 {len(new_prod)}건")
                except Exception as e:
                    st.error(f"❌ 생산실적 저장 실패: {e}")

            if saved:
                st.success(f"✅ 저장 완료: {', '.join(saved)}")
                st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 3. 분석 결과 ─────
    if ship_db.empty or prod_db.empty:
        st.info("👆 출하계획과 생산실적이 모두 저장되면 분석 결과가 표시됩니다.")
        return

    st.markdown("### 🚨 분석 결과")

    # ───── 디버그 (필요 시) ─────
    with st.expander("🔧 디버그 정보 (컬럼 매칭 결과)", expanded=False):
        st.markdown("**출하계획 컬럼 목록:**")
        st.code(", ".join(str(c) for c in ship_db.columns))
        erp_col = find_column(ship_db, ['ERP'])
        st.markdown(f"**🔑 ERP 매칭 컬럼: `{erp_col}`**")
        st.markdown("**출하계획 미리보기 (처음 3행):**")
        st.dataframe(ship_db.head(3), use_container_width=True)

    # ───── 4. 분석 실행 ─────
    try:
        merged = build_alert_table(ship_db, prod_db)
    except Exception as e:
        st.error(f"❌ 데이터 처리 실패: {e}")
        import traceback
        with st.expander("상세 에러"):
            st.code(traceback.format_exc())
        return

    if merged.empty:
        st.warning("표시할 데이터가 없습니다.")
        return

    # ───── 5. KPI ─────
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(merged):,}건")
    k2.metric("🔴 긴급", int((merged['알람'] == '🔴 긴급').sum()))
    k3.metric("🟠 부족", int((merged['알람'] == '🟠 부족').sum()))
    k4.metric("🟡 지연", int((merged['알람'] == '🟡 지연').sum()))
    k5.metric("✅ 정상", int((merged['알람'] == '✅ 정상').sum()))

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 6. 필터 ─────
    f1, f2 = st.columns(2)
    alert_sel = f1.multiselect("알람 등급",
                                ['🔴 긴급', '🟠 부족', '🟡 지연', '✅ 정상'],
                                key="ship_grade")
    type_sel = f2.multiselect("모델 타입", ['PD', '3IN1'], key="ship_type")

    view = merged.copy()
    if alert_sel:
        view = view[view['알람'].isin(alert_sel)]
    if type_sel:
        view = view[view['MODEL_TYPE'].isin(type_sel)]

    # ───── 7. 알람 테이블 ─────
    st.markdown("##### 📋 알람 상세")

    # 컬럼 정리
    view_display = view.copy()
    model_col = find_column(view_display, ['model'])
    code_col = find_column(view_display, ['code'])
    note_col = find_column(view_display, ['Note'])

    rename_map = {'_ERP_KEY': 'ERP'}
    if model_col and model_col != 'model':
        rename_map[model_col] = 'model'
    if code_col and code_col != 'code':
        rename_map[code_col] = 'code'
    if note_col and note_col != 'Note':
        rename_map[note_col] = 'Note'
    view_display = view_display.rename(columns=rename_map)

    show_cols_pref = ['알람', 'model', 'code', 'ERP', 'MODEL_TYPE',
                      'SO', '기초재고', '누적실적', '달성률(%)', 'NEW_GAP', 'Note']
    show_cols = [c for c in show_cols_pref if c in view_display.columns]

    def color_alert_row(row):
        colors = {'🔴 긴급': '#ffcccc', '🟠 부족': '#ffe0b3',
                  '🟡 지연': '#fff5b3', '✅ 정상': '#d4edda'}
        bg = colors.get(row['알람'], '')
        return [f'background-color: {bg}'] * len(row)

    if not view_display.empty:
        st.dataframe(
            view_display[show_cols].style.apply(color_alert_row, axis=1),
            use_container_width=True, height=450
        )
    else:
        st.info("표시할 데이터가 없습니다.")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 8. 차트 ─────
    st.markdown("##### 📈 분석 차트")
    ch1, ch2 = st.columns(2)

    with ch1:
        st.markdown("**모델별 달성률 TOP 15**")
        model_col_v = find_column(view, ['model'])
        if not view.empty and model_col_v:
            chart1 = view[[model_col_v, '달성률(%)']].dropna().sort_values('달성률(%)').head(15)
            fig1 = px.bar(chart1, x='달성률(%)', y=model_col_v, orientation='h',
                          color='달성률(%)', color_continuous_scale='RdYlGn',
                          range_color=[0, 150], height=400)
            fig1.add_vline(x=100, line_dash="dash", line_color="green")
            fig1.update_layout(margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig1, use_container_width=True)

    with ch2:
        st.markdown("**알람 등급별 분포**")
        if not view.empty:
            chart2 = view['알람'].value_counts().reset_index()
            chart2.columns = ['알람', '건수']
            color_map = {'🔴 긴급': '#ef4444', '🟠 부족': '#f97316',
                         '🟡 지연': '#eab308', '✅ 정상': '#22c55e'}
            fig2 = px.bar(chart2, x='알람', y='건수', color='알람',
                          color_discrete_map=color_map, height=400)
            fig2.update_layout(margin=dict(l=0, r=0, t=10, b=0), showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)

    # ───── 9. DB 초기화 ─────
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("##### 🗑️ DB 관리")
    r1, r2, r3 = st.columns([1, 1, 3])
    with r1:
        if st.button("🗑️ 출하계획 초기화", use_container_width=True, key="reset_ship"):
            st.session_state.pop('ship_db', None)
            st.session_state.pop('ship_updated', None)
            st.rerun()
    with r2:
        if st.button("🗑️ 생산실적 초기화", use_container_width=True, key="reset_prod"):
            st.session_state.pop('prod_db', None)
            st.session_state.pop('prod_updated', None)
            st.rerun()
