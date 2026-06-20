"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 모듈 v2.0
- PD 모델: P-ATE 완성 기준
- 3in1 모델: ASSY 완성 기준
- 업로드 → 적용하기 버튼 방식
"""
import io
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px


# ════════════════════════════════════════════════════════════
# 모델 타입 판별 (PD vs 3in1)
# ════════════════════════════════════════════════════════════
def classify_model_type(code_3in1):
    if pd.isna(code_3in1):
        return "UNKNOWN"
    code = str(code_3in1).strip()
    if code.startswith("018"):
        return "3IN1"
    elif code.startswith("01"):
        return "PD"
    return "OTHER"


# ════════════════════════════════════════════════════════════
# 실적 집계 (PD=P-ATE, 3in1=ASSY)
# ════════════════════════════════════════════════════════════
def aggregate_actual(prod_df):
    df = prod_df.copy()
    df['MODEL_TYPE'] = df['FINAL_MAT_ID'].apply(classify_model_type)

    pd_actual = (df[(df['MODEL_TYPE'] == 'PD') & (df['OPER_DESC'] == 'P-ATE')]
                 .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())

    in1_actual = (df[(df['MODEL_TYPE'] == '3IN1') & (df['OPER_DESC'] == 'ASSY')]
                  .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())

    actual = pd.concat([pd_actual, in1_actual], ignore_index=True)
    actual = actual.rename(columns={'FINAL_MAT_ID': '3in1Code(FG)',
                                     'QTY': '누적실적'})
    return actual


# ════════════════════════════════════════════════════════════
# 알람 등급
# ════════════════════════════════════════════════════════════
def classify_alert(row, today):
    cutoff = pd.to_datetime(row.get('Cut off Cargo'), errors='coerce')
    if pd.isna(cutoff):
        return "⚪ 미정"
    days_left = (cutoff.date() - today).days
    gap = row.get('NEW_GAP', 0)
    rate = row.get('달성률(%)', 0)

    if gap < 0 and days_left <= 1:
        return "🔴 긴급"
    elif gap < 0:
        return "🟠 부족"
    elif rate < 70 and days_left <= 2:
        return "🟡 지연"
    else:
        return "✅ 정상"


# ════════════════════════════════════════════════════════════
# 헤더 자동 감지하여 엑셀 읽기
# ════════════════════════════════════════════════════════════
def smart_read_excel(file, sheet_name):
    """3in1Code(FG) 또는 TTL Ship 컬럼이 있는 행을 헤더로 자동 감지"""
    for header_row in [0, 1, 2, 3, 4]:
        try:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            cols_str = ' '.join(str(c) for c in df.columns)
            if '3in1Code' in cols_str and 'TTL Ship' in cols_str:
                return df, header_row
        except Exception:
            continue
    # 못 찾으면 header=2 기본값
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=2)
    df.columns = [str(c).strip() for c in df.columns]
    return df, 2


# ════════════════════════════════════════════════════════════
# 머지 + 계산
# ════════════════════════════════════════════════════════════
def build_alert_table(ship_df, prod_df, stock_df=None):
    actual = aggregate_actual(prod_df)

    merged = ship_df.merge(actual, on='3in1Code(FG)', how='left')
    merged['누적실적'] = merged['누적실적'].fillna(0)

    if stock_df is not None and not stock_df.empty:
        stock_col = [c for c in stock_df.columns if 'stock' in str(c).lower()]
        if stock_col:
            merged = merged.merge(
                stock_df[['3in1Code(FG)', stock_col[0]]],
                on='3in1Code(FG)', how='left'
            )
            merged['기초재고'] = pd.to_numeric(merged[stock_col[0]], errors='coerce').fillna(0)
        else:
            merged['기초재고'] = 0
    else:
        merged['기초재고'] = 0

    merged['MODEL_TYPE'] = merged['3in1Code(FG)'].apply(classify_model_type)

    # 숫자 변환
    merged['TTL Ship'] = pd.to_numeric(merged['TTL Ship'], errors='coerce').fillna(0)
    merged['누적실적'] = pd.to_numeric(merged['누적실적'], errors='coerce').fillna(0)

    merged['NEW_GAP'] = (merged['기초재고']
                         + merged['누적실적']
                         - merged['TTL Ship'])

    # Plan 컬럼 자동 감지
    plan_col = None
    for c in merged.columns:
        if 'Plan' in str(c) or 'plan' in str(c):
            plan_col = c
            break

    if plan_col:
        plan = pd.to_numeric(merged[plan_col], errors='coerce').replace(0, pd.NA)
        merged['달성률(%)'] = (merged['누적실적'] / plan * 100).round(1).fillna(0)
    else:
        merged['달성률(%)'] = 0

    today = datetime.today().date()
    merged['알람'] = merged.apply(lambda r: classify_alert(r, today), axis=1)

    return merged


# ════════════════════════════════════════════════════════════
# 메인 렌더링 (업로드 → 적용하기 버튼 방식)
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD 모델: P-ATE 완성 기준  |  3in1 모델: ASSY 완성 기준")

    # ───── 1. 파일 업로드 ─────
    up1, up2 = st.columns(2)
    with up1:
        ship_file = st.file_uploader(
            "📁 출하계획 (Shipment.xlsx)",
            type=["xlsx"], key="ship_alert_upload"
        )
    with up2:
        prod_file = st.file_uploader(
            "📁 생산실적 (PROD_RESULT.csv)",
            type=["csv"], key="prod_alert_upload"
        )

    # ───── 2. 적용하기 버튼 ─────
    apply_col1, apply_col2 = st.columns([1, 5])
    with apply_col1:
        apply_btn = st.button(
            "🚀 적용하기",
            type="primary",
            use_container_width=True,
            disabled=(ship_file is None or prod_file is None),
            key="ship_alert_apply"
        )

    # 두 파일 모두 없으면 안내만 출력 후 종료
    if ship_file is None or prod_file is None:
        st.info("👆 출하계획(.xlsx)과 생산실적(.csv)을 업로드하면 [적용하기] 버튼이 활성화됩니다.")
        return

    # 적용하기 클릭 시 세션에 저장
    if apply_btn:
        st.session_state['ship_alert_applied'] = True

    # 아직 적용 전이면 안내만 출력 후 종료
    if not st.session_state.get('ship_alert_applied', False):
        st.success("✅ 파일 업로드 완료. [🚀 적용하기] 버튼을 눌러주세요.")
        return

    # ═══════════════════════════════════════════════════════
    # 여기서부터는 적용하기를 누른 경우에만 실행
    # ═══════════════════════════════════════════════════════
    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 3. 데이터 로드 (헤더 자동 감지) ─────
    with st.spinner("데이터 로드 중..."):
        try:
            ship_df, detected_header = smart_read_excel(ship_file, "Shipment Rev")
        except Exception:
            try:
                ship_df, detected_header = smart_read_excel(ship_file, 0)
            except Exception as e:
                st.error(f"❌ 출하계획 파일 로드 실패: {e}")
                return

        try:
            ship_file.seek(0)
            stock_df = pd.read_excel(ship_file, sheet_name="Stock")
        except Exception:
            stock_df = pd.DataFrame()

        try:
            prod_df = pd.read_csv(prod_file)
        except Exception as e:
            st.error(f"❌ 생산실적 파일 로드 실패: {e}")
            return

    # ───── 4. 컬럼 검증 ─────
    required_ship = ['3in1Code(FG)', 'TTL Ship']
    missing_ship = [c for c in required_ship if c not in ship_df.columns]

    required_prod = ['FINAL_MAT_ID', 'OPER_DESC', 'QTY']
    missing_prod = [c for c in required_prod if c not in prod_df.columns]

    if missing_ship or missing_prod:
        if missing_ship:
            st.error(f"❌ 출하계획 누락 컬럼: {missing_ship}")
            st.caption(f"보유 컬럼: {list(ship_df.columns)[:15]}...")
        if missing_prod:
            st.error(f"❌ 생산실적 누락 컬럼: {missing_prod}")
            st.caption(f"보유 컬럼: {list(prod_df.columns)}")
        with st.expander("🔧 디버그: Ship DF 미리보기"):
            st.dataframe(ship_df.head())
        return

    # ───── 5. 알람 테이블 생성 ─────
    try:
        merged = build_alert_table(ship_df, prod_df, stock_df)
    except Exception as e:
        st.error(f"❌ 데이터 처리 실패: {e}")
        return

    # ───── 6. KPI 카드 ─────
    st.markdown("##### 📊 알람 현황")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(merged):,}건")
    k2.metric("🔴 긴급", int((merged['알람'] == '🔴 긴급').sum()))
    k3.metric("🟠 부족", int((merged['알람'] == '🟠 부족').sum()))
    k4.metric("🟡 지연", int((merged['알람'] == '🟡 지연').sum()))
    k5.metric("✅ 정상", int((merged['알람'] == '✅ 정상').sum()))

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 7. 필터 ─────
    f1, f2, f3 = st.columns(3)
    cus_options = sorted(merged['Cus'].dropna().unique()) if 'Cus' in merged.columns else []
    cus_sel = f1.multiselect("거래선", cus_options, key="ship_alert_cus")
    alert_sel = f2.multiselect("알람 등급",
                                ['🔴 긴급', '🟠 부족', '🟡 지연', '✅ 정상'],
                                key="ship_alert_grade")
    type_sel = f3.multiselect("모델 타입", ['PD', '3IN1'], key="ship_alert_type")

    view = merged.copy()
    if cus_sel and 'Cus' in view.columns:
        view = view[view['Cus'].isin(cus_sel)]
    if alert_sel:
        view = view[view['알람'].isin(alert_sel)]
    if type_sel:
        view = view[view['MODEL_TYPE'].isin(type_sel)]

    # ───── 8. 알람 테이블 ─────
    st.markdown("##### 📋 알람 상세")
    show_cols_pref = ['알람', 'Cus', 'Model', 'MODEL_TYPE',
                      'Cut off Cargo', 'HQ Request',
                      'TTL Ship', '누적실적', '달성률(%)', 'NEW_GAP']
    show_cols = [c for c in show_cols_pref if c in view.columns]

    def color_alert_row(row):
        colors = {'🔴 긴급': '#ffcccc', '🟠 부족': '#ffe0b3',
                  '🟡 지연': '#fff5b3', '✅ 정상': '#d4edda'}
        bg = colors.get(row['알람'], '')
        return [f'background-color: {bg}'] * len(row)

    if not view.empty:
        st.dataframe(
            view[show_cols].style.apply(color_alert_row, axis=1),
            use_container_width=True, height=450
        )
    else:
        st.info("표시할 데이터가 없습니다.")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 9. 차트 ─────
    st.markdown("##### 📈 분석 차트")
    ch1, ch2 = st.columns(2)

    with ch1:
        st.markdown("**거래선별 평균 달성률**")
        if not view.empty and 'Cus' in view.columns:
            chart1 = (view.groupby('Cus')['달성률(%)'].mean()
                      .reset_index().sort_values('달성률(%)'))
            fig1 = px.bar(chart1, x='달성률(%)', y='Cus', orientation='h',
                          color='달성률(%)', color_continuous_scale='RdYlGn',
                          range_color=[0, 150], height=350)
            fig1.add_vline(x=100, line_dash="dash", line_color="green",
                           annotation_text="100%")
            fig1.update_layout(margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("데이터 없음")

    with ch2:
        st.markdown("**Cut-off 일자별 부족 수량**")
        shortage = view[view['NEW_GAP'] < 0].copy() if not view.empty else pd.DataFrame()
        if not shortage.empty and 'Cut off Cargo' in shortage.columns:
            shortage['부족수량'] = shortage['NEW_GAP'].abs()
            chart2 = (shortage.groupby('Cut off Cargo')['부족수량']
                      .sum().reset_index())
            chart2['Cut off Cargo'] = chart2['Cut off Cargo'].astype(str)
            fig2 = px.bar(chart2, x='Cut off Cargo', y='부족수량',
                          color='부족수량', color_continuous_scale='Reds',
                          height=350)
            fig2.update_layout(margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.success("✅ 부족 항목 없음")

    # ───── 10. 초기화 버튼 ─────
    st.markdown("<hr>", unsafe_allow_html=True)
    reset_col1, reset_col2 = st.columns([1, 5])
    with reset_col1:
        if st.button("🔄 초기화", use_container_width=True, key="ship_alert_reset"):
            st.session_state['ship_alert_applied'] = False
            st.rerun()
