"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 모듈 v1.0
- PD 모델: P-ATE 완성 기준
- 3in1 모델: ASSY 완성 기준
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
    """3in1Code(FG) 앞자리로 PD/3in1 구분"""
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
    """모델 타입별로 다른 공정 기준 집계"""
    df = prod_df.copy()
    df['MODEL_TYPE'] = df['FINAL_MAT_ID'].apply(classify_model_type)

    # PD: P-ATE 완성 기준
    pd_actual = (df[(df['MODEL_TYPE'] == 'PD') & (df['OPER_DESC'] == 'P-ATE')]
                 .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())

    # 3in1: ASSY 완성 기준
    in1_actual = (df[(df['MODEL_TYPE'] == '3IN1') & (df['OPER_DESC'] == 'ASSY')]
                  .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())

    actual = pd.concat([pd_actual, in1_actual], ignore_index=True)
    actual = actual.rename(columns={'FINAL_MAT_ID': '3in1Code(FG)',
                                     'QTY': '누적실적'})
    return actual


# ════════════════════════════════════════════════════════════
# 알람 등급 분류
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
# 메인 머지 + 계산
# ════════════════════════════════════════════════════════════
def build_alert_table(ship_df, prod_df, stock_df=None):
    """출하계획 + 실적 + 재고 머지 후 GAP 재계산"""
    actual = aggregate_actual(prod_df)

    merged = ship_df.merge(actual, on='3in1Code(FG)', how='left')
    merged['누적실적'] = merged['누적실적'].fillna(0)

    # 재고 머지
    if stock_df is not None and not stock_df.empty:
        stock_col = [c for c in stock_df.columns if 'stock' in c.lower()]
        if stock_col:
            merged = merged.merge(
                stock_df[['3in1Code(FG)', stock_col[0]]],
                on='3in1Code(FG)', how='left'
            )
            merged['기초재고'] = merged[stock_col[0]].fillna(0)
        else:
            merged['기초재고'] = 0
    else:
        merged['기초재고'] = 0

    # 모델 타입
    merged['MODEL_TYPE'] = merged['3in1Code(FG)'].apply(classify_model_type)

    # NEW GAP = 기초재고 + 누적실적 - TTL Ship
    merged['NEW_GAP'] = (merged['기초재고']
                         + merged['누적실적']
                         - merged['TTL Ship'])

    # 달성률 (Plan 컬럼 자동 감지)
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

    # 알람 등급
    today = datetime.today().date()
    merged['알람'] = merged.apply(lambda r: classify_alert(r, today), axis=1)

    return merged


# ════════════════════════════════════════════════════════════
# Streamlit 탭 렌더링
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    """tab 안에서 호출하는 메인 함수"""
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 PD 모델: P-ATE 완성 기준  |  3in1 모델: ASSY 완성 기준")

    # ───── 파일 업로드 ─────
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

    if ship_file is None or prod_file is None:
        st.info("👆 출하계획(.xlsx)과 생산실적(.csv) 파일을 모두 업로드해주세요.")
        return

    # ───── 데이터 로드 (헤더 자동 감지) ─────
    def smart_read_excel(file, sheet):
        """헤더 행을 자동으로 찾아서 읽기"""
        for header_row in [0, 1, 2, 3]:
            try:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row)
                # '3in1Code(FG)' 또는 유사 컬럼이 있으면 OK
                cols_str = ' '.join(str(c) for c in df.columns)
                if '3in1Code' in cols_str or 'TTL Ship' in cols_str:
                    return df
            except Exception:
                continue
        # 못 찾으면 기본값(header=2) 시도
        return pd.read_excel(file, sheet_name=sheet, header=2)
    
    try:
        ship_df = smart_read_excel(ship_file, "Shipment Rev")
    except Exception:
        ship_df = smart_read_excel(ship_file, 0)

    try:
        stock_df = pd.read_excel(ship_file, sheet_name="Stock")
    except Exception:
        stock_df = pd.DataFrame()
    
    # 컬럼명 공백 정리
    ship_df.columns = [str(c).strip() for c in ship_df.columns]
    
    # ───── 컬럼 검증 ─────
    required_ship = ['3in1Code(FG)', 'TTL Ship']
    missing = [c for c in required_ship if c not in ship_df.columns]
    if missing:
        st.error(f"❌ 출하계획 파일에 필수 컬럼 누락: {missing}")
        st.info(f"보유 컬럼: {list(ship_df.columns)}")
        return

    required_prod = ['FINAL_MAT_ID', 'OPER_DESC', 'QTY']
    missing = [c for c in required_prod if c not in prod_df.columns]
    if missing:
        st.error(f"❌ 생산실적 파일에 필수 컬럼 누락: {missing}")
        st.info(f"보유 컬럼: {list(prod_df.columns)}")
        return

    # ───── 알람 테이블 생성 ─────
    merged = build_alert_table(ship_df, prod_df, stock_df)

    # ───── KPI 카드 ─────
    st.markdown("##### 📊 알람 현황")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(merged):,}건")
    k2.metric("🔴 긴급", int((merged['알람'] == '🔴 긴급').sum()))
    k3.metric("🟠 부족", int((merged['알람'] == '🟠 부족').sum()))
    k4.metric("🟡 지연", int((merged['알람'] == '🟡 지연').sum()))
    k5.metric("✅ 정상", int((merged['알람'] == '✅ 정상').sum()))

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 필터 ─────
    f1, f2, f3 = st.columns(3)
    cus_options = sorted(merged['Cus'].dropna().unique()) if 'Cus' in merged.columns else []
    cus_sel = f1.multiselect("거래선", cus_options)
    alert_sel = f2.multiselect("알람 등급",
                                ['🔴 긴급', '🟠 부족', '🟡 지연', '✅ 정상'])
    type_sel = f3.multiselect("모델 타입", ['PD', '3IN1'])

    view = merged.copy()
    if cus_sel and 'Cus' in view.columns:
        view = view[view['Cus'].isin(cus_sel)]
    if alert_sel:
        view = view[view['알람'].isin(alert_sel)]
    if type_sel:
        view = view[view['MODEL_TYPE'].isin(type_sel)]

    # ───── 알람 테이블 ─────
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

    # ───── 차트 ─────
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
                          annotation_text="목표 100%")
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

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 다운로드 ─────
    st.markdown("##### 📥 다운로드")
    d1, d2 = st.columns(2)
    today_str = datetime.today().strftime('%Y%m%d')

    with d1:
        csv = view.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 CSV 다운로드", csv,
            f"shipment_alert_{today_str}.csv",
            "text/csv",
            use_container_width=True
        )

    with d2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            view.to_excel(writer, index=False, sheet_name='Alert')
        st.download_button(
            "📊 Excel 다운로드", buffer.getvalue(),
            f"shipment_alert_{today_str}.xlsx",
            use_container_width=True
        )
