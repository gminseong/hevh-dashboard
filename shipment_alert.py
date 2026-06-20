"""
Shipment Cut-off 알람 모듈
- PD 모델: P-ATE 공정 완성 기준
- 3in1 모델: ASSY 공정 완성 기준
"""
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import io


# ════════════════════════════════════════════════════════════
# 1. 모델 타입 판별 (PD vs 3in1)
# ════════════════════════════════════════════════════════════
def classify_model_type(code_3in1: str) -> str:
    """
    3in1Code(FG) 앞자리로 PD/3in1 구분
    - 01xxxxxx = PD (Power Display)
    - 018xxxxx = 3in1
    """
    if pd.isna(code_3in1):
        return "UNKNOWN"
    code = str(code_3in1)
    if code.startswith("018"):
        return "3IN1"
    elif code.startswith("01"):
        return "PD"
    return "OTHER"


# ════════════════════════════════════════════════════════════
# 2. 실적 집계 (PD=P-ATE, 3in1=ASSY)
# ════════════════════════════════════════════════════════════
def aggregate_actual(prod_df: pd.DataFrame) -> pd.DataFrame:
    """
    PROD_RESULT에서 모델 타입별로 다른 공정 기준 집계
    """
    df = prod_df.copy()
    df['MODEL_TYPE'] = df['FINAL_MAT_ID'].apply(classify_model_type)
    
    # PD: P-ATE 완성 기준
    pd_actual = (df[(df['MODEL_TYPE']=='PD') & (df['OPER_DESC']=='P-ATE')]
                 .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())
    
    # 3in1: ASSY 완성 기준
    in1_actual = (df[(df['MODEL_TYPE']=='3IN1') & (df['OPER_DESC']=='ASSY')]
                  .groupby('FINAL_MAT_ID')['QTY'].sum().reset_index())
    
    actual = pd.concat([pd_actual, in1_actual], ignore_index=True)
    actual = actual.rename(columns={'FINAL_MAT_ID':'3in1Code(FG)',
                                     'QTY':'누적실적'})
    return actual


# ════════════════════════════════════════════════════════════
# 3. 알람 등급 분류
# ════════════════════════════════════════════════════════════
def classify_alert(row, today: date) -> str:
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
# 4. 메인 머지 + 계산
# ════════════════════════════════════════════════════════════
def build_alert_table(ship_df: pd.DataFrame, 
                      prod_df: pd.DataFrame,
                      stock_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    출하계획 + 실적 + 재고 머지 후 GAP 재계산
    """
    actual = aggregate_actual(prod_df)
    
    merged = ship_df.merge(actual, on='3in1Code(FG)', how='left')
    merged['누적실적'] = merged['누적실적'].fillna(0)
    
    # 재고 머지 (있을 경우)
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
    
    # 모델 타입 판별
    merged['MODEL_TYPE'] = merged['3in1Code(FG)'].apply(classify_model_type)
    
    # NEW GAP = 기초재고 + 누적실적 - TTL Ship
    merged['NEW_GAP'] = (merged['기초재고'] 
                         + merged['누적실적'] 
                         - merged['TTL Ship'])
    
    # 달성률
    plan = merged.get('Plan.W25', 0)
    if isinstance(plan, pd.Series):
        merged['달성률(%)'] = (merged['누적실적'] 
                              / plan.replace(0, pd.NA) * 100).round(1).fillna(0)
    
    # 알람 등급
    today = datetime.today().date()
    merged['알람'] = merged.apply(lambda r: classify_alert(r, today), axis=1)
    
    return merged


# ════════════════════════════════════════════════════════════
# 5. 이메일 발송
# ════════════════════════════════════════════════════════════
def send_alert_email(df: pd.DataFrame, 
                     receiver: str,
                     smtp_config: dict) -> bool:
    """
    긴급/부족 알람 메일 발송
    smtp_config: {host, port, user, password, sender}
    """
    urgent = df[df['알람'].isin(['🔴 긴급','🟠 부족'])]
    if urgent.empty:
        return False
    
    # HTML 본문
    html = f"""
    <html><body style="font-family:맑은 고딕,sans-serif;">
    <h2 style="color:#d32f2f;">🚨 Shipment Cut-off 알람</h2>
    <p>발생 일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
    <p><b>긴급 {(df['알람']=='🔴 긴급').sum()}건 / 부족 {(df['알람']=='🟠 부족').sum()}건</b></p>
    <table border="1" cellpadding="6" cellspacing="0" 
           style="border-collapse:collapse; width:100%;">
        <tr style="background:#1976d2; color:white;">
            <th>알람</th><th>거래선</th><th>모델</th><th>Cut-off</th>
            <th>TTL Ship</th><th>누적실적</th><th>달성률</th><th>NEW GAP</th>
        </tr>
    """
    for _, row in urgent.iterrows():
        color = "#ffcccc" if row['알람']=='🔴 긴급' else "#ffe0b3"
        html += f"""
        <tr style="background:{color};">
            <td align="center">{row['알람']}</td>
            <td>{row.get('Cus','')}</td>
            <td>{row.get('Model','')}</td>
            <td align="center">{row.get('Cut off Cargo','')}</td>
            <td align="right">{int(row.get('TTL Ship',0)):,}</td>
            <td align="right">{int(row.get('누적실적',0)):,}</td>
            <td align="right">{row.get('달성률(%)',0)}%</td>
            <td align="right" style="color:red;"><b>{int(row.get('NEW_GAP',0)):,}</b></td>
        </tr>
        """
    html += "</table><p>본 메일은 HEVH Dashboard에서 자동 발송됩니다.</p></body></html>"
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"[HEVH 알람] Shipment Cut-off {date.today()}"
    msg['From'] = smtp_config['sender']
    msg['To'] = receiver
    msg.attach(MIMEText(html, 'html'))
    
    try:
        with smtplib.SMTP(smtp_config['host'], smtp_config['port']) as server:
            server.starttls()
            server.login(smtp_config['user'], smtp_config['password'])
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"메일 발송 실패: {e}")
        return False


# ════════════════════════════════════════════════════════════
# 6. Streamlit UI 렌더링
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab(prod_df: pd.DataFrame = None):
    """
    Streamlit 탭에서 호출하는 메인 함수
    prod_df: 기존 대시보드에서 로드한 PROD_RESULT 데이터
    """
    st.markdown("## 🚨 Shipment Cut-off 알람")
    st.caption("PD 모델: P-ATE 완성 기준 | 3in1 모델: ASSY 완성 기준")
    
    # ───── 1. 파일 업로드 ─────
    col1, col2 = st.columns(2)
    with col1:
        ship_file = st.file_uploader(
            "📁 출하계획 (Shipment.xlsx)",
            type=["xlsx"], key="ship_upload"
        )
    with col2:
        if prod_df is None:
            prod_file = st.file_uploader(
                "📁 생산실적 (PROD_RESULT.csv)",
                type=["csv"], key="prod_upload"
            )
            if prod_file:
                prod_df = pd.read_csv(prod_file)
        else:
            st.success(f"✅ 기존 DB 실적 로드: {len(prod_df):,}건")
    
    if ship_file is None or prod_df is None:
        st.info("👆 출하계획과 생산실적 파일을 모두 준비해주세요.")
        return
    
    # ───── 2. 데이터 로드 ─────
    try:
        ship_df = pd.read_excel(ship_file, sheet_name="Shipment Rev")
    except Exception:
        ship_df = pd.read_excel(ship_file, sheet_name=0)
    
    try:
        stock_df = pd.read_excel(ship_file, sheet_name="Stock")
    except Exception:
        stock_df = pd.DataFrame()
    
    # ───── 3. 알람 테이블 생성 ─────
    merged = build_alert_table(ship_df, prod_df, stock_df)
    
    # ───── 4. KPI 카드 ─────
    st.markdown("### 📊 알람 현황")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📦 전체", f"{len(merged):,}건")
    c2.metric("🔴 긴급", (merged['알람']=='🔴 긴급').sum())
    c3.metric("🟠 부족", (merged['알람']=='🟠 부족').sum())
    c4.metric("🟡 지연", (merged['알람']=='🟡 지연').sum())
    c5.metric("✅ 정상", (merged['알람']=='✅ 정상').sum())
    
    # ───── 5. 필터 ─────
    st.markdown("### 🔍 필터")
    f1, f2, f3 = st.columns(3)
    cus_sel = f1.multiselect("거래선", sorted(merged['Cus'].dropna().unique()))
    alert_sel = f2.multiselect("알람 등급", 
                                ['🔴 긴급','🟠 부족','🟡 지연','✅ 정상'])
    type_sel = f3.multiselect("모델 타입", ['PD','3IN1'])
    
    view = merged.copy()
    if cus_sel: view = view[view['Cus'].isin(cus_sel)]
    if alert_sel: view = view[view['알람'].isin(alert_sel)]
    if type_sel: view = view[view['MODEL_TYPE'].isin(type_sel)]
    
    # ───── 6. 테이블 ─────
    st.markdown("### 📋 알람 상세")
    show_cols = ['알람','Cus','Model','MODEL_TYPE','Cut off Cargo',
                 'HQ Request','TTL Ship','누적실적','달성률(%)','NEW_GAP']
    show_cols = [c for c in show_cols if c in view.columns]
    
    def color_alert_row(row):
        colors = {'🔴 긴급':'#ffcccc','🟠 부족':'#ffe0b3',
                  '🟡 지연':'#fff5b3','✅ 정상':'#d4edda'}
        bg = colors.get(row['알람'], '')
        return [f'background-color: {bg}'] * len(row)
    
    st.dataframe(
        view[show_cols].style.apply(color_alert_row, axis=1),
        use_container_width=True, height=450
    )
    
    # ───── 7. 차트 ─────
    st.markdown("### 📈 분석 차트")
    ch1, ch2 = st.columns(2)
    
    with ch1:
        st.markdown("#### 거래선별 평균 달성률")
        if not view.empty:
            chart1 = (view.groupby('Cus')['달성률(%)'].mean()
                      .reset_index().sort_values('달성률(%)'))
            fig1 = px.bar(chart1, x='달성률(%)', y='Cus', orientation='h',
                          color='달성률(%)', color_continuous_scale='RdYlGn',
                          range_color=[0,150])
            fig1.add_vline(x=100, line_dash="dash", line_color="green")
            st.plotly_chart(fig1, use_container_width=True)
    
    with ch2:
        st.markdown("#### Cut-off 일자별 부족 수량")
        shortage = view[view['NEW_GAP']<0].copy()
        if not shortage.empty:
            shortage['부족수량'] = shortage['NEW_GAP'].abs()
            chart2 = (shortage.groupby('Cut off Cargo')['부족수량']
                      .sum().reset_index())
            fig2 = px.bar(chart2, x='Cut off Cargo', y='부족수량',
                          color='부족수량', color_continuous_scale='Reds')
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.success("✅ 부족 항목 없음")
    
    # ───── 8. 액션 버튼 ─────
    st.markdown("### 🎬 액션")
    a1, a2, a3 = st.columns(3)
    
    with a1:
        # CSV 다운로드
        csv = view.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 CSV 다운로드", csv,
            f"shipment_alert_{date.today()}.csv",
            "text/csv"
        )
    
    with a2:
        # Excel 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            view.to_excel(writer, index=False, sheet_name='Alert')
        st.download_button(
            "📊 Excel 다운로드", buffer.getvalue(),
            f"shipment_alert_{date.today()}.xlsx"
        )
    
    with a3:
        # 메일 발송
        with st.expander("📧 알람 메일 발송"):
            receiver = st.text_input("수신자 이메일", "minseong@hansol.com")
            # SMTP 설정은 secrets.toml에서 로드
            if st.button("발송", type="primary"):
                if 'smtp' in st.secrets:
                    smtp_config = dict(st.secrets['smtp'])
                    ok = send_alert_email(view, receiver, smtp_config)
                    if ok:
                        st.success("✅ 메일 발송 완료")
                else:
                    st.warning("⚠️ .streamlit/secrets.toml에 SMTP 설정 필요")
