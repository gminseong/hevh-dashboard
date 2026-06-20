"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 모듈 v4.0
- DB 영구 저장 방식 (LOSSTIME과 동일 패턴)
- 각 파일 독립 업로드/갱신
- 자동 매칭 분석
"""
import io
import os
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px


# ════════════════════════════════════════════════════════════
# DB 파일 경로
# ════════════════════════════════════════════════════════════
SHIP_DB_PATH = "shipment_db.csv"
PROD_DB_PATH = "prod_result_db.csv"
META_DB_PATH = "shipment_meta.csv"  # 업로드 시각 기록


# ════════════════════════════════════════════════════════════
# DB 로드/저장
# ════════════════════════════════════════════════════════════
def load_ship_db():
    if os.path.exists(SHIP_DB_PATH):
        try:
            return pd.read_csv(SHIP_DB_PATH)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def load_prod_db():
    if os.path.exists(PROD_DB_PATH):
        try:
            return pd.read_csv(PROD_DB_PATH)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def load_meta():
    if os.path.exists(META_DB_PATH):
        try:
            return pd.read_csv(META_DB_PATH).set_index('key')['value'].to_dict()
        except Exception:
            return {}
    return {}


def save_meta(meta):
    pd.DataFrame(list(meta.items()), columns=['key', 'value']).to_csv(META_DB_PATH, index=False)


# ════════════════════════════════════════════════════════════
# 모델 타입
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
    actual = actual.rename(columns={'FINAL_MAT_ID': 'ERP', 'QTY': '누적실적'})
    return actual


# ════════════════════════════════════════════════════════════
# 알람 등급
# ════════════════════════════════════════════════════════════
def classify_alert(row, today):
    cutoff_cols = [c for c in row.index if 'Cut off' in str(c)
                   and pd.notna(row[c]) and row[c] != 0]
    days_left = 999
    for c in cutoff_cols:
        try:
            date_str = str(c).split('Cut off.')[-1].split(' ')[0]
            d = pd.to_datetime(date_str, errors='coerce')
            if pd.notna(d):
                diff = (d.date() - today).days
                if diff >= 0:
                    days_left = min(days_left, diff)
        except Exception:
            continue

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
# 머지 + 계산
# ════════════════════════════════════════════════════════════
def build_alert_table(ship_df, prod_df):
    actual = aggregate_actual(prod_df)

    if 'ERP' not in ship_df.columns:
        st.error("❌ 출하계획에 ERP 컬럼이 없습니다.")
        return pd.DataFrame()

    merged = ship_df.merge(actual, on='ERP', how='left')
    merged['누적실적'] = pd.to_numeric(merged['누적실적'], errors='coerce').fillna(0)
    merged['MODEL_TYPE'] = merged['ERP'].apply(classify_model_type)

    if 'SO' in merged.columns:
        merged['SO'] = pd.to_numeric(merged['SO'], errors='coerce').fillna(0)

    base_stock_col = None
    for c in merged.columns:
        if 'base stock' in str(c).lower():
            base_stock_col = c
            break

    if base_stock_col:
        merged['기초재고'] = pd.to_numeric(merged[base_stock_col], errors='coerce').fillna(0)
    else:
        merged['기초재고'] = 0.0

    merged['NEW_GAP'] = (merged['기초재고']
                         + merged['누적실적']
                         - merged.get('SO', 0))

    plan_col = None
    for c in merged.columns:
        if 'Plan' in str(c) or 'plan' in str(c):
            plan_col = c
            break

    if plan_col:
        plan = pd.to_numeric(merged[plan_col], errors='coerce')
        rate = merged['누적실적'] / plan.where(plan != 0)
        rate = pd.to_numeric(rate, errors='coerce').replace([float('inf'), -float('inf')], 0).fillna(0)
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
    st.caption("📌 PD 모델: P-ATE 완성 기준  |  3in1 모델: ASSY 완성 기준  |  💾 자동 저장 (각 파일 독립 갱신)")

    # ───── DB 상태 확인 ─────
    ship_db = load_ship_db()
    prod_db = load_prod_db()
    meta = load_meta()

    # ───── 1. DB 상태 표시 ─────
    st.markdown("##### 📊 데이터 현황")
    s1, s2 = st.columns(2)
    
    with s1:
        if not ship_db.empty:
            ship_time = meta.get('ship_updated', '미상')
            st.success(f"✅ **출하계획**: {len(ship_db):,}건  |  업데이트: {ship_time}")
        else:
            st.warning("⚠️ **출하계획**: 데이터 없음 — 아래에서 업로드해주세요")
    
    with s2:
        if not prod_db.empty:
            prod_time = meta.get('prod_updated', '미상')
            st.success(f"✅ **생산실적**: {len(prod_db):,}건  |  업데이트: {prod_time}")
        else:
            st.warning("⚠️ **생산실적**: 데이터 없음 — 아래에서 업로드해주세요")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 2. 파일 업로드 (독립 갱신) ─────
    st.markdown("##### 📤 파일 업로드 (필요한 것만 갱신)")
    
    up1, up2 = st.columns(2)
    
    with up1:
        st.markdown("**📁 출하계획 (Shipment.xlsx)**")
        ship_file = st.file_uploader(
            " ", type=["xlsx"], key="ship_upload", label_visibility="collapsed"
        )
        if ship_file is not None:
            if st.button("💾 출하계획 저장", type="primary", use_container_width=True, key="save_ship"):
                try:
                    new_ship = pd.read_excel(ship_file, sheet_name="Sheet1")
                    new_ship.columns = [str(c).strip() for c in new_ship.columns]
                    if 'model' in new_ship.columns:
                        new_ship = new_ship[new_ship['model'].notna()].copy()
                    new_ship.to_csv(SHIP_DB_PATH, index=False)
                    meta['ship_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                    save_meta(meta)
                    st.success(f"✅ 출하계획 저장 완료 ({len(new_ship):,}건)")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 저장 실패: {e}")

    with up2:
        st.markdown("**📁 생산실적 (PROD_RESULT.csv)**")
        prod_file = st.file_uploader(
            "  ", type=["csv"], key="prod_upload", label_visibility="collapsed"
        )
        if prod_file is not None:
            if st.button("💾 생산실적 저장", type="primary", use_container_width=True, key="save_prod"):
                try:
                    new_prod = pd.read_csv(prod_file)
                    new_prod.to_csv(PROD_DB_PATH, index=False)
                    meta['prod_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                    save_meta(meta)
                    st.success(f"✅ 생산실적 저장 완료 ({len(new_prod):,}건)")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 저장 실패: {e}")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 3. 자동 분석 (둘 다 있을 때만) ─────
    if ship_db.empty or prod_db.empty:
        st.info("👆 출하계획과 생산실적이 모두 저장되면 자동으로 분석 결과가 표시됩니다.")
        return

    st.markdown("### 🚨 분석 결과")

    try:
        merged = build_alert_table(ship_db, prod_db)
    except Exception as e:
        st.error(f"❌ 데이터 처리 실패: {e}")
        return

    if merged.empty:
        st.warning("표시할 데이터가 없습니다.")
        return

    # ───── KPI ─────
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 전체", f"{len(merged):,}건")
    k2.metric("🔴 긴급", int((merged['알람'] == '🔴 긴급').sum()))
    k3.metric("🟠 부족", int((merged['알람'] == '🟠 부족').sum()))
    k4.metric("🟡 지연", int((merged['알람'] == '🟡 지연').sum()))
    k5.metric("✅ 정상", int((merged['알람'] == '✅ 정상').sum()))

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 필터 ─────
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

    # ───── 테이블 ─────
    st.markdown("##### 📋 알람 상세")
    show_cols_pref = ['알람', 'model', 'code', 'ERP', 'MODEL_TYPE',
                      'SO', '기초재고', '누적실적', '달성률(%)', 'NEW_GAP', 'Note']
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

    st.markdown("<hr>", unsafe_allow_html=True)

    # ───── 차트 ─────
    st.markdown("##### 📈 분석 차트")
    ch1, ch2 = st.columns(2)
    
    with ch1:
        st.markdown("**모델별 달성률 TOP 15**")
        if not view.empty and 'model' in view.columns:
            chart1 = view[['model', '달성률(%)']].dropna().sort_values('달성률(%)').head(15)
            fig1 = px.bar(chart1, x='달성률(%)', y='model', orientation='h',
                          color='달성률(%)', color_continuous_scale='RdYlGn',
                          range_color=[0, 150], height=400)
            fig1.add_vline(x=100, line_dash="dash", line_color="green",
                           annotation_text="100%")
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

    # ───── DB 초기화 ─────
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("##### 🗑️ DB 관리")
    r1, r2 = st.columns(2)
    with r1:
        if st.button("🗑️ 출하계획 DB 초기화", use_container_width=True, key="reset_ship"):
            if os.path.exists(SHIP_DB_PATH):
                os.remove(SHIP_DB_PATH)
            meta.pop('ship_updated', None)
            save_meta(meta)
            st.success("출하계획 DB 초기화 완료")
            st.rerun()
    with r2:
        if st.button("🗑️ 생산실적 DB 초기화", use_container_width=True, key="reset_prod"):
            if os.path.exists(PROD_DB_PATH):
                os.remove(PROD_DB_PATH)
            meta.pop('prod_updated', None)
            save_meta(meta)
            st.success("생산실적 DB 초기화 완료")
            st.rerun()
