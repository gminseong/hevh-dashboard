"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v30.1
- 통합 비교 표 (문제만)
- 테이블 1: 계획 기준
- 테이블 2: 실적 반영 + Cutoff 시점 재고
- 정렬: components.html (JS 동작)
"""
from datetime import datetime
import io
import re
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


@st.cache_data(show_spinner=False)
def read_excel_raw(file_bytes, sheet_name):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def read_csv_cached(file_bytes):
    return pd.read_csv(io.BytesIO(file_bytes))


def classify(x):
    x = str(x).strip()
    if x.startswith('018'): return '3IN1'
    if x.startswith('01'): return 'PD'
    return 'OTHER'


def parse_date_from_col(col_name, year=2026):
    try:
        s = str(col_name).lower()
        s = re.sub(r'(plan|actual|cut\s*off|cargo)[\.\s_&]*', '', s).strip()
        nums = re.findall(r'\d+', s)
        if len(nums) >= 2:
            month, day = int(nums[0]), int(nums[1])
            if 1 <= month <= 12 and 1 <= day <= 31:
                return pd.Timestamp(year, month, day)
        elif len(nums) == 1:
            day = int(nums[0])
            if 1 <= day <= 31:
                return pd.Timestamp(year, 6, day)
        return None
    except Exception:
        return None


def merge_two_row_header(raw, header_row):
    main = raw.iloc[header_row].values
    sub = raw.iloc[header_row + 1].values if header_row + 1 < len(raw) else [None] * len(main)
    merged = []
    for mh, sh in zip(main, sub):
        m_str = str(mh).strip() if pd.notna(mh) else ''
        s_str = str(sh).strip() if pd.notna(sh) else ''
        if m_str in ('nan', 'None'): m_str = ''
        if s_str in ('nan', 'None'): s_str = ''
        if m_str and s_str:
            merged.append(f"{m_str}.{s_str}")
        elif m_str:
            merged.append(m_str)
        elif s_str:
            merged.append(s_str)
        else:
            merged.append(f"col_{len(merged)}")
    return merged


def load_sheet1_notes(file_bytes):
    try:
        sheet_names = list_sheets(file_bytes)
        target = None
        for s in sheet_names:
            if str(s).strip().lower() == 'sheet1':
                target = s
                break
        if target is None:
            return {}
        raw = read_excel_raw(file_bytes, target)
        header_row = None
        for i in range(min(8, len(raw))):
            row_vals = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
            if 'model' in row_vals and 'erp' in row_vals:
                header_row = i
                break
        if header_row is None:
            return {}
        merged = merge_two_row_header(raw, header_row)
        df = raw.iloc[header_row + 2:].copy()
        df.columns = merged
        df.columns = [str(c).strip() for c in df.columns]
        df = df.reset_index(drop=True)
        if 'Note' not in df.columns:
            for c in df.columns:
                if str(c).lower().strip() == 'note':
                    df = df.rename(columns={c: 'Note'})
                    break
        if 'ERP' not in df.columns or 'Note' not in df.columns:
            return {}
        df['ERP'] = df['ERP'].astype(str).str.strip()
        df = df[df['ERP'].str.startswith(('013', '018'))]
        return dict(zip(df['ERP'], df['Note'].fillna('')))
    except Exception:
        return {}


def load_shipment_rev(file_bytes):
    sheet_names = list_sheets(file_bytes)
    target = None
    for s in sheet_names:
        cl = str(s).strip().lower()
        if 'shipment' in cl and 'rev' in cl:
            target = s
            break
    if target is None:
        for s in sheet_names:
            if str(s).strip().lower() == 'shipment':
                target = s
                break
    if target is None:
        target = sheet_names[0]

    st.success(f"✅ 사용 시트: '{target}'")
    raw = read_excel_raw(file_bytes, target)

    header_row = None
    for i in range(min(8, len(raw))):
        row_vals = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
        if 'cus' in row_vals:
            header_row = i
            break
    if header_row is None:
        header_row = 2

    merged = merge_two_row_header(raw, header_row)
    df = raw.iloc[header_row + 2:].copy()
    df.columns = merged
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)

    rename_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        cl_clean = cl.replace(' ', '').replace('(', '').replace(')', '')
        if cl == 'cus' or 'customer' in cl:
            rename_map[c] = 'Cus'
        elif 'cut off cargo' in cl:
            rename_map[c] = 'Cut off Cargo'
        elif cl == 'hq' or 'hq request' in cl or 'hq.request' in cl:
            rename_map[c] = 'HQ Request'
        elif cl == 'model':
            rename_map[c] = 'model'
        elif cl == 'inch':
            rename_map[c] = 'Inch'
        elif cl == 'bncode' or ('bn' in cl and 'code' in cl):
            rename_map[c] = 'code'
        elif '3in1code' in cl_clean or cl == 'erp':
            rename_map[c] = 'ERP'
        elif 'po remain' in cl or 'po.remain' in cl:
            rename_map[c] = 'PO'
        elif 'ttl ship' in cl or 'ttl.ship' in cl:
            rename_map[c] = '_TTLShip'
        elif 'ttl plan' in cl or 'ttl.plan' in cl:
            rename_map[c] = '예상계획'
        elif 'o/stock' in cl or 'o.stock' in cl:
            rename_map[c] = '현재재고'
    df = df.rename(columns=rename_map)

    plan_date_cols = []
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'plan' in cl and any(ch.isdigit() for ch in cl):
            if 'ttl' in cl or 'actual' in cl:
                continue
            if cl == '예상계획'.lower():
                continue
            plan_date_cols.append(c)

    if 'ERP' not in df.columns:
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(30).tolist()
            erp_count = sum(1 for v in sample 
                            if (v.startswith('013') or v.startswith('018')) 
                            and len(v) >= 10)
            if erp_count >= 3:
                df = df.rename(columns={col: 'ERP'})
                break

    if 'ERP' not in df.columns:
        st.error("❌ ERP 컬럼 매핑 실패")
        return pd.DataFrame(), []

    df['ERP'] = df['ERP'].astype(str).str.strip()
    df = df[df['ERP'].str.startswith(('013', '018'))].copy()
    df = df.reset_index(drop=True)

    with st.expander("🔍 컬럼 매핑 진단", expanded=False):
        st.caption(f"전체 컬럼: {list(df.columns)}")
        st.caption(f"일자별 Plan: {plan_date_cols}")

    return df, plan_date_cols


# ════════════════════════════════════════════════════════════
# 분석 v30.1
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db, plan_date_cols, note_dict):
    prod_db = prod_db.copy()
    prod_db['TRAN_WORK_DATE'] = pd.to_datetime(prod_db['TRAN_WORK_DATE'], errors='coerce')
    today = prod_db['TRAN_WORK_DATE'].max()
    
    if pd.isna(today):
        st.error("❌ 생산실적 날짜 파싱 실패")
        return pd.DataFrame(), ''

    today_str = today.strftime('%m/%d')
    today_norm = today.normalize()
    st.info(f"📅 기준일: **{today.strftime('%Y-%m-%d')}** | 실적 마지막 일자")

    prod_db['ERP'] = prod_db['FINAL_MAT_ID'].astype(str).str.strip()
    prod_db['TYPE'] = prod_db['ERP'].apply(classify)
    valid = prod_db[
        ((prod_db['TYPE']=='PD') & (prod_db['OPER_DESC']=='P-ATE')) |
        ((prod_db['TYPE']=='3IN1') & (prod_db['OPER_DESC']=='ASSY'))
    ].copy()

    daily = valid.groupby(['ERP', valid['TRAN_WORK_DATE'].dt.normalize()])['QTY'].sum().reset_index()
    daily.columns = ['ERP', 'DATE', 'QTY']
    daily_dict_erp = {(r['ERP'], r['DATE']): r['QTY'] for _, r in daily.iterrows()}
    erp_total_actual = valid.groupby('ERP')['QTY'].sum().to_dict()

    m = ship_db.copy()
    m['MODEL_TYPE'] = m['ERP'].apply(classify)
    m['Note'] = m['ERP'].map(note_dict).fillna('')

    for col in ['PO', '예상계획', '현재재고', '_TTLShip'] + plan_date_cols:
        if col in m.columns:
            m[col] = pd.to_numeric(m[col], errors='coerce').fillna(0)

    if 'PO' not in m.columns and '_TTLShip' in m.columns:
        m['PO'] = m['_TTLShip']
    elif 'PO' in m.columns and '_TTLShip' in m.columns:
        m['PO'] = m.apply(lambda r: r['PO'] if r['PO'] > 0 else r['_TTLShip'], axis=1)

    if '현재재고' not in m.columns:
        m['현재재고'] = 0
    if '예상계획' not in m.columns:
        m['예상계획'] = 0
    m['현재재고'] = m['현재재고'].astype(int)
    m['예상계획'] = m['예상계획'].astype(int)

    if 'code' not in m.columns:
        m['code'] = m['ERP']
    m['code'] = m['code'].astype(str).str.strip()

    code_stock = m.groupby('code')['현재재고'].max().to_dict()
    code_plan = m.groupby('code')['예상계획'].max().to_dict()
    code_erp_map = m.groupby('code')['ERP'].apply(lambda x: list(set(x))).to_dict()
    
    code_total_actual = {}
    for code, erp_list in code_erp_map.items():
        code_total_actual[code] = sum(erp_total_actual.get(erp, 0) for erp in erp_list)
    
    m['현재재고'] = m['code'].map(code_stock).fillna(0).astype(int)
    m['예상계획'] = m['code'].map(code_plan).fillna(0).astype(int)
    m['_현재실적'] = m['code'].map(code_total_actual).fillna(0).astype(int)
    m[f'현재실적({today_str})'] = m['_현재실적']

    plan_date_map = {col: parse_date_from_col(col) for col in plan_date_cols}
    valid_plan_cols = {col: dt for col, dt in plan_date_map.items() if dt is not None}

    # code별 일자별 계획 사전
    code_daily_plan = {}
    for code in m['code'].unique():
        first_row = m[m['code'] == code].iloc[0]
        for col, dt in valid_plan_cols.items():
            plan_val = pd.to_numeric(first_row.get(col, 0), errors='coerce')
            if pd.isna(plan_val):
                plan_val = 0
            code_daily_plan[(code, dt.normalize())] = int(plan_val)

    # 미래 계획
    code_future_plan = {}
    for code in m['code'].unique():
        fp = sum(
            qty for (c, d), qty in code_daily_plan.items()
            if c == code and d > today_norm
        )
        code_future_plan[code] = int(fp)

    m['_미래계획'] = m['code'].map(code_future_plan).fillna(0).astype(int)
    m['_현재기반_생산'] = m['_현재실적'] + m['_미래계획']
    m['실적차이'] = m['_현재기반_생산'] - m['예상계획']

    # ⭐ Cutoff 시점 재고 계산 헬퍼
    def calc_cutoff_stock(code, cutoff_dt):
       """Cutoff 시점까지 누적 생산 + 현재재고 (작성 기준)"""
       stock = code_stock.get(code, 0)
        
       if pd.isna(cutoff_dt):
           # 전체 (모든 일자 합산)
           cutoff_norm = pd.Timestamp(2030, 1, 1)
       else:
           cutoff_norm = cutoff_dt.normalize()
        
        erp_list = code_erp_map.get(code, [])
        production = 0
        
        # code의 모든 일자별 계획 순회
       for (c, d), plan_qty in code_daily_plan.items():
           if c != code:
               continue
           if d > cutoff_norm:
               continue  # cutoff 이후 제외
            
           if d <= today_norm:
               # 과거/오늘: 실적 사용 (해당 일자의 ERP별 실적 합)
               for erp in erp_list:
                   production += daily_dict_erp.get((erp, d), 0)
           else:
               # 미래: 계획 사용
               production += plan_qty
        
       return stock + production

    # Cut off 정렬
    m['_cutoff_dt'] = pd.to_datetime(m['Cut off Cargo'], errors='coerce')
    m = m.sort_values(['code', '_cutoff_dt']).reset_index(drop=True)

    # 시뮬레이션 1: 계획 기준
    m['예상제품재고_계획'] = 0
    m['_부족1'] = 0
    
    for code in m['code'].unique():
        code_idx = m[m['code'] == code].index.tolist()
        total = len(code_idx)
        stock = code_stock.get(code, 0)
        plan = code_plan.get(code, 0)
        available = stock + plan
        
        for i, idx in enumerate(code_idx):
            po = int(m.loc[idx, 'PO'])
            m.at[idx, '순서'] = f"{i+1}/{total}"
            
            if available >= po:
                available -= po
                m.at[idx, '예상제품재고_계획'] = available
                m.at[idx, '_부족1'] = 0
            elif available > 0:
                m.at[idx, '예상제품재고_계획'] = -(po - available)
                m.at[idx, '_부족1'] = po - available
                available = 0
            else:
                m.at[idx, '예상제품재고_계획'] = -po
                m.at[idx, '_부족1'] = po

    # 시뮬레이션 2: 실적 반영 + Cutoff 시점 재고
    m['예상제품재고_실적'] = 0
    m['Cutoff시점재고'] = 0
    m['_부족2'] = 0
    
    for code in m['code'].unique():
        code_idx = m[m['code'] == code].index.tolist()
        stock = code_stock.get(code, 0)
        current_production = code_total_actual.get(code, 0) + code_future_plan.get(code, 0)
        available = stock + current_production
        
        for i, idx in enumerate(code_idx):
            po = int(m.loc[idx, 'PO'])
            cutoff_dt = m.loc[idx, '_cutoff_dt']
            
            # ⭐ Cutoff 시점 재고 (해당 cutoff 날짜까지 누적 생산 + 재고)
            m.at[idx, 'Cutoff시점재고'] = calc_cutoff_stock(code, cutoff_dt)
            
            # 예상제품재고 (Cut off FIFO 누적 차감)
            if available >= po:
                available -= po
                m.at[idx, '예상제품재고_실적'] = available
                m.at[idx, '_부족2'] = 0
            elif available > 0:
                m.at[idx, '예상제품재고_실적'] = -(po - available)
                m.at[idx, '_부족2'] = po - available
                available = 0
            else:
                m.at[idx, '예상제품재고_실적'] = -po
                m.at[idx, '_부족2'] = po

    # 알람
    def get_alert(shortage, gap=0):
        if shortage >= 5000:
            return "🔴 출하불가"
        elif shortage > 0:
            return "🟠 부족"
        elif gap < -1000:
            return "🟡 차질"
        else:
            return "✅ 정상"

    m['알람_계획'] = m['_부족1'].apply(lambda x: get_alert(x))
    m['알람_실적'] = m.apply(lambda r: get_alert(r['_부족2'], r['실적차이']), axis=1)

    # 정수화
    for col in ['PO', '예상계획', '현재재고', '실적차이', 
                '예상제품재고_계획', '예상제품재고_실적', 'Cutoff시점재고',
                f'현재실적({today_str})']:
        if col in m.columns:
            m[col] = m[col].astype(int)

    m = m.drop(columns=['_부족1','_부족2','_현재실적','_cutoff_dt','_현재기반_생산','_미래계획'], errors='ignore')

    return m, today_str


# ════════════════════════════════════════════════════════════
# HTML 테이블 (components.html + JS 정렬)
# ════════════════════════════════════════════════════════════
def render_html_table(df, height=500, table_id="t1"):
    numeric_cols = set()
    for c in df.columns:
        if df[c].dtype.kind in 'iuf':
            numeric_cols.add(c)

    def fmt(val, is_num):
        if pd.isna(val) or val == '':
            return '-'
        if is_num:
            try:
                n = int(float(val))
                if n < 0:
                    return f'<span style="color:#DC2626;font-weight:700;">{n:,}</span>'
                if n == 0:
                    return '<span style="color:#9CA3AF;">0</span>'
                return f'{n:,}'
            except Exception:
                pass
        return str(val)

    def bg(row):
        alerts = []
        for col in ['알람_실적', '알람_계획', '알람']:
            if col in row.index:
                alerts.append(str(row.get(col, '')))
        combined = ' '.join(alerts)
        if '출하불가' in combined: return '#FEE2E2'
        if '부족' in combined: return '#FFEDD5'
        if '차질' in combined: return '#FEF3C7'
        return '#FFFFFF'

    parts = [f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
body {{ margin:0; padding:0; font-family:-apple-system,sans-serif; font-size:12px; }}
#{table_id}_wrap {{ max-height:{height}px; overflow:auto; border:1px solid #E5E7EB; border-radius:6px; }}
#{table_id} {{ border-collapse:collapse; width:100%; }}
#{table_id} th {{ 
    cursor:pointer; user-select:none; 
    padding:8px 6px; text-align:center; 
    border:1px solid #4B5563; font-weight:700; 
    white-space:nowrap; background-color:#374151; color:white;
    position:sticky; top:0; z-index:10;
}}
#{table_id} th:hover {{ background-color:#4B5563 !important; }}
#{table_id} th.sort-asc::after {{ content:" ▲"; color:#FCD34D; }}
#{table_id} th.sort-desc::after {{ content:" ▼"; color:#FCD34D; }}
#{table_id} td {{ 
    padding:6px; border-bottom:1px solid #E5E7EB; 
    white-space:nowrap; 
}}
</style></head><body>
<div id="{table_id}_wrap">
<table id="{table_id}">
<thead><tr>''']
    
    for i, col in enumerate(df.columns):
        is_num = col in numeric_cols
        data_type = 'number' if is_num else 'string'
        parts.append(f'<th data-col="{i}" data-type="{data_type}" onclick="sortTable({i},\'{data_type}\')">{col}</th>')
    parts.append('</tr></thead><tbody>')

    for _, row in df.iterrows():
        row_bg = bg(row)
        parts.append(f'<tr style="background-color:{row_bg};">')
        for col in df.columns:
            is_num = col in numeric_cols
            val = row[col]
            cell = fmt(val, is_num)
            align = 'right' if is_num else 'center'
            sort_val = ''
            if is_num and not pd.isna(val):
                try:
                    sort_val = f' data-sort="{int(float(val))}"'
                except Exception:
                    pass
            parts.append(f'<td{sort_val} style="text-align:{align};">{cell}</td>')
        parts.append('</tr>')
    
    parts.append('''</tbody></table></div>
<script>
var sortStates = {};
function sortTable(colIdx, type) {
    var table = document.querySelector("table");
    var tbody = table.querySelector("tbody");
    var rows = Array.from(tbody.querySelectorAll("tr"));
    var ths = table.querySelectorAll("th");
    
    var currentAsc = sortStates[colIdx] === 'asc';
    var newDir = currentAsc ? 'desc' : 'asc';
    sortStates = {};
    sortStates[colIdx] = newDir;
    
    ths.forEach(function(th) { th.classList.remove('sort-asc','sort-desc'); });
    ths[colIdx].classList.add('sort-' + newDir);
    
    rows.sort(function(a, b) {
        var ca = a.cells[colIdx];
        var cb = b.cells[colIdx];
        var va, vb;
        if (type === 'number') {
            va = parseFloat(ca.getAttribute('data-sort') || ca.textContent.replace(/[,\\s]/g,'')) || 0;
            vb = parseFloat(cb.getAttribute('data-sort') || cb.textContent.replace(/[,\\s]/g,'')) || 0;
        } else {
            va = ca.textContent.trim();
            vb = cb.textContent.trim();
        }
        if (va < vb) return newDir === 'asc' ? -1 : 1;
        if (va > vb) return newDir === 'asc' ? 1 : -1;
        return 0;
    });
    rows.forEach(function(r) { tbody.appendChild(r); });
}
</script></body></html>''')
    
    html = ''.join(parts)
    components.html(html, height=height + 50, scrolling=False)


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람")
    st.caption("📌 계획 vs 실적 비교 | 컬럼 클릭으로 정렬")

    ship_db = st.session_state.get('ship_db', pd.DataFrame())
    plan_cols = st.session_state.get('plan_date_cols', [])
    note_dict = st.session_state.get('note_dict', {})
    prod_db = st.session_state.get('prod_db', pd.DataFrame())
    ship_t = st.session_state.get('ship_updated', '-')
    prod_t = st.session_state.get('prod_updated', '-')

    s1, s2 = st.columns(2)
    if not ship_db.empty:
        s1.success(f"📁 출하계획: **{len(ship_db)}건** | Note {len(note_dict)}건 | {ship_t}")
    else:
        s1.warning("📁 출하계획: 없음")
    if not prod_db.empty:
        s2.success(f"📊 생산실적: **{len(prod_db)}건** | {prod_t}")
    else:
        s2.warning("📊 생산실적: 없음")

    st.markdown("##### 📤 파일 업로드")
    files = st.file_uploader(" ", type=["xlsx","csv"], accept_multiple_files=True,
                              key="ship_up_v30", label_visibility="collapsed")

    if files:
        if st.button("🚀 저장 및 분석", type="primary", use_container_width=True, key="apply_v30"):
            with st.spinner("처리 중..."):
                for f in files:
                    fname = f.name.lower()
                    f.seek(0)
                    file_bytes = f.read()
                    if fname.endswith('.xlsx'):
                        df, p_cols = load_shipment_rev(file_bytes)
                        notes = load_sheet1_notes(file_bytes)
                        if not df.empty:
                            st.session_state['ship_db'] = df
                            st.session_state['plan_date_cols'] = p_cols
                            st.session_state['note_dict'] = notes
                            st.session_state['ship_updated'] = datetime.now().strftime('%m-%d %H:%M')
                            st.success(f"✅ 출하계획 ({len(df)}건) + Note ({len(notes)}건)")
                    elif fname.endswith('.csv'):
                        try:
                            df = read_csv_cached(file_bytes)
                            st.session_state['prod_db'] = df
                            st.session_state['prod_updated'] = datetime.now().strftime('%m-%d %H:%M')
                            st.success(f"✅ 생산실적 ({len(df)}건)")
                        except Exception as e:
                            st.error(f"❌ {f.name}: {e}")
            st.rerun()

    st.markdown("---")

    if ship_db.empty or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    with st.spinner("분석 중..."):
        try:
            m, today_str = analyze(ship_db, prod_db, plan_cols, note_dict)
        except Exception as e:
            st.error(f"❌ 분석 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
            return

    if m.empty:
        return

    actual_col = f'현재실적({today_str})'

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 통합 비교 표
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.markdown("---")
    st.markdown("### 🚨 즉시 조치 필요 (계획 또는 실적 기준 문제)")
    st.caption("⚠️ 두 시뮬레이션 중 하나라도 문제 발생한 행만 표시 | 컬럼 클릭 정렬")
    
    problem_alerts = ['🔴 출하불가', '🟠 부족', '🟡 차질']
    urgent = m[(m['알람_계획'].isin(problem_alerts)) | (m['알람_실적'].isin(problem_alerts))].copy()
    
    if not urgent.empty:
        k1, k2, k3 = st.columns(3)
        k1.metric("문제 행", f"{len(urgent)}건")
        k2.metric("🔴 출하불가 (실적)", int((urgent['알람_실적']=='🔴 출하불가').sum()))
        k3.metric("🟠 부족 (실적)", int((urgent['알람_실적']=='🟠 부족').sum()))
        
        compare_cols = ['알람_계획','알람_실적','순서','Cut off Cargo','Cus','model','code','ERP','MODEL_TYPE',
                        'PO','현재재고','예상계획',actual_col,'실적차이',
                        '예상제품재고_계획','예상제품재고_실적','Cutoff시점재고','Note']
        compare_cols = [c for c in compare_cols if c in urgent.columns]
        render_html_table(urgent[compare_cols], height=500, table_id="compare")
    else:
        st.success("✅ 모든 모델 정상")

    # 필터
    st.markdown("---")
    st.markdown("### 🎛️ 필터")
    f1, f2, f3 = st.columns(3)
    a_sel = f1.multiselect("알람 (실적 기준)", ['🔴 출하불가','🟠 부족','🟡 차질','✅ 정상'], key="a_v30")
    t_sel = f2.multiselect("모델", ['PD','3IN1','OTHER'], key="t_v30")
    cus_opt = sorted(m['Cus'].dropna().unique()) if 'Cus' in m.columns else []
    c_sel = f3.multiselect("거래선", cus_opt, key="c_v30") if cus_opt else []

    def apply_filter(df, alert_col='알람_실적'):
        v = df.copy()
        if a_sel: v = v[v[alert_col].isin(a_sel)]
        if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
        if c_sel: v = v[v['Cus'].isin(c_sel)]
        return v

    # 테이블 1
    st.markdown("---")
    st.markdown("### 📋 1) 이번 주 Cut off 예상 (계획 100%)")
    st.caption("예상제품재고 = 현재재고 + 예상계획 - 누적 PO")
    
    t1 = m.rename(columns={'알람_계획':'알람', '예상제품재고_계획':'예상제품재고'}).copy()
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("전체", f"{len(t1)}건")
    k2.metric("🔴 출하불가", int((t1['알람']=='🔴 출하불가').sum()))
    k3.metric("🟠 부족", int((t1['알람']=='🟠 부족').sum()))
    k4.metric("🟡 차질", int((t1['알람']=='🟡 차질').sum()))
    k5.metric("✅ 정상", int((t1['알람']=='✅ 정상').sum()))
    
    show1 = ['알람','순서','Cut off Cargo','Cus','model','code','ERP','MODEL_TYPE',
             'PO','현재재고','예상계획','예상제품재고','Note']
    show1 = [c for c in show1 if c in t1.columns]
    v1 = apply_filter(t1, '알람')
    if not v1.empty:
        render_html_table(v1[show1], height=400, table_id="t1")

    # 테이블 2
    st.markdown("---")
    st.markdown("### 📋 2) 현재 실적 반영 시 Cut off 예상")
    st.caption(f"예상제품재고 = 현재재고 + ({actual_col} + 미래 계획) - 누적 PO | Cutoff시점재고 = Cut off 날짜까지 재고+생산")
    
    t2 = m.rename(columns={'알람_실적':'알람', '예상제품재고_실적':'예상제품재고'}).copy()
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("전체", f"{len(t2)}건")
    k2.metric("🔴 출하불가", int((t2['알람']=='🔴 출하불가').sum()))
    k3.metric("🟠 부족", int((t2['알람']=='🟠 부족').sum()))
    k4.metric("🟡 차질", int((t2['알람']=='🟡 차질').sum()))
    k5.metric("✅ 정상", int((t2['알람']=='✅ 정상').sum()))
    
    show2 = ['알람','순서','Cut off Cargo','Cus','model','code','ERP','MODEL_TYPE',
             'PO','현재재고','예상계획',actual_col,'실적차이','예상제품재고','Cutoff시점재고','Note']
    show2 = [c for c in show2 if c in t2.columns]
    v2 = apply_filter(t2, '알람')
    if not v2.empty:
        render_html_table(v2[show2], height=500, table_id="t2")
        csv = v2[show2].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 CSV 다운로드", csv,
                           f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           "text/csv", key="dl_v30")

    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화", use_container_width=True, key="rs_v30"):
        for k in ['ship_db','ship_updated','plan_date_cols','note_dict']:
            st.session_state.pop(k, None)
        st.cache_data.clear()
        st.rerun()
    if r2.button("🗑️ 생산실적 초기화", use_container_width=True, key="rp_v30"):
        for k in ['prod_db','prod_updated']:
            st.session_state.pop(k, None)
        st.cache_data.clear()
        st.rerun()
