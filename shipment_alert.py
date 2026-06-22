"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v30.8
- ORDER_MAD_ID 컬럼명 수정
- Cutoff시점재고 = 재고 + cutoff까지 실적 - cutoff까지 누적PO
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
    if x.startswith('01'):  return 'PD'
    return 'OTHER'


def parse_date_from_col(col_name, year=2026):
    try:
        s = str(col_name).lower()
        s = re.sub(r'(plan|actual|cut\s*off|cargo)[\.\s_&]*', '', s).strip()
        nums = re.findall(r'\d+', s)
        if len(nums) >= 2:
            m2, d2 = int(nums[0]), int(nums[1])
            if 1 <= m2 <= 12 and 1 <= d2 <= 31:
                return pd.Timestamp(year, m2, d2)
        elif len(nums) == 1:
            d2 = int(nums[0])
            if 1 <= d2 <= 31:
                return pd.Timestamp(year, 6, d2)
        return None
    except Exception:
        return None


def normalize_cutoff(cdt):
    """Cut off Cargo 날짜 → 2026년 기준 Timestamp"""
    if pd.isna(cdt):
        return pd.Timestamp(2030, 1, 1)
    ts = pd.Timestamp(cdt).normalize()
    if ts.year < 2026:
        ts = ts.replace(year=2026)
    return ts


def merge_two_row_header(raw, header_row):
    main = raw.iloc[header_row].values
    sub  = (raw.iloc[header_row + 1].values
            if header_row + 1 < len(raw) else [None] * len(main))
    merged = []
    for mh, sh in zip(main, sub):
        ms = str(mh).strip() if pd.notna(mh) else ''
        ss = str(sh).strip() if pd.notna(sh) else ''
        if ms in ('nan', 'None'): ms = ''
        if ss in ('nan', 'None'): ss = ''
        if ms and ss:   merged.append(f"{ms}.{ss}")
        elif ms:        merged.append(ms)
        elif ss:        merged.append(ss)
        else:           merged.append(f"col_{len(merged)}")
    return merged


# 저장 (업로드 버튼 클릭 시)
github_save_xlsx(file_bytes, "data/shipment.xlsx")
github_save_csv(prod_df, "data/production.csv")

# 로드 (앱 시작 시)
if 'ship_db' not in st.session_state:
    df = github_load_xlsx("data/shipment.xlsx")
    if not df.empty:
        st.session_state['ship_db'] = df
        
if 'prod_db' not in st.session_state:
    df = github_load_csv("data/production.csv")
    if not df.empty:
        st.session_state['prod_db'] = df



def load_sheet1_notes(file_bytes):
    try:
        sheet_names = list_sheets(file_bytes)
        target = next((s for s in sheet_names
                       if str(s).strip().lower() == 'sheet1'), None)
        if target is None:
            return {}
        raw = read_excel_raw(file_bytes, target)
        header_row = None
        for i in range(min(8, len(raw))):
            rv = [str(v).lower().strip()
                  for v in raw.iloc[i].values if pd.notna(v)]
            if 'model' in rv and 'erp' in rv:
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
        target = next((s for s in sheet_names
                       if str(s).strip().lower() == 'shipment'), None)
    if target is None:
        target = sheet_names[0]

    st.success(f"✅ 사용 시트: '{target}'")
    raw = read_excel_raw(file_bytes, target)

    header_row = None
    for i in range(min(8, len(raw))):
        rv = [str(v).lower().strip()
              for v in raw.iloc[i].values if pd.notna(v)]
        if 'cus' in rv:
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
        cl  = str(c).lower().strip()
        clc = cl.replace(' ', '').replace('(', '').replace(')', '')
        if cl == 'cus' or 'customer' in cl:            rename_map[c] = 'Cus'
        elif 'cut off cargo' in cl:                    rename_map[c] = 'Cut off Cargo'
        elif cl == 'hq' or 'hq request' in cl \
                or 'hq.request' in cl:                 rename_map[c] = 'HQ Request'
        elif cl == 'model':                            rename_map[c] = 'model'
        elif cl == 'inch':                             rename_map[c] = 'Inch'
        elif cl == 'bncode' or ('bn' in cl and 'code' in cl):
                                                       rename_map[c] = 'code'
        elif '3in1code' in clc or cl == 'erp':         rename_map[c] = 'ERP'
        elif 'po remain' in cl or 'po.remain' in cl:   rename_map[c] = 'PO'
        elif 'ttl ship'  in cl or 'ttl.ship'  in cl:   rename_map[c] = '_TTLShip'
        elif 'ttl plan'  in cl or 'ttl.plan'  in cl:   rename_map[c] = '예상계획'
        elif 'o/stock'   in cl or 'o.stock'   in cl:   rename_map[c] = '현재재고'
    df = df.rename(columns=rename_map)

    plan_date_cols = []
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'plan' in cl and any(ch.isdigit() for ch in cl):
            if 'ttl' in cl or 'actual' in cl:
                continue
            plan_date_cols.append(c)

    if 'ERP' not in df.columns:
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(30).tolist()
            if sum(1 for v in sample
                   if (v.startswith('013') or v.startswith('018'))
                   and len(v) >= 10) >= 3:
                df = df.rename(columns={col: 'ERP'})
                break

    if 'ERP' not in df.columns:
        st.error("❌ ERP 컬럼 매핑 실패")
        return pd.DataFrame(), []

    df['ERP'] = df['ERP'].astype(str).str.strip()
    df = df[df['ERP'].str.startswith(('013', '018'))].copy().reset_index(drop=True)

    with st.expander("🔍 컬럼 매핑 진단", expanded=False):
        st.caption(f"전체 컬럼: {list(df.columns)}")
        st.caption(f"일자별 Plan: {plan_date_cols}")

    return df, plan_date_cols


# ════════════════════════════════════════════════════════════
# 분석 v30.8
# ════════════════════════════════════════════════════════════
def analyze(ship_db, prod_db, plan_date_cols, note_dict):
    st.warning("🚀 **v30.8 실행 중**")

    prod_db = prod_db.copy()
    prod_db['TRAN_WORK_DATE'] = pd.to_datetime(
        prod_db['TRAN_WORK_DATE'], errors='coerce')
    today = prod_db['TRAN_WORK_DATE'].max()
    if pd.isna(today):
        st.error("❌ 생산실적 날짜 파싱 실패")
        return pd.DataFrame(), ''

    today_str  = today.strftime('%m/%d')
    today_norm = today.normalize()
    st.info(f"📅 기준일: **{today.strftime('%Y-%m-%d')}** | 실적 마지막 일자")

    # ── MES 실적 ──────────────────────────────────────────
    # ⭐ ORDER_MAD_ID 컬럼 사용 (FINAL_MAT_ID → ORDER_MAD_ID)
    erp_col = None
    for candidate in ['ORDER_MAD_ID', 'FINAL_MAT_ID', 'MAT_ID', 'ERP']:
        if candidate in prod_db.columns:
            erp_col = candidate
            break
    if erp_col is None:
        st.error(f"❌ ERP 컬럼 없음. 컬럼 목록: {list(prod_db.columns)}")
        return pd.DataFrame(), ''

    st.info(f"📊 실적 ERP 컬럼: **{erp_col}**")
    prod_db['ERP'] = prod_db[erp_col].astype(str).str.strip()
    prod_db['TYPE'] = prod_db['ERP'].apply(classify)
    valid = prod_db[
        ((prod_db['TYPE'] == 'PD')   & (prod_db['OPER_DESC'] == 'P-ATE')) |
        ((prod_db['TYPE'] == '3IN1') & (prod_db['OPER_DESC'] == 'ASSY'))
    ].copy()

    daily = (valid
             .groupby(['ERP', valid['TRAN_WORK_DATE'].dt.normalize()])['QTY']
             .sum().reset_index())
    daily.columns = ['ERP', 'DATE', 'QTY']
    daily_dict_erp   = {(r.ERP, r.DATE): r.QTY for r in daily.itertuples()}
    erp_total_actual = valid.groupby('ERP')['QTY'].sum().to_dict()

    # ── Shipment Rev ──────────────────────────────────────
    mdf = ship_db.copy()
    mdf['MODEL_TYPE'] = mdf['ERP'].apply(classify)
    mdf['Note']       = mdf['ERP'].map(note_dict).fillna('')

    for col in ['PO', '예상계획', '현재재고', '_TTLShip'] + plan_date_cols:
        if col in mdf.columns:
            mdf[col] = pd.to_numeric(mdf[col], errors='coerce').fillna(0)

    if 'PO' not in mdf.columns and '_TTLShip' in mdf.columns:
        mdf['PO'] = mdf['_TTLShip']
    elif 'PO' in mdf.columns and '_TTLShip' in mdf.columns:
        mdf['PO'] = mdf.apply(
            lambda r: r['PO'] if r['PO'] > 0 else r['_TTLShip'], axis=1)

    mdf['현재재고'] = mdf.get('현재재고', pd.Series(0, index=mdf.index)).astype(int)
    mdf['예상계획'] = mdf.get('예상계획', pd.Series(0, index=mdf.index)).astype(int)

    if 'code' not in mdf.columns:
        mdf['code'] = mdf['ERP']
    mdf['code'] = mdf['code'].astype(str).str.strip()

    # ── code 단위 집계 ────────────────────────────────────
    code_stock   = mdf.groupby('code')['현재재고'].max().to_dict()
    code_plan    = mdf.groupby('code')['예상계획'].max().to_dict()
    code_erp_map = (mdf.groupby('code')['ERP']
                       .apply(lambda x: list(set(x))).to_dict())
    code_total_actual = {
        ck: sum(erp_total_actual.get(e, 0) for e in ev)
        for ck, ev in code_erp_map.items()
    }

    mdf['현재재고'] = mdf['code'].map(code_stock).fillna(0).astype(int)
    mdf['예상계획'] = mdf['code'].map(code_plan).fillna(0).astype(int)
    mdf['_실적']   = mdf['code'].map(code_total_actual).fillna(0).astype(int)
    mdf[f'현재실적({today_str})'] = mdf['_실적']

    # ── 일자별 계획 ───────────────────────────────────────
    plan_date_map = {c: parse_date_from_col(c) for c in plan_date_cols}
    valid_plan    = {c: dt for c, dt in plan_date_map.items() if dt is not None}

    code_daily_plan = {}
    for ck in mdf['code'].unique():
        fr = mdf[mdf['code'] == ck].iloc[0]
        dp = {}
        for col, dt in valid_plan.items():
            v = pd.to_numeric(fr.get(col, 0), errors='coerce')
            dp[dt.normalize()] = int(v) if not pd.isna(v) else 0
        code_daily_plan[ck] = dp

    code_future_plan = {
        ck: sum(q for d, q in dp.items() if d > today_norm)
        for ck, dp in code_daily_plan.items()
    }

    mdf['_미래계획'] = mdf['code'].map(code_future_plan).fillna(0).astype(int)
    mdf['실적차이']  = (mdf['_실적'] + mdf['_미래계획']) - mdf['예상계획']

    # ── Cut-off 정렬 ──────────────────────────────────────
    mdf['_cdt'] = pd.to_datetime(mdf['Cut off Cargo'], errors='coerce')
    mdf = mdf.sort_values(['code', '_cdt']).reset_index(drop=True)

    # ── 시뮬레이션 1: 계획 100% ──────────────────────────
    mdf['예상제품재고_계획'] = 0.0
    mdf['_부족1']           = 0.0

    for ck in mdf['code'].unique():
        idxs  = mdf[mdf['code'] == ck].index.tolist()
        avail = float(code_stock.get(ck, 0) + code_plan.get(ck, 0))
        for i, idx in enumerate(idxs):
            po = float(mdf.loc[idx, 'PO'])
            mdf.loc[idx, '순서'] = f"{i+1}/{len(idxs)}"
            if avail >= po:
                avail -= po
                mdf.loc[idx, '예상제품재고_계획'] = avail
                mdf.loc[idx, '_부족1']           = 0.0
            elif avail > 0:
                mdf.loc[idx, '예상제품재고_계획'] = -(po - avail)
                mdf.loc[idx, '_부족1']           = po - avail
                avail = 0.0
            else:
                mdf.loc[idx, '예상제품재고_계획'] = -po
                mdf.loc[idx, '_부족1']           = po

    # ── 시뮬레이션 2: 실적 반영 ──────────────────────────
    mdf['예상제품재고_실적'] = 0.0
    mdf['Cutoff시점재고']   = 0.0
    mdf['_부족2']           = 0.0

    for ck in mdf['code'].unique():
        idxs       = mdf[mdf['code'] == ck].index.tolist()
        stk        = float(code_stock.get(ck, 0))
        avail      = stk + float(code_total_actual.get(ck, 0)) \
                         + float(code_future_plan.get(ck, 0))
        erp_set_lc = set(code_erp_map.get(ck, []))
        cumul_po   = 0.0  # ⭐ cutoff까지 누적 PO

        for idx in idxs:
            po  = float(mdf.loc[idx, 'PO'])
            cdt = mdf.loc[idx, '_cdt']

            # ⭐ cutoff 날짜 정규화 (년도 보정)
            cutoff_n = normalize_cutoff(cdt)

            # ⭐ cutoff까지 실적 합산
            actual_until = sum(
                qty for (ek, d), qty in daily_dict_erp.items()
                if ek in erp_set_lc and d <= cutoff_n
            )

            # ⭐ cutoff까지 누적 PO 차감
            cumul_po += po
            mdf.loc[idx, 'Cutoff시점재고'] = float(
                stk + actual_until - cumul_po)

            # 예상제품재고_실적 (전체 주간 기준)
            if avail >= po:
                avail -= po
                mdf.loc[idx, '예상제품재고_실적'] = avail
                mdf.loc[idx, '_부족2']           = 0.0
            elif avail > 0:
                mdf.loc[idx, '예상제품재고_실적'] = -(po - avail)
                mdf.loc[idx, '_부족2']           = po - avail
                avail = 0.0
            else:
                mdf.loc[idx, '예상제품재고_실적'] = -po
                mdf.loc[idx, '_부족2']           = po

    # ── 알람 ─────────────────────────────────────────────
    def get_alert(shortage, gap=0):
        if shortage >= 5000: return "🔴 출하불가"
        if shortage > 0:     return "🟠 부족"
        if gap < -1000:      return "🟡 차질"
        return "✅ 정상"

    mdf['알람_계획'] = mdf['_부족1'].apply(lambda x: get_alert(x))
    mdf['알람_실적'] = mdf.apply(
        lambda r: get_alert(r['_부족2'], r['실적차이']), axis=1)

    # ── 정수 변환 ─────────────────────────────────────────
    for col in ['PO', '예상계획', '현재재고', '실적차이',
                '예상제품재고_계획', '예상제품재고_실적',
                'Cutoff시점재고', f'현재실적({today_str})']:
        if col in mdf.columns:
            mdf[col] = mdf[col].astype(int)

    # ── 디버그 expander ───────────────────────────────────
    with st.expander("🐛 디버그 v30.8", expanded=False):
        rows = []
        for ck in sorted(mdf['code'].unique()):
            erp_list_v = code_erp_map.get(ck, [])
            dp_v       = code_daily_plan.get(ck, {})
            act_days   = {}
            for e in erp_list_v:
                for (ek, d), q in daily_dict_erp.items():
                    if ek == e:
                        act_days[d] = act_days.get(d, 0) + q
            act_s = ', '.join(
                f"{d.strftime('%m/%d')}:{q:,}"
                for d, q in sorted(act_days.items()))
            rows.append({
                'code': ck,
                'ERP': ', '.join(erp_list_v),
                '현재재고': code_stock.get(ck, 0),
                '실적합': code_total_actual.get(ck, 0),
                '일자별실적': act_s or '(없음)',
            })
        st.dataframe(pd.DataFrame(rows),
                     use_container_width=True, height=400)

    mdf = mdf.drop(columns=[
        '_부족1', '_부족2', '_실적', '_cdt', '_미래계획'
    ], errors='ignore')

    return mdf, today_str


# ════════════════════════════════════════════════════════════
# HTML 테이블
# ════════════════════════════════════════════════════════════
def render_html_table(df, height=500, table_id="t1"):
    num_cols = {c for c in df.columns if df[c].dtype.kind in 'iuf'}

    def fmt(val, is_num):
        if pd.isna(val) or val == '': return '-'
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

    def row_bg(row):
        txt = ' '.join(str(row.get(c, ''))
                       for c in ['알람_실적', '알람_계획', '알람']
                       if c in row.index)
        if '출하불가' in txt: return '#FEE2E2'
        if '부족'    in txt: return '#FFEDD5'
        if '차질'    in txt: return '#FEF3C7'
        return '#FFFFFF'

    parts = [f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
body{{margin:0;padding:0;font-family:-apple-system,sans-serif;font-size:12px;}}
#{table_id}_w{{max-height:{height}px;overflow:auto;
  border:1px solid #E5E7EB;border-radius:6px;}}
#{table_id}{{border-collapse:collapse;width:100%;}}
#{table_id} th{{cursor:pointer;user-select:none;padding:8px 6px;
  text-align:center;border:1px solid #4B5563;font-weight:700;
  white-space:nowrap;background:#374151;color:#fff;
  position:sticky;top:0;z-index:10;}}
#{table_id} th:hover{{background:#4B5563!important;}}
#{table_id} th.sa::after{{content:" ▲";color:#FCD34D;}}
#{table_id} th.sd::after{{content:" ▼";color:#FCD34D;}}
#{table_id} td{{padding:6px;border-bottom:1px solid #E5E7EB;white-space:nowrap;}}
</style></head><body>
<div id="{table_id}_w"><table id="{table_id}"><thead><tr>''']

    for i, col in enumerate(df.columns):
        dt = 'number' if col in num_cols else 'string'
        parts.append(f'<th onclick="sT({i},\'{dt}\')">{col}</th>')
    parts.append('</tr></thead><tbody>')

    for _, row in df.iterrows():
        bg = row_bg(row)
        parts.append(f'<tr style="background:{bg};">')
        for col in df.columns:
            is_n = col in num_cols
            val  = row[col]
            cell = fmt(val, is_n)
            aln  = 'right' if is_n else 'center'
            sv   = ''
            if is_n and not pd.isna(val):
                try: sv = f' data-sort="{int(float(val))}"'
                except Exception: pass
            parts.append(f'<td{sv} style="text-align:{aln};">{cell}</td>')
        parts.append('</tr>')

    parts.append(f'''</tbody></table></div>
<script>
var S_{table_id}={{}};
function sT(ci,tp){{
  var tb=document.getElementById("{table_id}");
  var bd=tb.querySelector("tbody");
  var rs=Array.from(bd.querySelectorAll("tr"));
  var hs=tb.querySelectorAll("th");
  var asc=S_{table_id}[ci]==='asc';
  var dir=asc?'desc':'asc';
  S_{table_id}={{}};S_{table_id}[ci]=dir;
  hs.forEach(h=>{{h.classList.remove('sa','sd');}});
  hs[ci].classList.add(dir==='asc'?'sa':'sd');
  rs.sort((a,b)=>{{
    var ca=a.cells[ci],cb=b.cells[ci],va,vb;
    if(tp==='number'){{
      va=parseFloat(ca.getAttribute('data-sort')||
         ca.textContent.replace(/[,\s]/g,''))||0;
      vb=parseFloat(cb.getAttribute('data-sort')||
         cb.textContent.replace(/[,\s]/g,''))||0;
    }}else{{va=ca.textContent.trim();vb=cb.textContent.trim();}}
    return dir==='asc'?(va<vb?-1:va>vb?1:0):(va>vb?-1:va<vb?1:0);
  }});
  rs.forEach(r=>bd.appendChild(r));
}}
</script></body></html>''')

    components.html(''.join(parts), height=height + 50, scrolling=False)


# ════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════
def render_shipment_alert_tab():
    st.markdown("#### 🚨 Shipment Cut-off 알람 **v30.8**")
    st.caption("📌 계획 vs 실적 비교 | 컬럼 클릭 정렬")

    ship_db   = st.session_state.get('ship_db',        pd.DataFrame())
    plan_cols = st.session_state.get('plan_date_cols', [])
    note_dict = st.session_state.get('note_dict',      {})
    prod_db   = st.session_state.get('prod_db',        pd.DataFrame())
    ship_t    = st.session_state.get('ship_updated',   '-')

    c1, c2 = st.columns(2)
    c1.success(f"📁 출하계획: **{len(ship_db)}건** | {ship_t}") \
        if not ship_db.empty else c1.warning("📁 출하계획: 없음")
    c2.success(f"📊 생산실적: **{len(prod_db)}건**") \
        if not prod_db.empty else c2.warning("📊 생산실적: 없음")

    st.markdown("##### 📤 파일 업로드")
    files = st.file_uploader(
        " ", type=["xlsx", "csv"], accept_multiple_files=True,
        key="ship_up_v30", label_visibility="collapsed")

    if files:
        if st.button("🚀 저장 및 분석", type="primary",
                     use_container_width=True, key="apply_v30"):
            with st.spinner("처리 중..."):
                for f in files:
                    f.seek(0); fb = f.read()
                    if f.name.lower().endswith('.xlsx'):
                        df, p = load_shipment_rev(fb)
                        ns    = load_sheet1_notes(fb)
                        if not df.empty:
                            st.session_state.update({
                                'ship_db': df, 'plan_date_cols': p,
                                'note_dict': ns,
                                'ship_updated': datetime.now().strftime('%m-%d %H:%M')
                            })
                            st.success(f"✅ 출하계획 {len(df)}건")
                    elif f.name.lower().endswith('.csv'):
                        try:
                            st.session_state['prod_db'] = read_csv_cached(fb)
                            st.session_state['prod_updated'] = \
                                datetime.now().strftime('%m-%d %H:%M')
                            st.success("✅ 생산실적 저장")
                        except Exception as e:
                            st.error(f"❌ {e}")
            st.rerun()

    st.markdown("---")
    if ship_db.empty or prod_db.empty:
        st.info("출하계획(.xlsx)과 생산실적(.csv) 모두 업로드 시 분석됩니다.")
        return

    with st.spinner("분석 중..."):
        try:
            mdf, today_str = analyze(ship_db, prod_db, plan_cols, note_dict)
        except Exception as e:
            st.error(f"❌ 분석 오류: {e}")
            import traceback; st.code(traceback.format_exc())
            return

    if mdf.empty:
        return

    actual_col     = f'현재실적({today_str})'
    problem_alerts = ['🔴 출하불가', '🟠 부족', '🟡 차질']

    # ── 통합 비교 표 ──────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🚨 즉시 조치 필요")
    st.caption("두 시뮬레이션 중 하나라도 문제 발생한 행 | 컬럼 클릭 정렬")

    urgent = mdf[
        mdf['알람_계획'].isin(problem_alerts) |
        mdf['알람_실적'].isin(problem_alerts)
    ].copy()

    if not urgent.empty:
        k1, k2, k3 = st.columns(3)
        k1.metric("문제 행", f"{len(urgent)}건")
        k2.metric("🔴 출하불가 (실적)",
                  int((urgent['알람_실적'] == '🔴 출하불가').sum()))
        k3.metric("🟠 부족 (실적)",
                  int((urgent['알람_실적'] == '🟠 부족').sum()))
        cc = ['알람_계획', '알람_실적', '순서', 'Cut off Cargo', 'Cus',
              'model', 'code', 'ERP', 'MODEL_TYPE', 'PO',
              '현재재고', '예상계획', actual_col, '실적차이',
              '예상제품재고_계획', '예상제품재고_실적', 'Cutoff시점재고', 'Note']
        cc = [c for c in cc if c in urgent.columns]
        render_html_table(urgent[cc], height=500, table_id="compare")
    else:
        st.success("✅ 모든 모델 정상")

    # ── 필터 ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🎛️ 필터")
    f1, f2, f3 = st.columns(3)
    a_sel = f1.multiselect("알람 (실적)",
        ['🔴 출하불가', '🟠 부족', '🟡 차질', '✅ 정상'], key="a_v30")
    t_sel = f2.multiselect("모델", ['PD', '3IN1', 'OTHER'], key="t_v30")
    cus_opt = sorted(mdf['Cus'].dropna().unique()) \
        if 'Cus' in mdf.columns else []
    c_sel = f3.multiselect("거래선", cus_opt, key="c_v30")

    def apply_filter(df, acol='알람_실적'):
        v = df.copy()
        if a_sel: v = v[v[acol].isin(a_sel)]
        if t_sel: v = v[v['MODEL_TYPE'].isin(t_sel)]
        if c_sel: v = v[v['Cus'].isin(c_sel)]
        return v

    # ── 테이블 1 ─────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📋 1) 이번 주 Cut off 예상 (계획 100%)")
    st.caption("예상제품재고 = 현재재고 + 예상계획 - 누적PO")

    t1 = mdf.rename(columns={
        '알람_계획': '알람', '예상제품재고_계획': '예상제품재고'}).copy()
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("전체",        f"{len(t1)}건")
    k2.metric("🔴 출하불가", int((t1['알람'] == '🔴 출하불가').sum()))
    k3.metric("🟠 부족",     int((t1['알람'] == '🟠 부족').sum()))
    k4.metric("🟡 차질",     int((t1['알람'] == '🟡 차질').sum()))
    k5.metric("✅ 정상",     int((t1['알람'] == '✅ 정상').sum()))

    s1c = ['알람', '순서', 'Cut off Cargo', 'Cus', 'model', 'code',
           'ERP', 'MODEL_TYPE', 'PO', '현재재고', '예상계획', '예상제품재고', 'Note']
    s1c = [c for c in s1c if c in t1.columns]
    v1 = apply_filter(t1, '알람')
    if not v1.empty:
        render_html_table(v1[s1c], height=400, table_id="t1")

    # ── 테이블 2 ─────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📋 2) 현재 실적 반영 시 Cut off 예상")
    st.caption(
        f"Cutoff시점재고 = 현재재고 + Cutoff까지실적 - Cutoff까지누적PO | "
        f"예상제품재고 = 현재재고 + ({actual_col}+미래계획) - 누적PO")

    t2 = mdf.rename(columns={
        '알람_실적': '알람', '예상제품재고_실적': '예상제품재고'}).copy()
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("전체",        f"{len(t2)}건")
    k2.metric("🔴 출하불가", int((t2['알람'] == '🔴 출하불가').sum()))
    k3.metric("🟠 부족",     int((t2['알람'] == '🟠 부족').sum()))
    k4.metric("🟡 차질",     int((t2['알람'] == '🟡 차질').sum()))
    k5.metric("✅ 정상",     int((t2['알람'] == '✅ 정상').sum()))

    s2c = ['알람', '순서', 'Cut off Cargo', 'Cus', 'model', 'code',
           'ERP', 'MODEL_TYPE', 'PO', '현재재고', '예상계획',
           actual_col, '실적차이', '예상제품재고', 'Cutoff시점재고', 'Note']
    s2c = [c for c in s2c if c in t2.columns]
    v2 = apply_filter(t2, '알람')
    if not v2.empty:
        render_html_table(v2[s2c], height=500, table_id="t2")
        csv = v2[s2c].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 CSV 다운로드", csv,
            f"shipment_alert_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv", key="dl_v30")

    st.markdown("---")
    r1, r2 = st.columns(2)
    if r1.button("🗑️ 출하계획 초기화",
                 use_container_width=True, key="rs_v30"):
        for k in ['ship_db', 'ship_updated', 'plan_date_cols', 'note_dict']:
            st.session_state.pop(k, None)
        st.cache_data.clear(); st.rerun()
    if r2.button("🗑️ 생산실적 초기화",
                 use_container_width=True, key="rp_v30"):
        for k in ['prod_db', 'prod_updated']:
            st.session_state.pop(k, None)
        st.cache_data.clear(); st.rerun()
