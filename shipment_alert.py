"""
[한솔테크닉스 HEVH] Shipment Cut-off 알람 v30.14
- 누적계획/누적실적/누적차이 추가
- 실적차이 → 누적차이로 교체
- Infor model 보정
"""
from datetime import datetime
import io
import re
import base64
import requests
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


def _gh_config():
    try: return st.secrets["github"]["token"], st.secrets["github"]["repo"]
    except: return None, None


def _gh_load_csv(filename):
    token, repo = _gh_config()
    if not token: return pd.DataFrame()
    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            return pd.read_csv(io.StringIO(base64.b64decode(r.json()["content"]).decode("utf-8")))
        return pd.DataFrame()
    except: return pd.DataFrame()


def _gh_save_csv(df, filename):
    token, repo = _gh_config()
    if not token: return False
    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    sha = None
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200: sha = r.json().get("sha")
    except: pass
    content = base64.b64encode(df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")).decode()
    payload = {"message": f"DB:{filename}", "content": content}
    if sha: payload["sha"] = sha
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=15)
        return r.status_code in [200, 201]
    except: return False


def _gh_delete(filename):
    token, repo = _gh_config()
    if not token: return False
    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha = r.json().get("sha")
            requests.delete(url, headers=headers,
                json={"message": f"DELETE:{filename}", "sha": sha}, timeout=10)
            return True
    except: pass
    return False


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
        # ⭐ 전체 날짜 형식 먼저 시도 (2026-06-23)
        ts = pd.to_datetime(s, errors='coerce')
        if not pd.isna(ts): return ts.normalize()
        # 기존 로직
        s = re.sub(r'(plan|actual|cut\s*off|cargo)[\.\s_&]*', '', s).strip()
        nums = re.findall(r'\d+', s)
        if len(nums) >= 2:
            m2, d2 = int(nums[-2]), int(nums[-1])  # ⭐ 마지막 2개 사용
            if 1 <= m2 <= 12 and 1 <= d2 <= 31:
                return pd.Timestamp(year, m2, d2)
        return None
    except: return None


def normalize_cutoff(cdt):
    if pd.isna(cdt): return pd.Timestamp(2030, 1, 1)
    ts = pd.Timestamp(cdt).normalize()
    if ts.year < 2026: ts = ts.replace(year=2026)
    return ts


def fmt_cutoff(val):
    if pd.isna(val) or val == '': return '-'
    s = str(val).replace(' 00:00:00', '').replace('T00:00:00', '')
    try:
        ts = pd.Timestamp(s)
        if ts.year == 2026: return f"{ts.month}/{ts.day}"
    except: pass
    return s


def merge_two_row_header(raw, header_row):
    main = raw.iloc[header_row].values
    sub = (raw.iloc[header_row + 1].values if header_row + 1 < len(raw) else [None] * len(main))
    merged = []
    for mh, sh in zip(main, sub):
        ms = str(mh).strip() if pd.notna(mh) else ''
        ss = str(sh).strip() if pd.notna(sh) else ''
        if ms in ('nan', 'None'): ms = ''
        if ss in ('nan', 'None'): ss = ''
        if ms and ss: merged.append(f"{ms}.{ss}")
        elif ms: merged.append(ms)
        elif ss: merged.append(ss)
        else: merged.append(f"col_{len(merged)}")
    return merged


def load_infor_model_map(file_bytes):
    try:
        sheet_names = list_sheets(file_bytes)
        target = next((s for s in sheet_names if str(s).strip().lower() == 'infor'), None)
        if target is None: return {}
        raw = read_excel_raw(file_bytes, target)
        header_row = None
        for i in range(min(5, len(raw))):
            rv = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
            if 'mes' in rv:
                header_row = i; break
        if header_row is None: header_row = 0
        df = raw.iloc[header_row + 1:].copy()
        df.columns = [str(c).strip() for c in raw.iloc[header_row].values]
        df = df.reset_index(drop=True)
        erp_col = None
        for c in df.columns:
            cl = str(c).lower().replace(' ', '').replace('(', '').replace(')', '')
            if '3in1code' in cl or 'fg' in cl:
                erp_col = c; break
        model_col = next((c for c in df.columns if 'project' in str(c).lower()), None)
        if erp_col is None or model_col is None: return {}
        result = {}
        for _, row in df.iterrows():
            erp = str(row.get(erp_col, '')).strip()
            prj = str(row.get(model_col, '')).strip()
            if not erp.startswith(('013', '018')): continue
            parts = prj.split(',')
            if len(parts) >= 2:
                model = parts[1].strip()
                if model and model not in ('nan', 'None', ''):
                    result[erp] = model
        return result
    except: return {}


def load_sheet1_notes(file_bytes):
    try:
        sheet_names = list_sheets(file_bytes)
        target = next((s for s in sheet_names if str(s).strip().lower() == 'sheet1'), None)
        if target is None: return {}
        raw = read_excel_raw(file_bytes, target)
        header_row = None
        for i in range(min(8, len(raw))):
            rv = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
            if 'model' in rv and 'erp' in rv:
                header_row = i; break
        if header_row is None: return {}
        merged = merge_two_row_header(raw, header_row)
        df = raw.iloc[header_row + 2:].copy()
        df.columns = [str(c).strip() for c in merged]
        df = df.reset_index(drop=True)
        if 'Note' not in df.columns:
            for c in df.columns:
                if str(c).lower().strip() == 'note':
                    df = df.rename(columns={c: 'Note'}); break
        if 'ERP' not in df.columns or 'Note' not in df.columns: return {}
        df['ERP'] = df['ERP'].astype(str).str.strip()
        df = df[df['ERP'].str.startswith(('013', '018'))]
        return dict(zip(df['ERP'], df['Note'].fillna('')))
    except: return {}


def load_shipment_rev(file_bytes):
    sheet_names = list_sheets(file_bytes)
    target = None
    for s in sheet_names:
        cl = str(s).strip().lower()
        if 'shipment' in cl and 'rev' in cl:
            target = s; break
    if target is None:
        target = next((s for s in sheet_names if str(s).strip().lower() == 'shipment'), None)
    if target is None:
        target = sheet_names[0]
    raw = read_excel_raw(file_bytes, target)
    header_row = None
    for i in range(min(8, len(raw))):
        rv = [str(v).lower().strip() for v in raw.iloc[i].values if pd.notna(v)]
        if 'cus' in rv:
            header_row = i; break
    if header_row is None: header_row = 2
    merged = merge_two_row_header(raw, header_row)
    df = raw.iloc[header_row + 2:].copy()
    df.columns = [str(c).strip() for c in merged]
    df = df.reset_index(drop=True)
    rename_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        clc = cl.replace(' ', '').replace('(', '').replace(')', '')
        if cl == 'cus' or 'customer' in cl: rename_map[c] = 'Cus'
        elif 'cut' in cl and ('off' in cl or 'cargo' in cl): rename_map[c] = 'Cut off Cargo'
        elif cl == 'model': rename_map[c] = 'model'
        elif cl == 'inch': rename_map[c] = 'Inch'
        elif 'hq' in cl and ('req' in cl or 'rqt' in cl or 'request' in cl): rename_map[c] = 'HQ Rqt'
        elif cl == 'bncode' or ('bn' in cl and 'code' in cl): rename_map[c] = 'code'
        elif '3in1code' in clc or cl == 'erp': rename_map[c] = 'ERP'
        elif 'po remain' in cl or 'po.remain' in cl: rename_map[c] = 'PO'
        elif 'ttl ship' in cl or 'ttl.ship' in cl: rename_map[c] = '_TTLShip'
        elif 'ttl plan' in cl or 'ttl.plan' in cl: rename_map[c] = '예상계획'
        elif 'o/stock' in cl or 'o.stock' in cl: rename_map[c] = '현재재고'
    df = df.rename(columns=rename_map)
    plan_date_cols = []
    for c in df.columns:
        cl = str(c).lower().strip()
        if ('plan' in cl or 'actual' in cl) and any(ch.isdigit() for ch in cl):
            if 'ttl' in cl: continue
            plan_date_cols.append(c)
    if 'ERP' not in df.columns:
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(30).tolist()
            if sum(1 for v in sample if (v.startswith('013') or v.startswith('018')) and len(v) >= 10) >= 3:
                df = df.rename(columns={col: 'ERP'}); break
    if 'ERP' not in df.columns:
        st.error("ERP 컬럼 매핑 실패")
        return pd.DataFrame(), []
    df['ERP'] = df['ERP'].astype(str).str.strip()
    df = df[df['ERP'].str.startswith(('013', '018'))].copy().reset_index(drop=True)
    infor_map = load_infor_model_map(file_bytes)
    if infor_map and 'model' in df.columns:
        df['model'] = df['ERP'].map(infor_map).fillna(df['model'])
    elif infor_map:
        df['model'] = df['ERP'].map(infor_map)
    return df, plan_date_cols


def save_shipment_db(df, note_dict):
    save_df = df.copy()
    if note_dict:
        save_df['Note'] = save_df['ERP'].map(note_dict).fillna(save_df.get('Note', ''))
    ok1 = _gh_save_csv(save_df, "data/shipment_db.csv")
    if note_dict:
        note_df = pd.DataFrame(list(note_dict.items()), columns=['ERP', 'Note'])
        _gh_save_csv(note_df, "data/note_db.csv")
    return ok1


def load_shipment_db():
    df = _gh_load_csv("data/shipment_db.csv")
    if df.empty: return pd.DataFrame(), [], {}
    p_cols = [c for c in df.columns
              if ('plan' in str(c).lower() or 'actual' in str(c).lower())
              and any(ch.isdigit() for ch in str(c))
              and 'ttl' not in str(c).lower()]
    note_df = _gh_load_csv("data/note_db.csv")
    if not note_df.empty and 'ERP' in note_df.columns and 'Note' in note_df.columns:
        note_dict = dict(zip(note_df['ERP'], note_df['Note'].fillna('')))
    elif 'Note' in df.columns and 'ERP' in df.columns:
        note_dict = dict(zip(df['ERP'], df['Note'].fillna('')))
    else:
        note_dict = {}
    return df, p_cols, note_dict


def analyze(ship_db, plan_date_cols, note_dict, prod_db=None):
    mdf = ship_db.copy()
    mdf['MODEL_TYPE'] = mdf['ERP'].apply(classify)
    if note_dict:
        mdf['Note'] = mdf['ERP'].map(note_dict).fillna(mdf['Note'] if 'Note' in mdf.columns else '')
    elif 'Note' not in mdf.columns:
        mdf['Note'] = ''

    for col in ['PO', '예상계획', '현재재고', '_TTLShip'] + plan_date_cols:
        if col in mdf.columns:
            mdf[col] = pd.to_numeric(mdf[col], errors='coerce').fillna(0)

    if 'PO' not in mdf.columns and '_TTLShip' in mdf.columns:
        mdf['PO'] = mdf['_TTLShip']
    elif 'PO' in mdf.columns and '_TTLShip' in mdf.columns:
        mdf['PO'] = mdf.apply(lambda r: r['PO'] if r['PO'] > 0 else r['_TTLShip'], axis=1)

    if '현재재고' not in mdf.columns: mdf['현재재고'] = 0
    if '예상계획' not in mdf.columns: mdf['예상계획'] = 0
    mdf['현재재고'] = mdf['현재재고'].astype(int)
    mdf['예상계획'] = mdf['예상계획'].astype(int)

    if 'code' not in mdf.columns: mdf['code'] = mdf['ERP']
    mdf['code'] = mdf['code'].astype(str).str.strip()

    code_stock   = mdf.groupby('code')['현재재고'].max().to_dict()
    code_plan    = mdf.groupby('code')['예상계획'].max().to_dict()
    code_erp_map = mdf.groupby('code')['ERP'].apply(lambda x: list(set(x))).to_dict()

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

    mdf['현재재고'] = mdf['code'].map(code_stock).fillna(0).astype(int)
    mdf['예상계획'] = mdf['code'].map(code_plan).fillna(0).astype(int)

    mdf['_cdt'] = pd.to_datetime(mdf['Cut off Cargo'], errors='coerce')
    mdf['Cut off Cargo'] = mdf['_cdt'].apply(fmt_cutoff)
    mdf = mdf.sort_values(['code', '_cdt']).reset_index(drop=True)

    # 시뮬레이션 1
    mdf['예상제품재고_계획'] = 0.0
    mdf['_부족1'] = 0.0
    for ck in mdf['code'].unique():
        idxs  = mdf[mdf['code'] == ck].index.tolist()
        avail = float(code_stock.get(ck, 0) + code_plan.get(ck, 0))
        for i, idx in enumerate(idxs):
            po = float(mdf.loc[idx, 'PO'])
            mdf.loc[idx, '순서'] = f"({i+1}/{len(idxs)})"
            if avail >= po:
                avail -= po
                mdf.loc[idx, '예상제품재고_계획'] = avail
                mdf.loc[idx, '_부족1'] = 0.0
            elif avail > 0:
                mdf.loc[idx, '예상제품재고_계획'] = -(po - avail)
                mdf.loc[idx, '_부족1'] = po - avail
                avail = 0.0
            else:
                mdf.loc[idx, '예상제품재고_계획'] = -po
                mdf.loc[idx, '_부족1'] = po

    mdf['알람_계획'] = mdf['_부족1'].apply(lambda x:
        "🔴 출하불가" if x >= 5000 else "🟠 부족" if x > 0 else "✅ 정상")

    today_str = None
    has_prod  = prod_db is not None and not prod_db.empty

    # 시뮬레이션 2
    if has_prod:
        prod_db = prod_db.copy()
        prod_db['TRAN_WORK_DATE'] = pd.to_datetime(prod_db['TRAN_WORK_DATE'], errors='coerce')
        today      = prod_db['TRAN_WORK_DATE'].max()
        today_str  = today.strftime('%m/%d')
        today_norm = today.normalize()
        stock_base = today_norm
        st.info(f"기준일: {today.strftime('%Y-%m-%d')} | 현재재고 기준일 이후 실적만 반영")

        erp_col = next((c for c in ['ORDER_MAD_ID', 'FINAL_MAT_ID', 'MAT_ID', 'ERP'] if c in prod_db.columns), None)
        if erp_col:
            prod_db['ERP']  = prod_db[erp_col].astype(str).str.strip()
            prod_db['TYPE'] = prod_db['ERP'].apply(classify)
            valid = prod_db[
                ((prod_db['TYPE'] == 'PD')   & (prod_db['OPER_DESC'] == 'P-ATE')) |
                ((prod_db['TYPE'] == '3IN1') & (prod_db['OPER_DESC'] == 'ASSY'))
            ].copy()

            daily = valid.groupby(['ERP', valid['TRAN_WORK_DATE'].dt.normalize()])['QTY'].sum().reset_index()
            daily.columns = ['ERP', 'DATE', 'QTY']
            daily_dict_erp   = {(r.ERP, r.DATE): r.QTY for r in daily.itertuples()}
            erp_after_actual = valid[valid['TRAN_WORK_DATE'].dt.normalize() > stock_base].groupby('ERP')['QTY'].sum().to_dict()
            erp_total_actual = valid.groupby('ERP')['QTY'].sum().to_dict()

            code_after_actual = {ck: sum(erp_after_actual.get(e, 0) for e in ev) for ck, ev in code_erp_map.items()}
            code_total_actual = {ck: sum(erp_total_actual.get(e, 0) for e in ev) for ck, ev in code_erp_map.items()}
            code_future_plan  = {ck: sum(q for d, q in dp.items() if d > today_norm) for ck, dp in code_daily_plan.items()}

            # ⭐ 누적계획 = Plan&actual 컬럼 중 today_norm 이하 날짜 합산
            code_cumul_plan = {}
            for ck in mdf['code'].unique():
                fr = mdf[mdf['code'] == ck].iloc[0]
                total = 0
                for col, dt in valid_plan.items():
                    if dt is not None and dt.normalize() <= today_norm:
                        v = pd.to_numeric(fr.get(col, 0), errors='coerce')
                        total += int(v) if not pd.isna(v) else 0
                code_cumul_plan[ck] = total

            mdf[f'현재실적({today_str})'] = mdf['code'].map(code_total_actual).fillna(0).astype(int)
            mdf['누적계획']               = mdf['code'].map(code_cumul_plan).fillna(0).astype(int)
            mdf['누적실적']               = mdf[f'현재실적({today_str})']
            mdf['누적차이']               = mdf['누적실적'] - mdf['누적계획']  # ⭐ 실적차이 대체
            mdf['_미래계획']              = mdf['code'].map(code_future_plan).fillna(0).astype(int)

            mdf['예상제품재고_실적'] = 0.0
            mdf['Cutoff시점재고']   = 0.0
            mdf['_부족2']           = 0.0

            for ck in mdf['code'].unique():
                idxs       = mdf[mdf['code'] == ck].index.tolist()
                stk        = float(code_stock.get(ck, 0))
                avail      = stk + float(code_after_actual.get(ck, 0)) + float(code_future_plan.get(ck, 0))
                erp_set_lc = set(code_erp_map.get(ck, []))
                cumul_po   = 0.0
                for idx in idxs:
                    po       = float(mdf.loc[idx, 'PO'])
                    cdt      = mdf.loc[idx, '_cdt']
                    cutoff_n = normalize_cutoff(cdt)
                    actual_until = sum(qty for (ek, d), qty in daily_dict_erp.items()
                                       if ek in erp_set_lc and d > stock_base and d <= cutoff_n)
                    cumul_po += po
                    mdf.loc[idx, 'Cutoff시점재고'] = float(stk + actual_until - cumul_po)
                    if avail >= po:
                        avail -= po
                        mdf.loc[idx, '예상제품재고_실적'] = avail
                        mdf.loc[idx, '_부족2'] = 0.0
                    elif avail > 0:
                        mdf.loc[idx, '예상제품재고_실적'] = -(po - avail)
                        mdf.loc[idx, '_부족2'] = po - avail
                        avail = 0.0
                    else:
                        mdf.loc[idx, '예상제품재고_실적'] = -po
                        mdf.loc[idx, '_부족2'] = po

            # ⭐ 알람 기준 누적차이로 교체
            mdf['알람_실적'] = mdf.apply(lambda r:
                "🔴 출하불가" if r['_부족2'] >= 5000
                else "🟠 부족" if r['_부족2'] > 0
                else "🟡 차질" if r['누적차이'] < -1000
                else "✅ 정상", axis=1)

    int_cols = ['PO', '예상계획', '현재재고', '예상제품재고_계획']
    if has_prod:
        int_cols += ['예상제품재고_실적', 'Cutoff시점재고', '누적계획', '누적실적', '누적차이',
                     f'현재실적({today_str})']
    for col in int_cols:
        if col in mdf.columns:
            mdf[col] = mdf[col].astype(int)

    mdf = mdf.drop(columns=['_부족1', '_부족2', '_cdt', '_미래계획'], errors='ignore')
    return mdf, today_str, has_prod


def render_html_table(df, height=500, table_id="t1"):
    num_cols = {c for c in df.columns if df[c].dtype.kind in 'iuf'}

    def fmt(val, is_num):
        if pd.isna(val) or val == '': return '-'
        s = str(val)
        if '00:00:00' in s:
            s = s.replace(' 00:00:00', '').replace('T00:00:00', '')
            try:
                ts = pd.Timestamp(s)
                if ts.year == 2026: return f"{ts.month}/{ts.day}"
            except: pass
            return s
        if is_num:
            try:
                n = int(float(val))
                if n < 0: return f'<span style="color:#DC2626;font-weight:700;">{n:,}</span>'
                if n == 0: return '<span style="color:#9CA3AF;">0</span>'
                return f'{n:,}'
            except: pass
        return s

    def row_bg(row):
        txt = ' '.join(str(row.get(c, '')) for c in ['알람_실적', '알람_계획', '알람'] if c in row.index)
        if '출하불가' in txt: return '#FEE2E2'
        if '부족' in txt: return '#FFEDD5'
        if '차질' in txt: return '#FEF3C7'
        return '#FFFFFF'

    head = f'''<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{{margin:0;padding:0;font-family:-apple-system,sans-serif;font-size:12px;}}
#{table_id}_w{{max-height:{height}px;overflow:auto;border:1px solid #E5E7EB;border-radius:6px;}}
#{table_id}{{border-collapse:collapse;width:100%;}}
#{table_id} th{{cursor:pointer;user-select:none;padding:8px 6px;text-align:center;
  border:1px solid #4B5563;font-weight:700;white-space:nowrap;
  background:#374151;color:#fff;position:sticky;top:0;z-index:10;}}
#{table_id} th:hover{{background:#4B5563!important;}}
#{table_id} th.sa::after{{content:" ▲";color:#FCD34D;}}
#{table_id} th.sd::after{{content:" ▼";color:#FCD34D;}}
#{table_id} td{{padding:6px;border-bottom:1px solid #E5E7EB;white-space:nowrap;}}
</style></head><body><div id="{table_id}_w"><table id="{table_id}"><thead><tr>'''

    ths = ''.join(f'<th onclick="sT({i},\'{"number" if c in num_cols else "string"}\')">{c}</th>'
                  for i, c in enumerate(df.columns))
    rows = []
    for _, row in df.iterrows():
        bg = row_bg(row)
        cells = []
        for col in df.columns:
            is_n = col in num_cols
            val  = row[col]
            cell = fmt(val, is_n)
            aln  = 'right' if is_n else 'center'
            sv   = ''
            if is_n and not pd.isna(val):
                try: sv = f' data-sort="{int(float(val))}"'
                except: pass
            cells.append(f'<td{sv} style="text-align:{aln};">{cell}</td>')
        rows.append(f'<tr style="background:{bg};">{"".join(cells)}</tr>')

    script = f'''<script>
var S_{table_id}={{}};
function sT(ci,tp){{
  var tb=document.getElementById("{table_id}");
  var bd=tb.querySelector("tbody");
  var rs=Array.from(bd.querySelectorAll("tr"));
  var hs=tb.querySelectorAll("th");
  var asc=S_{table_id}[ci]==="asc";
  var dir=asc?"desc":"asc";
  S_{table_id}={{}};S_{table_id}[ci]=dir;
  hs.forEach(h=>{{h.classList.remove("sa","sd");}});
  hs[ci].classList.add(dir==="asc"?"sa":"sd");
  rs.sort((a,b)=>{{
    var ca=a.cells[ci],cb=b.cells[ci],va,vb;
    if(tp==="number"){{
      va=parseFloat(ca.getAttribute("data-sort")||ca.textContent.replace(/[,\s]/g,""))||0;
      vb=parseFloat(cb.getAttribute("data-sort")||cb.textContent.replace(/[,\s]/g,""))||0;
    }}else{{va=ca.textContent.trim();vb=cb.textContent.trim();}}
    return dir==="asc"?(va<vb?-1:va>vb?1:0):(va>vb?-1:va<vb?1:0);
  }});
  rs.forEach(r=>bd.appendChild(r));
}}
</script></body></html>'''

    html = head + ths + '</tr></thead><tbody>' + ''.join(rows) + '</tbody></table></div>' + script
    components.html(html, height=height + 50, scrolling=False)


def render_shipment_alert_tab():
    if 'ship_db' not in st.session_state or st.session_state.get('ship_db', pd.DataFrame()).empty:
        with st.spinner("출하계획 DB 로드 중..."):
            df, p_cols, note_dict = load_shipment_db()
            if not df.empty:
                st.session_state['ship_db']        = df
                st.session_state['plan_date_cols'] = p_cols
                st.session_state['note_dict']      = note_dict
                st.session_state['ship_updated']   = 'DB 로드'

    if 'prod_db' not in st.session_state or st.session_state.get('prod_db', pd.DataFrame()).empty:
        with st.spinner("생산실적 DB 로드 중..."):
            df = _gh_load_csv("data/production.csv")
            if not df.empty:
                st.session_state['prod_db']      = df
                st.session_state['prod_updated'] = 'DB 로드'

    ship_db   = st.session_state.get('ship_db',        pd.DataFrame())
    plan_cols = st.session_state.get('plan_date_cols', [])
    note_dict = st.session_state.get('note_dict',      {})
    prod_db   = st.session_state.get('prod_db',        pd.DataFrame())
    ship_t    = st.session_state.get('ship_updated',   '-')
    prod_t    = st.session_state.get('prod_updated',   '-')

    with st.sidebar:
        st.markdown("---")
        st.markdown("#### Shipment 파일 업로드")
        if not ship_db.empty:
            st.success(f"출하계획: {len(ship_db)}건 | {ship_t}")
        else:
            st.warning("출하계획: 없음")
        if not prod_db.empty:
            st.success(f"생산실적: {len(prod_db)}건 | {prod_t}")
        else:
            st.info("생산실적: 없음 (선택사항)")

        uploaded = st.file_uploader(
            "출하계획(xlsx) + 생산실적(csv)",
            type=["xlsx", "csv"],
            accept_multiple_files=True,
            key="ship_files_v30")

        if st.button("저장 및 분석", type="primary", use_container_width=True, key="apply_v30"):
            for f in uploaded:
                f.seek(0); fb = f.read()
                if f.name.lower().endswith('.xlsx'):
                    with st.spinner("출하계획 처리 중..."):
                        df, p_cols = load_shipment_rev(fb)
                        notes      = load_sheet1_notes(fb)
                        if not df.empty:
                            st.session_state['ship_db']        = df
                            st.session_state['plan_date_cols'] = p_cols
                            st.session_state['note_dict']      = notes
                            st.session_state['ship_updated']   = datetime.now().strftime('%m-%d %H:%M')
                            ok = save_shipment_db(df, notes)
                            st.success(f"출하계획 {len(df)}건 {'저장완료' if ok else '저장실패'}")
                elif f.name.lower().endswith('.csv'):
                    with st.spinner("생산실적 처리 중..."):
                        df = read_csv_cached(fb)
                        st.session_state['prod_db']      = df
                        st.session_state['prod_updated'] = datetime.now().strftime('%m-%d %H:%M')
                        ok = _gh_save_csv(df, "data/production.csv")
                        st.success(f"생산실적 {len(df)}건 {'저장완료' if ok else '저장실패'}")
            st.rerun()

        st.markdown("---")
        if st.button("출하계획 초기화", use_container_width=True, key="rs_v30"):
            for k in ['ship_db', 'ship_updated', 'plan_date_cols', 'note_dict']:
                st.session_state.pop(k, None)
            _gh_delete("data/shipment_db.csv")
            _gh_delete("data/note_db.csv")
            st.cache_data.clear(); st.rerun()
        if st.button("생산실적 초기화", use_container_width=True, key="rp_v30"):
            for k in ['prod_db', 'prod_updated']:
                st.session_state.pop(k, None)
            _gh_delete("data/production.csv")
            st.cache_data.clear(); st.rerun()

    if ship_db.empty:
        st.info("사이드바에서 출하계획(.xlsx)을 업로드하세요.")
        return

    with st.spinner("분석 중..."):
        try:
            mdf, today_str, has_prod = analyze(
                ship_db, plan_cols, note_dict,
                prod_db if not prod_db.empty else None)
        except Exception as e:
            st.error(f"분석 오류: {e}")
            import traceback; st.code(traceback.format_exc())
            return

    if mdf.empty: return

    actual_col     = f'현재실적({today_str})' if today_str else None
    problem_alerts = ['🔴 출하불가', '🟠 부족', '🟡 차질']

    # 즉시 조치
    st.markdown("### 🚨 즉시 조치 필요")
    st.caption("두 시뮬레이션 중 하나라도 문제 발생한 행 | 컬럼 클릭 정렬")
    urgent_mask = mdf['알람_계획'].isin(problem_alerts)
    if has_prod:
        urgent_mask = urgent_mask | mdf['알람_실적'].isin(problem_alerts)
    urgent = mdf[urgent_mask].copy()

    if not urgent.empty:
        k1, k2, k3 = st.columns(3)
        k1.metric("문제 행", f"{len(urgent)}건")
        k2.metric("출하불가(계획)", int((urgent['알람_계획'] == '🔴 출하불가').sum()))
        if has_prod:
            k3.metric("출하불가(실적)", int((urgent['알람_실적'] == '🔴 출하불가').sum()))
        cc = ['알람_계획']
        if has_prod: cc.append('알람_실적')
        cc += ['순서', 'HQ Rqt', 'Cut off Cargo', 'Cus', 'model', 'code', 'ERP',
               'MODEL_TYPE', 'PO', '현재재고', '예상계획']
        if actual_col and actual_col in urgent.columns:
            cc += [actual_col, '누적계획', '누적실적', '누적차이']
        cc.append('예상제품재고_계획')
        if has_prod: cc += ['예상제품재고_실적', 'Cutoff시점재고']
        cc.append('Note')
        cc = [c for c in cc if c in urgent.columns]
        render_html_table(urgent[cc], height=500, table_id="compare")
    else:
        st.success("✅ 모든 모델 정상")

    # 테이블 1
    st.markdown("---")
    st.markdown("### 📋 1) Cut off 예상 (계획 100%)")
    st.caption("예상제품재고 = 현재재고 + 예상계획 - 누적PO")
    t1 = mdf.rename(columns={'알람_계획': '알람', '예상제품재고_계획': '예상제품재고'}).copy()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("전체", f"{len(t1)}건")
    k2.metric("🔴 출하불가", int((t1['알람'] == '🔴 출하불가').sum()))
    k3.metric("🟠 부족",     int((t1['알람'] == '🟠 부족').sum()))
    k4.metric("✅ 정상",     int((t1['알람'] == '✅ 정상').sum()))
    s1c = ['알람', '순서', 'HQ Rqt', 'Cut off Cargo', 'Cus', 'model', 'code',
           'ERP', 'MODEL_TYPE', 'PO', '현재재고', '예상계획', '예상제품재고', 'Note']
    s1c = [c for c in s1c if c in t1.columns]
    if not t1.empty:
        render_html_table(t1[s1c], height=400, table_id="t1")

    # 테이블 2
    if has_prod:
        st.markdown("---")
        st.markdown("### 📋 2) Cut off 예상 (실적 반영)")
        st.caption("누적차이 = 누적실적 - 누적계획 | Cutoff시점재고 = 현재재고 + 기준일이후실적 - 누적PO")
        t2 = mdf.rename(columns={'알람_실적': '알람', '예상제품재고_실적': '예상제품재고'}).copy()
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("전체",        f"{len(t2)}건")
        k2.metric("🔴 출하불가", int((t2['알람'] == '🔴 출하불가').sum()))
        k3.metric("🟠 부족",     int((t2['알람'] == '🟠 부족').sum()))
        k4.metric("🟡 차질",     int((t2['알람'] == '🟡 차질').sum()))
        k5.metric("✅ 정상",     int((t2['알람'] == '✅ 정상').sum()))
        s2c = ['알람', '순서', 'HQ Rqt', 'Cut off Cargo', 'Cus', 'model', 'code',
               'ERP', 'MODEL_TYPE', 'PO', '현재재고', '예상계획',
               actual_col, '누적계획', '누적실적', '누적차이',
               '예상제품재고', 'Cutoff시점재고', 'Note']
        s2c = [c for c in s2c if c in t2.columns]
        if not t2.empty:
            render_html_table(t2[s2c], height=500, table_id="t2")
            csv = t2[s2c].to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 CSV 다운로드", csv,
                f"shipment_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                "text/csv", key="dl_v30")
