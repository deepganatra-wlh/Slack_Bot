from flask import Blueprint, request, jsonify, send_file, render_template
import pandas as pd
import numpy as np
import os, json, uuid, traceback, requests
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

banca_bp = Blueprint('banca', __name__, template_folder='templates')

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── AI Assistant (Groq) config ────────────────────────────────────────────────
# Set your key: export GROQ_API_KEY="gsk_...."   (free, no card — get one at https://console.groq.com/keys)
GROQ_API_KEY  = os.environ.get('GROQ_API_KEY', '')
GROQ_MODEL    = os.environ.get('GROQ_MODEL', 'llama-3.3-70b-versatile')
GROQ_API_URL  = 'https://api.groq.com/openai/v1/chat/completions'

ALLOWED_EXTENSIONS = {'xlsx', 'xlsb', 'xls', 'csv'}
def allowed_file(f): return '.' in f and f.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

# ─── RTO ──────────────────────────────────────────────────────────────────────
def build_rto_mapping(path):
    try:
        xl = pd.ExcelFile(path)
        rto_df = None
        for s in ['Data','RTOSRWithoutChannel','RTO','Sheet1']:
            if s in xl.sheet_names:
                rto_df = pd.read_excel(path, sheet_name=s); break
        if rto_df is None or 'UW Budget Cluster' not in rto_df.columns: return {}
        rto_df['UW Budget Cluster'] = rto_df['UW Budget Cluster'].astype(str).str.strip()
        rto_df['Code'] = rto_df['Code'].astype(str).str.strip()
        rto_df = rto_df[rto_df['UW Budget Cluster'].notna() & (rto_df['UW Budget Cluster'] != '')]
        out = {}
        for cl in rto_df['UW Budget Cluster'].unique():
            codes = rto_df[rto_df['UW Budget Cluster']==cl]['Code'].dropna().unique().tolist()
            out[str(cl)] = ','.join(sorted(set(map(str,codes))))
        return out
    except Exception as e:
        print(f"RTO error: {e}"); return {}

# ─── Detriff ranges ───────────────────────────────────────────────────────────
DETRIFF_RANGES = [
    ('0',0,0.999),('0-5',1.0,5.001),('1-5',1.001,5.001),('0.0',0,1.001),
    ('6-10',5.001,10.001),('11-15',10.001,15.001),('16-20',15.001,20.001),
    ('21-25',20.001,25.001),('26-30',25.001,30.001),('31-35',30.001,35.001),
    ('36-40',35.001,40.001),('41-45',40.001,45.001),('46-50',45.001,50.001),
    ('51-55',50.001,55.001),('56-60',55.001,60.001),('61-65',60.001,65.001),
    ('66-70',65.001,70.001),('71-75',70.001,75.001),('76-80',75.001,80.001),
    ('81-85',80.001,85.001),('86-90',85.001,90.001),('91-95',90.001,95.001),
    ('96-100',95.001,100.001),
]
def get_dr(name):
    for n,l,u in DETRIFF_RANGES:
        if n==name: return l,u
    return None,None

# ─── Detriff merger ───────────────────────────────────────────────────────────
def apply_detriff_merger(df):
    ll = next((c for c in df.columns if 'detriff' in c.lower() and 'll' in c.lower()),None)
    ul = next((c for c in df.columns if 'detriff' in c.lower() and 'ul' in c.lower()),None)
    if not ll or not ul: return df, 0
    grp = [c for c in df.columns if c not in [ll,ul,'Unique Id','unique_id']]
    df[ll] = pd.to_numeric(df[ll], errors='coerce')
    df[ul] = pd.to_numeric(df[ul], errors='coerce')
    orig = len(df)
    agg = {ll:'min', ul:'max'}
    for u in [c for c in df.columns if c in ['Unique Id','unique_id']]: agg[u]='first'
    res = df.groupby(grp, as_index=False, dropna=False).agg(agg)
    res = res[[c for c in df.columns if c in res.columns]]
    return res, orig - len(res)

# ─── Transformation engine ────────────────────────────────────────────────────
def apply_transformations(df, transformations):
    df = df.copy()
    for t in transformations:
        col = t.get('column')
        ops = t.get('ops', [])
        if col not in df.columns: continue
        for op in ops:
            ot = op.get('type')
            if ot == 'case':
                c = op.get('value','upper')
                if c=='upper': df[col]=df[col].astype(str).str.upper()
                elif c=='lower': df[col]=df[col].astype(str).str.lower()
                elif c=='title': df[col]=df[col].astype(str).str.title()
            elif ot == 'trim':
                df[col]=df[col].astype(str).str.strip()
            elif ot == 'strip_chars':
                df[col]=df[col].astype(str).str.strip(op.get('chars',''))
            elif ot == 'replace':
                df[col]=df[col].astype(str).str.replace(op.get('find',''),op.get('with',''),regex=False)
            elif ot == 'regex_replace':
                df[col]=df[col].astype(str).str.replace(op.get('pattern',''),op.get('with',''),regex=True)
            elif ot == 'value_map':
                vm = op.get('map',{})
                df[col]=df[col].astype(str).map(lambda x: vm.get(x,x))
            elif ot == 'prefix':
                df[col]=op.get('value','')+df[col].astype(str)
            elif ot == 'suffix':
                df[col]=df[col].astype(str)+op.get('value','')
            elif ot == 'default_if_empty':
                dv = op.get('value','ANY')
                df[col]=df[col].apply(lambda x: dv if (pd.isna(x) or str(x).strip() in ('','nan','None')) else x)
            elif ot == 'number_format':
                df[col]=pd.to_numeric(df[col],errors='coerce').round(op.get('decimals',2))
            elif ot == 'math_op':
                expr = op.get('expression','').strip()
                decimals = op.get('decimals', None)   # optional rounding
                if expr:
                    import math as _math
                    _safe_ns = {
                        'math': _math, 'abs': abs, 'round': round,
                        'min': min, 'max': max, 'pow': pow,
                        'sqrt': _math.sqrt, 'log': _math.log,
                        'log10': _math.log10, 'ceil': _math.ceil,
                        'floor': _math.floor, 'pi': _math.pi, 'e': _math.e,
                    }
                    def _apply_math(cell_val, _expr=expr, _ns=_safe_ns, _dec=decimals):
                        try:
                            x = float(cell_val)          # try numeric conversion
                            result = eval(_expr, {"__builtins__": {}}, {**_ns, 'x': x})
                            result = float(result)
                            if _dec is not None:
                                result = round(result, _dec)
                            # Return int string if whole number, else float string
                            return int(result) if result == int(result) else result
                        except Exception:
                            return cell_val              # keep original on any error
                    df[col] = df[col].apply(_apply_math)
            elif ot == 'conditional':
                for rule in op.get('rules',[]):
                    if_col   = rule.get('if_col')
                    if_op    = rule.get('if_op','eq')
                    if_val   = str(rule.get('if_val','')).lower()
                    then_col = rule.get('then_col', col)
                    if if_col not in df.columns or then_col not in df.columns: continue
                    if if_op=='eq':      mask=df[if_col].astype(str).str.lower().str.strip()==if_val
                    elif if_op=='neq':   mask=df[if_col].astype(str).str.lower().str.strip()!=if_val
                    elif if_op=='contains': mask=df[if_col].astype(str).str.lower().str.contains(if_val,na=False)
                    elif if_op=='empty': mask=df[if_col].astype(str).str.strip().isin(['','nan','None','ANY'])
                    elif if_op=='notempty': mask=~df[if_col].astype(str).str.strip().isin(['','nan','None','ANY'])
                    else: mask=pd.Series([False]*len(df))
                    for tr in rule.get('then_rules',[]):
                        wv  = str(tr.get('when_val','')).lower()
                        sv  = tr.get('set_val','')
                        wop = tr.get('when_op','eq')
                        if wop=='any' or wv=='*':      rm=mask
                        elif wop=='empty' or wv=='':   rm=mask&df[then_col].astype(str).str.strip().isin(['','nan','None','ANY'])
                        elif wop=='contains':          rm=mask&df[then_col].astype(str).str.lower().str.contains(wv,na=False)
                        else:                          rm=mask&(df[then_col].astype(str).str.lower().str.strip()==wv)
                        df.loc[rm, then_col]=sv
    return df

# ─── Read Excel ───────────────────────────────────────────────────────────────
def read_excel_file(filepath, sheet_name, header_row, start_col):
    ext = filepath.rsplit('.',1)[1].lower()
    if ext=='csv': df=pd.read_csv(filepath,header=None)
    elif ext=='xlsb': df=pd.read_excel(filepath,sheet_name=sheet_name,header=None,engine='pyxlsb')
    else: df=pd.read_excel(filepath,sheet_name=sheet_name,header=None)
    hr,sc = header_row-1, start_col-1
    headers = df.iloc[hr,sc:].values.tolist()
    data    = df.iloc[hr+1:,sc:sc+len(headers)].values
    sdf = pd.DataFrame(data, columns=headers)
    sdf.columns = sdf.columns.map(lambda c: str(c).strip())
    sdf.dropna(how='all',inplace=True); sdf.reset_index(drop=True,inplace=True)
    return sdf

# ─── Standard grid processor ──────────────────────────────────────────────────
def process_grid(source_df, mapping_config, rto_mapping=None, passthrough_cols=None):
    th  = mapping_config.get('target_headers',[])
    cm  = mapping_config.get('column_mapping',{})
    rcs = mapping_config.get('rto_cluster_source','UW Budget Cluster')
    sv  = mapping_config.get('static_values',{})
    # Only carry through cols that actually exist in source and aren't already in target
    pt  = [c for c in (passthrough_cols or []) if c in source_df.columns and c not in th]
    seen=set(); avd=[]
    for name,_,_ in DETRIFF_RANGES:
        if name in source_df.columns and name not in seen:
            seen.add(name); avd.append(name)
    rows=[]
    for idx,row in source_df.iterrows():
        if avd:
            for rn in avd:
                val=row.get(rn)
                if pd.isna(val): continue
                ll,ul=get_dr(rn)
                rows.append(_build_row(row,idx,source_df,th,cm,sv,rto_mapping,rcs,val,rn,ll,ul,pt))
        else:
            rows.append(_build_row(row,idx,source_df,th,cm,sv,rto_mapping,rcs,None,None,None,None,pt))
    if not rows: return None
    out_cols = th + pt
    return pd.DataFrame(rows, columns=out_cols)

def _build_row(row,idx,src,th,cm,sv,rtom,rcs,dval,rn,ll,ul,passthrough_cols=None):
    nr={}
    for col in th:
        cl=col.lower().replace('*','').strip()
        if cl.endswith('ll') and ('capacity' in cl or 'weight' in cl): nr[col]=-10000000
        elif cl.endswith('ul') and ('capacity' in cl or 'weight' in cl): nr[col]=10000000
        elif cl.endswith('ll') and 'detriff' in cl: nr[col]=ll if ll is not None else 'ANY'
        elif cl.endswith('ul') and 'detriff' in cl: nr[col]=ul if ul is not None else 'ANY'
        elif 'total gwp ll' in cl or 'gwp_ll' in cl: nr[col]=-10000000
        elif 'total gwp ul' in cl or 'gwp_ul' in cl: nr[col]=10000000
        else: nr[col]='ANY'
    for tc,si in cm.items():
        if tc not in th: continue
        if si=='__DETRIFF_VALUE__': nr[tc]=dval if dval is not None else 'ANY'
        elif si=='__DETRIFF_RANGE__': nr[tc]=rn if rn else 'ANY'
        elif si=='__UNIQUE_ID__': nr[tc]=f"row_{idx}_{rn}" if rn else f"row_{idx}"
        elif si=='__RTO_CODES__':
            cl=row.get(rcs); nr[tc]=rtom.get(str(cl).strip(),'ANY') if rtom and pd.notna(cl) else 'ANY'
        elif isinstance(si,str) and si.startswith('__LITERAL__:'): nr[tc]=si[len('__LITERAL__:'):]
        else:
            if isinstance(si,str) and si in src.columns:
                v=row.get(si); nr[tc]=v if pd.notna(v) else 'ANY'
            else: nr[tc]=si
    for col,val in sv.items():
        if col in th: nr[col]=val
    uid='Unique Id'
    if uid in th and nr.get(uid)=='ANY':
        nr[uid]=f"row_{idx}_{rn}" if rn else f"row_{idx}"
    # Carry passthrough source columns verbatim
    for pc in (passthrough_cols or []):
        nr[pc] = row.get(pc, None)
    return nr

# ─── SK Finance processor ──────────────────────────────────────────────────────
def cc_to_ll_ul(v):
    v=str(v).strip()
    if v in ('NA','None','nan',''): return -999999999,999999999
    if v=='<1000 CC': return 0,999
    if v=='1000-1500 CC': return 1000,1500
    if v=='>1500 CC': return 1501,999999999
    return -999999999,999999999

def ncb_to_ll_ul(v):
    v=str(v).strip()
    if v=='All': return 0,101
    if v=='Yes': return 1,101
    if v=='No': return 0,0
    return 0,101

def pb_to_outgo(v):
    return {'on OD':'OD','on TP':'TP'}.get(str(v).strip(),str(v).strip())

def process_sk_finance(config, rto_mapping):
    """
    SK Finance processor with OD+TP merge logic:

    The sheet has mirrored OD columns and TP columns for the same LOBs.
    For each unique (LOB, NCB, CC) combination:
      - If only OD rate exists  → banca_outgo = OD,  banca_total_comm_rate = OD rate, banca_tp_prct = 0
      - If only TP rate exists  → banca_outgo = TP,  banca_total_comm_rate = TP rate, banca_tp_prct = 0
      - If BOTH OD and TP exist → banca_outgo = OD AND TP, banca_total_comm_rate = OD rate, banca_tp_prct = TP rate
    This produces ONE output row per LOB (instead of two separate OD/TP rows).
    """
    filepath = config['filepath']
    sheet    = config.get('sheet_name','SK Finance')
    wb = load_workbook(filepath, data_only=True)
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    ws = wb[sheet]

    ar,ac = config.get('agent_code_cell',[1,2])
    rr,rc = config.get('relationship_code_cell',[3,2])
    agent_code = ws.cell(row=ar,column=ac).value
    rel_raw    = ws.cell(row=rr,column=rc).value
    try: rel_code = str(int(float(str(rel_raw))))
    except: rel_code = str(rel_raw or '')

    pb_ri = config.get('payment_basis_row',9)
    nb_ri = config.get('ncb_row',10)
    cc_ri = config.get('cc_row',11)
    lb_ri = config.get('lob_row',12)
    ds    = config.get('data_start_row',14)
    sc    = config.get('state_col',1)
    cc_col= config.get('cluster_col',2)
    dc    = config.get('data_start_col',3)
    vpb   = config.get('valid_payment_basis',['on OD','on TP'])
    ver   = config.get('version_id','')
    # The output column name used for OD+TP combined outgo label (configurable)
    od_tp_label = config.get('od_tp_combined_label','OD AND TP')
    mc    = ws.max_column; mr = ws.max_row

    pbv=[ws.cell(row=pb_ri,column=c).value for c in range(1,mc+1)]
    nbv=[ws.cell(row=nb_ri,column=c).value for c in range(1,mc+1)]
    ccv=[ws.cell(row=cc_ri,column=c).value for c in range(1,mc+1)]
    lbv=[ws.cell(row=lb_ri,column=c).value for c in range(1,mc+1)]

    # Separate OD and TP data columns
    # Each entry: {'col_idx', 'payment_basis', 'ncb', 'cc', 'lob', 'pb_type'}
    # pb_type = 'OD' or 'TP'
    od_cols = []
    tp_cols = []
    for c in range(dc, mc+1):
        pb  = pbv[c-1]
        lob = lbv[c-1]
        ncb = nbv[c-1]
        cc  = ccv[c-1]
        if not lob or pb not in vpb:
            continue
        entry = {'col_idx':c,'payment_basis':pb,'ncb':ncb,'cc':cc,'lob':lob}
        if str(pb).strip() == 'on OD':
            od_cols.append(entry)
        elif str(pb).strip() == 'on TP':
            tp_cols.append(entry)

    # Build lookup: (lob, ncb, cc) → col descriptor, for both OD and TP
    def _key(e):
        return (str(e['lob']).strip(), str(e['ncb'] or '').strip(), str(e['cc'] or '').strip())

    od_map = {_key(e): e for e in od_cols}
    tp_map = {_key(e): e for e in tp_cols}

    # All unique LOB keys across both OD and TP
    all_keys = list(dict.fromkeys(list(od_map.keys()) + list(tp_map.keys())))

    out_cols = config.get('output_columns',[
        'version_id','pol_agent_code','primary_agent_code','product_code','proposal_id',
        'biz_mix','two_wheeler_category','type_of_business','oem_type','final_region',
        'rto_code','detriff_discount_ll','detriff_discount_ul','gross_vehicle_weight_ll',
        'gross_vehicle_weight_ul','cubic_capacity_ll','cubic_capacity_ul','bus_type',
        'model','fuel_type','banca_cpa_flag','banca_misp','banca_flat_amount',
        'banca_capping','banca_tp_prct','banca_tp_outgo','banca_total_comm_rate',
        'banca_outgo','product_desc','state','ncb_percentage_ll','ncb_percentage_ul',
        'body_type','txt_motor_nonmotor'
    ])
    out_map = config.get('output_mapping', _default_sk_mapping())

    rows = []
    for ri in range(ds, mr+1):
        sv = ws.cell(row=ri,column=sc).value
        cv = ws.cell(row=ri,column=cc_col).value
        if not sv or not cv: continue
        ss = str(sv).strip()
        cs = str(cv).strip()
        rto = rto_mapping.get(cs,'ANY') if rto_mapping else 'ANY'

        for lob_key in all_keys:
            od_entry = od_map.get(lob_key)
            tp_entry = tp_map.get(lob_key)

            # Read OD rate (None if column absent)
            od_rate = None
            if od_entry:
                v = ws.cell(row=ri, column=od_entry['col_idx']).value
                if v is not None:
                    od_rate = v

            # Read TP rate (None if column absent)
            tp_rate = None
            if tp_entry:
                v = ws.cell(row=ri, column=tp_entry['col_idx']).value
                if v is not None:
                    tp_rate = v

            # Determine which rates are "real" (non-zero, non-None)
            has_od = od_rate is not None and float(od_rate) != 0
            has_tp = tp_rate is not None and float(tp_rate) != 0

            # Skip if neither OD nor TP has a meaningful rate
            if not has_od and not has_tp:
                continue

            # ── Merge logic ────────────────────────────────────────────────
            # Use the representative entry for NCB/CC metadata
            rep = od_entry if od_entry else tp_entry

            if has_od and has_tp:
                # Both present → combined row
                banca_outgo         = od_tp_label   # e.g. "OD AND TP"
                banca_total_comm    = od_rate        # OD rate goes in main comm col
                banca_tp_prct_val   = tp_rate        # TP rate goes in TP prct col
                banca_tp_outgo_val  = 'TP'
            elif has_od:
                # Only OD
                banca_outgo         = 'OD'
                banca_total_comm    = od_rate
                banca_tp_prct_val   = 0
                banca_tp_outgo_val  = 'NA'
            else:
                # Only TP
                banca_outgo         = 'TP'
                banca_total_comm    = tp_rate
                banca_tp_prct_val   = 0
                banca_tp_outgo_val  = 'NA'

            ccll, ccul = cc_to_ll_ul(rep['cc'])
            nll,  nul  = ncb_to_ll_ul(rep['ncb'])

            ctx = {
                '__agent_code__':        agent_code,
                '__relationship_code__': rel_code,
                '__state__':             ss,
                '__cluster__':           cs,
                '__rto_codes__':         rto,
                '__comm_rate__':         banca_total_comm,
                '__tp_rate__':           banca_tp_prct_val,
                '__cc_ll__':             ccll,
                '__cc_ul__':             ccul,
                '__ncb_ll__':            nll,
                '__ncb_ul__':            nul,
                '__banca_outgo__':       banca_outgo,
                '__banca_tp_outgo__':    banca_tp_outgo_val,
                '__banca_tp_prct__':     banca_tp_prct_val,
                '__lob__':               rep['lob'],
                '__payment_basis__':     rep['payment_basis'],
                '__ncb__':               rep['ncb'],
                '__cc__':                rep['cc'],
                '__version_id__':        ver,
            }

            ro = {}
            for oc in out_cols:
                if oc in out_map:
                    vs = out_map[oc]
                    if isinstance(vs,str) and vs.startswith('__LITERAL__:'):
                        ro[oc] = vs[len('__LITERAL__:'):]
                    elif isinstance(vs,str) and vs in ctx:
                        ro[oc] = ctx[vs]
                    else:
                        ro[oc] = vs
                else:
                    ro[oc] = 'ANY'
            rows.append(ro)

    total_data_cols = len(od_cols) + len(tp_cols)
    meta = {
        'agent_code':       str(agent_code),
        'relationship_code': rel_code,
        'data_cols_count':  total_data_cols,
        'od_cols':          len(od_cols),
        'tp_cols':          len(tp_cols),
        'unique_lob_keys':  len(all_keys),
        'sheets':           wb.sheetnames,
    }
    return pd.DataFrame(rows, columns=out_cols), meta

def _default_sk_mapping():
    return {
        'version_id':'__LITERAL__:','pol_agent_code':'__agent_code__',
        'primary_agent_code':'__relationship_code__','product_code':'__LITERAL__:ANY',
        'proposal_id':'__LITERAL__:ANY','biz_mix':'__lob__',
        'two_wheeler_category':'__LITERAL__:ANY','type_of_business':'__LITERAL__:ANY',
        'oem_type':'__LITERAL__:ANY','final_region':'__state__','rto_code':'__rto_codes__',
        'detriff_discount_ll':'__LITERAL__:0','detriff_discount_ul':'__LITERAL__:100.001',
        'gross_vehicle_weight_ll':'__LITERAL__:-999999999','gross_vehicle_weight_ul':'__LITERAL__:999999999',
        'cubic_capacity_ll':'__cc_ll__','cubic_capacity_ul':'__cc_ul__',
        'bus_type':'__LITERAL__:ANY','model':'__LITERAL__:ANY','fuel_type':'__LITERAL__:ANY',
        'banca_cpa_flag':'__LITERAL__:0','banca_misp':'__LITERAL__:No',
        'banca_flat_amount':'__LITERAL__:0','banca_capping':'__LITERAL__:0',
        # banca_tp_prct: receives TP rate when both OD+TP present, else 0
        'banca_tp_prct':'__banca_tp_prct__',
        # banca_tp_outgo: 'TP' when combined, 'NA' otherwise
        'banca_tp_outgo':'__banca_tp_outgo__',
        # banca_total_comm_rate: always gets the OD rate (or sole rate if only one)
        'banca_total_comm_rate':'__comm_rate__',
        # banca_outgo: 'OD AND TP' / 'OD' / 'TP'
        'banca_outgo':'__banca_outgo__',
        'product_desc':'__LITERAL__:ANY','state':'__LITERAL__:ANY',
        'ncb_percentage_ll':'__ncb_ll__','ncb_percentage_ul':'__ncb_ul__',
        'body_type':'__LITERAL__:ANY','txt_motor_nonmotor':'__LITERAL__:Motor',
    }

# ─── AI Assistant ──────────────────────────────────────────────────────────────
AI_ACTION_TYPES = (
    'set_mapping', 'set_target_headers', 'add_transformation',
    'remove_transformation', 'navigate', 'none'
)

def build_ai_system_prompt(context, has_config):
    """Builds the system prompt sent to the LLM, grounding it in this portal's
    actual data model so it never invents features that don't exist."""
    mode = context.get('mode', 'std')
    tx_types_doc = (
        "case{value:upper|lower|title}, trim{}, strip_chars{chars}, "
        "replace{find,with}, regex_replace{pattern,with}, value_map{map:{old:new}}, "
        "prefix{value}, suffix{value}, default_if_empty{value}, "
        "number_format{decimals}, math_op{expression (python expr using x),decimals}, "
        "conditional{rules:[{if_col,if_op:eq|neq|contains|empty|notempty,if_val,then_col,"
        "then_rules:[{when_op:eq|contains|empty|any,when_val,set_val}]}]}"
    )
    special_markers_doc = (
        "__LITERAL__:value (fixed constant), __DETRIFF_VALUE__ (the detriff column's value for that row), "
        "__RTO_CODES__ (comma list of RTO codes from the RTO cluster file), "
        "__UNIQUE_ID__ (auto generated row id), __DETRIFF_RANGE__ (the detriff band name, e.g. 6-10)"
    )
    ctx_json = json.dumps({
        'mode': mode,
        'sheet_name': context.get('sheet_name'),
        'target_headers': context.get('target_headers', []),
        'column_mapping': context.get('column_mapping', {}),
        'source_columns': (context.get('source_columns') or [])[:80],
        'transformations': context.get('transformations', []),
        'has_initial_config': has_config,
    }, ensure_ascii=False)

    return f"""You are the in-app AI Assistant for "Grid Processor Portal", a tool that converts \
insurance/OEM commission grid spreadsheets into standardized target CSV outputs.

The portal has two modes, each a 5-step wizard:
- Standard Grid (mode 'std'): 1 Upload Files, 2 Configure Source (sheet/header row/start col), \
3 Column Mapping (Target Headers + Column Mapping JSON, mapping each target header to a source \
column name, a __LITERAL__:value, or a special marker), 4 Transformations (per-column op chains), \
5 Process & Export.
- SK Finance (mode 'sk'): 1 Upload Files, 2 Sheet Structure, 3 Output Mapping, 4 Transformations, \
5 Process & Export. This mode has its own fixed cell-based config (agent code cell, payment basis \
row, etc.) rather than a target_headers/column_mapping pair.

Special markers usable as a mapping value: {special_markers_doc}
Transformation operation types and params: {tx_types_doc}
Built-in presets exist for common OEMs (TATA PV, TATA CV, Banca Motor, Hyundai, Nissan/Renault, \
OEM Other) — each is a ready-made target_headers + column_mapping pair the user can pick from a \
preset grid in Step 2/3 instead of typing config by hand.

The user's CURRENT state in the app (ground truth — do not assume anything not shown here):
{ctx_json}

CRITICAL RULE — initial config gate:
The user can only ask you to EDIT config (mapping, target headers, or transformations) once they \
already have a basic initial config loaded — meaning target_headers and column_mapping are non-empty \
(std mode) or they've completed Sheet Structure/Output Mapping (sk mode). "has_initial_config" above \
tells you whether that's true right now.
- If has_initial_config is false and the user asks you to edit/add/change something, do NOT return \
any edit actions. Instead explain kindly that they should first pick a preset or fill in Target \
Headers + Column Mapping in Step 3 (or complete Step 2/3 in SK Finance mode), and offer a "navigate" \
action to take them there.
- If has_initial_config is true, you may return concrete edit actions.
- You can ALWAYS answer general "how do I..." / "where is..." questions about using the portal, \
regardless of has_initial_config — that never requires an initial config.

You must reply with ONLY a single JSON object (no markdown fences, no prose outside it) matching \
exactly this shape:
{{"reply": "<short, friendly, concrete answer to show the user, 1-6 sentences>",
  "actions": [ {{"type": "<one of {', '.join(AI_ACTION_TYPES)}>", ...fields}} ]}}

Action field shapes:
- set_mapping: {{"type":"set_mapping","target_header":"<exact target header text>","value":"<source column, __LITERAL__:x, or special marker>"}}
- set_target_headers: {{"type":"set_target_headers","headers":["Header One*","Header Two*", ...]}} (full replacement list)
- add_transformation: {{"type":"add_transformation","column":"<target header>","op":{{"type":"<tx op type>", ...op params}}}}
- remove_transformation: {{"type":"remove_transformation","column":"<target header>","op_index":0}}
- navigate: {{"type":"navigate","mode":"std|sk","panel":1-5}}
- none: {{"type":"none"}} — use when no state change is needed (pure Q&A)

Only ever use column names that exist in target_headers or source_columns above — never invent one. \
Keep "actions" as an empty array when you're just answering a question. Keep replies concise and \
concrete; say exactly what you changed (e.g. which header, which op) rather than being vague."""


@banca_bp.route('/api/ai/status')
def ai_status():
    return jsonify({'configured': bool(GROQ_API_KEY), 'model': GROQ_MODEL})


@banca_bp.route('/api/ai/chat', methods=['POST'])
def ai_chat():
    try:
        if not GROQ_API_KEY:
            return jsonify({'error': 'AI Assistant is not configured. Ask your admin to set the '
                                      'GROQ_API_KEY environment variable on the server (free key at '
                                      'console.groq.com/keys), then restart it.'}), 400
        d = request.json or {}
        user_msg = (d.get('message') or '').strip()
        if not user_msg:
            return jsonify({'error': 'Empty message'}), 400
        history = d.get('history') or []
        context = d.get('context') or {}
        mode = context.get('mode', 'std')

        if mode == 'sk':
            has_config = bool(context.get('sk_has_config'))
        else:
            has_config = bool(context.get('target_headers')) and bool(context.get('column_mapping'))

        messages = [{'role': 'system', 'content': build_ai_system_prompt(context, has_config)}]
        for h in history[-10:]:
            if h.get('role') in ('user', 'assistant') and h.get('content'):
                messages.append({'role': h['role'], 'content': str(h['content'])[:4000]})
        messages.append({'role': 'user', 'content': user_msg})

        resp = requests.post(GROQ_API_URL,
            headers={'Authorization': f'Bearer {GROQ_API_KEY}', 'Content-Type': 'application/json'},
            json={'model': GROQ_MODEL, 'messages': messages, 'temperature': 0.2,
                  'response_format': {'type': 'json_object'}},
            timeout=60)
        if resp.status_code >= 400:
            return jsonify({'error': f'Groq API error ({resp.status_code}): {resp.text[:300]}'}), 502
        data = resp.json()
        raw = data['choices'][0]['message']['content']
        try:
            parsed = json.loads(raw)
        except Exception:
            parsed = {'reply': raw, 'actions': []}

        actions = parsed.get('actions') or []
        if not isinstance(actions, list):
            actions = []
        actions = [a for a in actions if isinstance(a, dict) and a.get('type') in AI_ACTION_TYPES]
        # Hard-enforce the config gate server-side too, don't just trust the model.
        if not has_config:
            actions = [a for a in actions if a.get('type') in ('navigate', 'none')]

        return jsonify({'reply': parsed.get('reply', ''), 'actions': actions, 'has_config': has_config})
    except requests.exceptions.RequestException as e:
        return jsonify({'error': f'Could not reach Groq API: {e}'}), 502
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

# ─── API Routes ───────────────────────────────────────────────────────────────
@banca_bp.route('/')
def index(): return render_template('banca_index.html')

@banca_bp.route('/api/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files: return jsonify({'error':'No file'}),400
        f=request.files['file']
        if not f.filename or not allowed_file(f.filename): return jsonify({'error':'Invalid type'}),400
        sid=str(uuid.uuid4())[:8]
        fn=secure_filename(f.filename)
        sp=os.path.join(UPLOAD_DIR,f"{sid}_{fn}"); f.save(sp)
        rp=None
        if 'rto_file' in request.files:
            rf=request.files['rto_file']
            if rf.filename and allowed_file(rf.filename):
                rp=os.path.join(UPLOAD_DIR,f"{sid}_rto_{secure_filename(rf.filename)}"); rf.save(rp)
        sheets=[]
        try:
            ext=fn.rsplit('.',1)[1].lower()
            if ext=='csv': sheets=['default']
            elif ext=='xlsb': sheets=pd.ExcelFile(sp,engine='pyxlsb').sheet_names
            else: sheets=pd.ExcelFile(sp).sheet_names
        except: pass
        return jsonify({'session_id':sid,'filename':fn,'filepath':sp,'rto_filepath':rp,'sheets':sheets})
    except Exception as e: return jsonify({'error':str(e)}),500

@banca_bp.route('/api/preview', methods=['POST'])
def preview_columns():
    try:
        d=request.json
        sdf=read_excel_file(d['filepath'],d.get('sheet_name','Sheet1'),
                            int(d.get('header_row',2)),int(d.get('start_col',2)))
        cols=[c for c in sdf.columns if c and c.lower()!='nan']
        dcs=list(dict.fromkeys([c for c in cols for n,_,_ in DETRIFF_RANGES if c==n]))
        return jsonify({'columns':cols,'sample':sdf.head(3).fillna('').astype(str).to_dict('records'),
                        'detriff_cols':dcs,'has_detriff':len(dcs)>0,'row_count':len(sdf)})
    except Exception as e: return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

@banca_bp.route('/api/process', methods=['POST'])
def process():
    try:
        d=request.json
        sdf=read_excel_file(d['filepath'],d.get('sheet_name','Sheet1'),
                            int(d.get('header_row',2)),int(d.get('start_col',2)))
        rtom=build_rto_mapping(d['rto_filepath']) if d.get('rto_filepath') and os.path.exists(d['rto_filepath']) else {}
        # Detect source-only columns referenced in conditional if_col — carry them
        # through row expansion inside process_grid so row counts always match,
        # then drop them after transformation so they never appear in the output.
        real_target_headers = d.get('mapping_config',{}).get('target_headers',[])
        passthrough_cols = []
        if d.get('transformations'):
            for t in d['transformations']:
                for op in t.get('ops', []):
                    if op.get('type') == 'conditional':
                        for rule in op.get('rules', []):
                            ic = rule.get('if_col')
                            if ic and ic not in passthrough_cols:
                                passthrough_cols.append(ic)

        rdf=process_grid(sdf,d.get('mapping_config',{}),rtom,passthrough_cols=passthrough_cols)
        if rdf is None or len(rdf)==0: return jsonify({'error':'No data produced'}),400

        if d.get('transformations'):
            rdf = apply_transformations(rdf, d['transformations'])
            # Strip passthrough cols — keep only real target headers in output
            if real_target_headers:
                rdf = rdf[[c for c in real_target_headers if c in rdf.columns]]

        fn=f"{d.get('session_id','x')}_{secure_filename(d.get('output_name','output'))}.csv"
        op=os.path.join(OUTPUT_DIR,fn); rdf.to_csv(op,index=False)
        return jsonify({'success':True,'output_path':op,'output_filename':fn,'rows':len(rdf),
                        'cols':len(rdf.columns),'preview':rdf.head(5).fillna('').astype(str).to_dict('records'),
                        'columns':list(rdf.columns)})
    except Exception as e: return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

@banca_bp.route('/api/sk_finance/preview', methods=['POST'])
def sk_preview():
    try:
        d=request.json; fp=d.get('filepath'); sh=d.get('sheet_name','SK Finance')
        if not fp or not os.path.exists(fp): return jsonify({'error':'File not found'}),400
        wb=load_workbook(fp,data_only=True)
        if sh not in wb.sheetnames: return jsonify({'error':f"Sheet '{sh}' not found",'sheets':wb.sheetnames}),400
        ws=wb[sh]
        snap=[]
        for r in range(1,min(16,ws.max_row+1)):
            snap.append([str(ws.cell(row=r,column=c).value or '') for c in range(1,min(15,ws.max_column+1))])
        return jsonify({'sheets':wb.sheetnames,'max_row':ws.max_row,'max_col':ws.max_column,'snapshot':snap})
    except Exception as e: return jsonify({'error':str(e)}),500

@banca_bp.route('/api/sk_finance/process', methods=['POST'])
def sk_process():
    try:
        d=request.json
        cfg=d.get('config',{}); cfg['filepath']=d['filepath']
        rtom=build_rto_mapping(d['rto_filepath']) if d.get('rto_filepath') and os.path.exists(d['rto_filepath']) else {}
        rdf,meta=process_sk_finance(cfg,rtom)
        if d.get('transformations'): rdf=apply_transformations(rdf,d['transformations'])
        fn=f"{d.get('session_id','x')}_{secure_filename(d.get('output_name','sk_output'))}.csv"
        op=os.path.join(OUTPUT_DIR,fn); rdf.to_csv(op,index=False)
        return jsonify({'success':True,'output_path':op,'output_filename':fn,'rows':len(rdf),
                        'cols':len(rdf.columns),'meta':{k:str(v) for k,v in meta.items()},
                        'preview':rdf.head(5).fillna('').astype(str).to_dict('records'),
                        'columns':list(rdf.columns)})
    except Exception as e: return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

@banca_bp.route('/api/merge_detriff', methods=['POST'])
def merge_detriff():
    try:
        d=request.json; ip=d.get('output_path')
        if not ip or not os.path.exists(ip): return jsonify({'error':'File not found'}),400
        df=pd.read_csv(ip); mdf,red=apply_detriff_merger(df)
        mfn=os.path.basename(ip).replace('.csv','')+'_merged.csv'
        mp=os.path.join(OUTPUT_DIR,mfn); mdf.to_csv(mp,index=False)
        return jsonify({'success':True,'merged_path':mp,'merged_filename':mfn,
                        'original_rows':len(df),'merged_rows':len(mdf),'rows_reduced':red,
                        'preview':mdf.head(5).fillna('').astype(str).to_dict('records')})
    except Exception as e: return jsonify({'error':str(e)}),500

@banca_bp.route('/api/download/<filename>')
def download(filename):
    p=os.path.join(OUTPUT_DIR,filename)
    if not os.path.exists(p): return jsonify({'error':'Not found'}),404
    return send_file(p,as_attachment=True,download_name=filename)

@banca_bp.route('/api/sk_finance/default_config')
def sk_default_config():
    return jsonify({'agent_code_cell':[1,2],'relationship_code_cell':[3,2],
        'payment_basis_row':9,'ncb_row':10,'cc_row':11,'lob_row':12,
        'data_start_row':14,'state_col':1,'cluster_col':2,'data_start_col':3,
        'valid_payment_basis':['on OD','on TP'],'version_id':'',
        'output_columns':['version_id','pol_agent_code','primary_agent_code','product_code','proposal_id',
            'biz_mix','two_wheeler_category','type_of_business','oem_type','final_region',
            'rto_code','detriff_discount_ll','detriff_discount_ul','gross_vehicle_weight_ll',
            'gross_vehicle_weight_ul','cubic_capacity_ll','cubic_capacity_ul','bus_type',
            'model','fuel_type','banca_cpa_flag','banca_misp','banca_flat_amount',
            'banca_capping','banca_tp_prct','banca_tp_outgo','banca_total_comm_rate',
            'banca_outgo','product_desc','state','ncb_percentage_ll','ncb_percentage_ul',
            'body_type','txt_motor_nonmotor'],
        'output_mapping':_default_sk_mapping(),
        'available_source_keys':['__agent_code__','__relationship_code__','__state__','__cluster__',
            '__rto_codes__','__comm_rate__','__tp_rate__','__cc_ll__','__cc_ul__','__ncb_ll__','__ncb_ul__',
            '__banca_outgo__','__banca_tp_prct__','__banca_tp_outgo__',
            '__lob__','__payment_basis__','__ncb__','__cc__','__version_id__']})

@banca_bp.route('/api/presets')
def get_presets():
    return jsonify({
        "TATA PV":{"sheet_name":"Sheet2","header_row":2,"start_col":2,
            "target_headers":["Version Id*","Oem Type*","Pol Agent Code*","Primary Agent Code*","Supplier Code*",
                "Vehicle Type*","Add On Premium Flag*","Biz Mix*","Model*","Two Wheeler Category*",
                "Type Of Business*","Tata Vehicle Age Ll*","Tata Vehicle Age Ul*","Nop Ll*","Nop Ul*",
                "Total Gwp Ll*","Total Gwp Ul*","Detriff Discount Ll*","Detriff Discount Ul*",
                "Cubic Capacity Ll*","Cubic Capacity Ul*","Gross Vehicle Weight Ll*","Gross Vehicle Weight Ul*",
                "Tatacvpv Misp*","Flat Amnt*","Outgo*","Tatacvpv Capping*","Tatacvpv Comm Prct*",
                "Od Rate*","Tp Rate*","Tata Provision Rate*","Tata Provision Outgo*","Tata Fuel Type*",
                "Tata Identifire*","Rto Code*","Unique Id"],
            "column_mapping":{"Version Id*":"__LITERAL__:tatacvpv_Grid_3","Oem Type*":"__LITERAL__:TMIBASL",
                "Pol Agent Code*":"__LITERAL__:BRC0000286","Primary Agent Code*":"__LITERAL__:CVPV01234",
                "Supplier Code*":"Dealer Name","Vehicle Type*":"VEHICLETYPE","Add On Premium Flag*":"__LITERAL__:No",
                "Biz Mix*":"LOB","Model*":"Model","Type Of Business*":"Old/New",
                "Tata Vehicle Age Ll*":"Vehicle Age","Tata Vehicle Age Ul*":"Vehicle Age",
                "Nop Ll*":"NCB (Y/N)","Nop Ul*":"NCB (Y/N)",
                "Total Gwp Ll*":"__LITERAL__:-10000000","Total Gwp Ul*":"__LITERAL__:10000000",
                "Tata Fuel Type*":"Fuel Type","Tatacvpv Misp*":"System CommissionExtra",
                "Tatacvpv Capping*":"Capping Maximum Limit %","Outgo*":"Payment Type",
                "Tatacvpv Comm Prct*":"__DETRIFF_VALUE__","Tata Identifire*":"__LITERAL__:TATA PV",
                "Rto Code*":"__RTO_CODES__","Unique Id":"__UNIQUE_ID__"},
            "rto_cluster_source":"UW Budget Cluster"},
        "TATA CV":{"sheet_name":"Sheet1","header_row":2,"start_col":2,
            "target_headers":["Version Id*","Oem Type*","Pol Agent Code*","Primary Agent Code*","Supplier Code*",
                "Vehicle Type*","Add On Premium Flag*","Biz Mix*","Model*","Two Wheeler Category*",
                "Type Of Business*","Tata Vehicle Age Ll*","Tata Vehicle Age Ul*","Nop Ll*","Nop Ul*",
                "Total Gwp Ll*","Total Gwp Ul*","Detriff Discount Ll*","Detriff Discount Ul*",
                "Cubic Capacity Ll*","Cubic Capacity Ul*","Gross Vehicle Weight Ll*","Gross Vehicle Weight Ul*",
                "Tatacvpv Misp*","Flat Amnt*","Outgo*","Tatacvpv Capping*","Tatacvpv Comm Prct*",
                "Od Rate*","Tp Rate*","Tata Provision Rate*","Tata Provision Outgo*","Tata Fuel Type*",
                "Tata Identifire*","Rto Code*","Unique Id"],
            "column_mapping":{"Version Id*":"__LITERAL__:tatacvpv_Grid_3","Oem Type*":"__LITERAL__:TATA",
                "Pol Agent Code*":"__LITERAL__:BRC0000286","Primary Agent Code*":"__LITERAL__:CVPV01234",
                "Supplier Code*":"Dealer Name","Biz Mix*":"LOB","Model*":"Model","Type Of Business*":"Old/New",
                "Tata Vehicle Age Ll*":"Vehicle Age","Tata Vehicle Age Ul*":"Vehicle Age",
                "Nop Ll*":"NCB (Y/N)","Nop Ul*":"NCB (Y/N)",
                "Total Gwp Ll*":"__LITERAL__:-10000000","Total Gwp Ul*":"__LITERAL__:10000000",
                "Tata Fuel Type*":"Fuel Type","Tatacvpv Misp*":"System CommissionExtra",
                "Tatacvpv Capping*":"Capping Maximum Limit %","Outgo*":"Payment Type",
                "Tatacvpv Comm Prct*":"__DETRIFF_VALUE__","Tata Identifire*":"__LITERAL__:TATA CV",
                "Rto Code*":"__RTO_CODES__","Unique Id":"__UNIQUE_ID__"},
            "rto_cluster_source":"UW Budget Cluster"},
        "Banca Motor":{"sheet_name":"Sheet1","header_row":2,"start_col":2,
            "target_headers":["Version Id*","Parent Agent Code*","Product Code*","Proposal Id*","Biz Mix*",
                "Two Wheeler Category*","Type Of Business*","Oem Type*","Final Region*",
                "Detriff Discount Ll*","Detriff Discount Ul*","Gross Vehicle Weight Ll*","Gross Vehicle Weight Ul*",
                "Cubic Capacity Ll*","Cubic Capacity Ul*","Primary Agent Code*","Total Gwp Ll*","Total Gwp Ul*",
                "Make*","Fuel Type*","Banca Misp*","Banca Flat Amount*","Tp Prct*","Banca Total Comm Rate*",
                "Banca Outgo*","Irda Flag*","Banca Capping*","Capping Outgo*","Model*","Banca Cpa Flag*",
                "Vehicle Class*","Bus Type*","Tp Outgo*","Unique Id","Special Remarks","Remarks","Rto Code*"],
            "column_mapping":{"Version Id*":"__LITERAL__:banca_oem_july_Grid_2",
                "Parent Agent Code*":"IMD Code (If not based on relationship code)","Biz Mix*":"LOB",
                "Vehicle Class*":"Vehicle Class","Two Wheeler Category*":"Bike / Scooter","Bus Type*":"Bus Type",
                "Type Of Business*":"Old/New","Cubic Capacity Ll*":"CC Lower Limit","Cubic Capacity Ul*":"CC Upper Limit",
                "Primary Agent Code*":"Relationship Code","Make*":"Make","Fuel Type*":"Fuel Type",
                "Banca Misp*":"System CommissionExtra","Tp Prct*":"TP Comm%","Tp Outgo*":"TP Comm Basis",
                "Banca Outgo*":"Payment Type","Banca Capping*":"Capping % if Any","Model*":"Model",
                "Banca Total Comm Rate*":"__DETRIFF_VALUE__","Special Remarks":"Special Remarks",
                "Remarks":"Remarks","Rto Code*":"__RTO_CODES__","Unique Id":"__UNIQUE_ID__"},
            "rto_cluster_source":"UW Budget Cluster"},
        "Hyundai":{"sheet_name":"Sheet1","header_row":2,"start_col":2,
            "target_headers":["Version Id*","Oem Type*","Parent Agent Code*","Supplier Code*","Primary Agent Code*",
                "Biz Mix*","Two Wheeler Category*","Type Of Business*","Detriff Discount Ll*","Detriff Discount Ul*",
                "Cubic Capacity Ll*","Cubic Capacity Ul*","Gross Vehicle Weight Ll*","Gross Vehicle Weight Ul*",
                "Misp Hyundai*","Outgo Hyundai*","Comm Percentage Hyundai*","Tp Irda Flag*","Unique Id"],
            "column_mapping":{"Version Id*":"__LITERAL__:magma_hyundai_oem_Grid_1","Oem Type*":"OEM",
                "Supplier Code*":"Dealer Code","Primary Agent Code*":"Relationship Code","Biz Mix*":"LOB",
                "Type Of Business*":"Old/New","Misp Hyundai*":"System CommissionExtra",
                "Outgo Hyundai*":"Payment Type","Comm Percentage Hyundai*":"__DETRIFF_VALUE__",
                "Tp Irda Flag*":"TP Comm","Unique Id":"__UNIQUE_ID__"}},
        "Nissan / Renault":{"sheet_name":"Sheet1","header_row":2,"start_col":2,
            "target_headers":["Version Id*","Parent Agent Code*","Supplier Code*","Primary Agent Code*",
                "Channel Code Parent Txn*","Subchannel Code Parent Txn*","Biz Mix*","Two Wheeler Category*",
                "Type Of Business*","Region State*","Rto State*","Region For Calculation*","Misp*",
                "Budget Cluster*","Capping*","Capping Outgo*","Detriff Discount Ll*","Detriff Discount Ul*",
                "Cubic Capacity Ll*","Cubic Capacity Ul*","Gross Vehicle Weight Ll*","Gross Vehicle Weight Ul*",
                "Outgo Nissan*","Comm Percentage*","Fuel Type*","No Claim Bonus Ll*","No Claim Bonus Ul*",
                "Oem*","RTO Codes*","Unique Id"],
            "column_mapping":{"Version Id*":"__LITERAL__:magma_nissan_oem_Grid_1",
                "Supplier Code*":"Dealer Code Nissan_Renault (Actual)","Primary Agent Code*":"Relationship Code",
                "Channel Code Parent Txn*":"IMD Code (If not based on relationship code)","Biz Mix*":"LOB",
                "Type Of Business*":"Old/New","Region State*":"Zone","Budget Cluster*":"UW Budget Cluster",
                "Rto State*":"State","Misp*":"System CommissionExtra","Capping*":"Capping Maximum Limit %",
                "Capping Outgo*":"Capping Basis","Cubic Capacity Ll*":"CC Lower Limit","Cubic Capacity Ul*":"CC Upper Limit",
                "Outgo Nissan*":"Payment Type","Comm Percentage*":"__DETRIFF_VALUE__","Fuel Type*":"Fuel Type",
                "No Claim Bonus Ll*":"NCB (Y/N)","Oem*":"OEM","RTO Codes*":"__RTO_CODES__","Unique Id":"__UNIQUE_ID__"},
            "rto_cluster_source":"UW Budget Cluster"},
        "OEM Other":{"sheet_name":"Sheet1","header_row":2,"start_col":2,
            "target_headers":["Version Id*","Primary Agent Code*","Make*","Vehicle Type*","Biz Mix*",
                "Detriff Discount Ll*","Detriff Discount Ul*","Cubic Capacity Ll*","Cubic Capacity Ul*",
                "Gross Vehicle Weight Ll*","Gross Vehicle Weight Ul*","NCB*","Two Wheeler Category*",
                "Type Of Business*","Fuel Type*","Oem Type*","Bus Type*","Rto Cluster*","RTO Codes*",
                "Oem Misp*","Oem Other Capping*","Oem Other Capping Amnt*","Oem Outgo*","Oem Flat Amount*",
                "Oem Comm Prct*","Oem Other Cpa Flag*","Vehicle Age Ll*","Vehicle Age Ul*","Unique Id"],
            "column_mapping":{"Version Id*":"__LITERAL__:oem_other_Grid_2","Primary Agent Code*":"Relationship Code",
                "Make*":"Make","Vehicle Type*":"Vehicle Class","Biz Mix*":"LOB",
                "Cubic Capacity Ll*":"CC Lower Limit","Cubic Capacity Ul*":"CC Upper Limit","NCB*":"NCB (Y/N)",
                "Two Wheeler Category*":"Two Wheeler Type","Type Of Business*":"Old/New","Fuel Type*":"Fuel Type",
                "Oem Type*":"OEM","Bus Type*":"Bus / Vehicle Type","Rto Cluster*":"UW Budget Cluster",
                "Oem Misp*":"System CommissionExtra","Oem Other Capping*":"Capping Maximum Limit %",
                "Oem Outgo*":"Payment Type","Oem Comm Prct*":"__DETRIFF_VALUE__","Oem Other Cpa Flag*":"CPA Flag",
                "Vehicle Age Ll*":"Vehicle Age","Vehicle Age Ul*":"Vehicle Age",
                "RTO Codes*":"__RTO_CODES__","Unique Id":"__UNIQUE_ID__"},
            "rto_cluster_source":"UW Budget Cluster"}
    })

