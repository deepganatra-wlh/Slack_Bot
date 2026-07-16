"""
GridBot — Slack bot that drives the two grid-processing portals
(Standard/Banca Grid Processor + Agency Special Motor Matrix) from a Slack
mention, using the same JSON config the portals' "Export Config" button produces.

Usage in Slack:
    @gridbot use banca portal
    (attach: grid.xlsx, config.json, optionally rto.xlsx)

    @gridbot use agency portal
    (attach: matrix.xlsx, rto.xlsx, config.json)

    @gridbot use sk finance
    (attach: grid.xlsx, config.json)   -- uses the "sk" section of a banca-portal config export

Architecture note: this Blueprint is mounted in the SAME Flask app as the two
portal Blueprints (see main.py). Rather than making real HTTP calls to itself
(which would count against PythonAnywhere free-tier outbound whitelist /
CPU-second limits), it drives the portals in-process via Flask's test client —
functionally identical to what the browser does, just without a network hop.
The only real outbound HTTP in this file is to Slack's own API.
"""

import os
import re
import io
import json
import time
import hmac
import hashlib
import logging
import tempfile
import threading

import requests
from flask import Blueprint, request, jsonify, current_app
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger('gridbot')

# ── Config ──────────────────────────────────────────────────────────────────
SLACK_BOT_TOKEN      = os.environ.get('SLACK_BOT_TOKEN', '')       # xoxb-...
SLACK_SIGNING_SECRET = os.environ.get('SLACK_SIGNING_SECRET', '')  # from Basic Information page

# Mount points of the two portal Blueprints within THIS SAME app (see main.py).
# Only change these if you re-register the blueprints under different prefixes.
BANCA_PREFIX  = '/banca'
AGENCY_PREFIX = '/agency'

if not SLACK_BOT_TOKEN or not SLACK_SIGNING_SECRET:
    log.warning('SLACK_BOT_TOKEN / SLACK_SIGNING_SECRET not set — the bot will not be able to '
                'verify requests or talk to Slack until these are exported as env vars.')

slack_bp = Blueprint('slackbot', __name__)
slack = WebClient(token=SLACK_BOT_TOKEN)

_seen_event_ids = {}   # simple in-memory de-dup for Slack's at-least-once delivery / retries
_SEEN_TTL = 600


# ══════════════════════════════════════════════════════════════════════════
# Slack request verification
# ══════════════════════════════════════════════════════════════════════════

def verify_slack_signature(req) -> bool:
    if not SLACK_SIGNING_SECRET:
        return False
    timestamp = req.headers.get('X-Slack-Request-Timestamp', '')
    if not timestamp or abs(time.time() - int(timestamp)) > 60 * 5:
        return False  # too old — replay-attack guard
    sig_basestring = f"v0:{timestamp}:{req.get_data(as_text=True)}"
    my_sig = 'v0=' + hmac.new(
        SLACK_SIGNING_SECRET.encode(), sig_basestring.encode(), hashlib.sha256
    ).hexdigest()
    their_sig = req.headers.get('X-Slack-Signature', '')
    return hmac.compare_digest(my_sig, their_sig)


def _dedup_prune():
    now = time.time()
    for k, ts in list(_seen_event_ids.items()):
        if now - ts > _SEEN_TTL:
            _seen_event_ids.pop(k, None)


# ══════════════════════════════════════════════════════════════════════════
# Portal / mode detection from the mention text
# ══════════════════════════════════════════════════════════════════════════

def detect_target(text: str):
    """Returns ('std'|'agency', 'std'|'sk') based on keywords in the message."""
    t = (text or '').lower()
    if any(k in t for k in ['agency', 'special motor', 'matrix']):
        return 'agency', None
    if 'sk finance' in t or re.search(r'\bsk\b', t):
        return 'std', 'sk'
    if any(k in t for k in ['banca', 'standard grid', 'grid processor', 'std']):
        return 'std', 'std'
    return None, None


def wants_detriff_merge(text: str) -> bool:
    return 'detriff' in (text or '').lower()


# ══════════════════════════════════════════════════════════════════════════
# Slack file download / upload helpers
# ══════════════════════════════════════════════════════════════════════════

def download_slack_file(file_info: dict, dest_dir: str) -> str:
    url = file_info['url_private_download']
    name = file_info.get('name', 'file')
    path = os.path.join(dest_dir, name)
    r = requests.get(url, headers={'Authorization': f'Bearer {SLACK_BOT_TOKEN}'}, timeout=120)
    r.raise_for_status()
    with open(path, 'wb') as f:
        f.write(r.content)
    return path


def classify_files(files: list, dest_dir: str) -> dict:
    """Splits attached Slack files into grid / rto / config by extension + filename hints."""
    out = {'grid': None, 'rto': None, 'config': None}
    excel_like = []
    for finfo in files:
        name = (finfo.get('name') or '').lower()
        path = download_slack_file(finfo, dest_dir)
        if name.endswith('.json'):
            out['config'] = path
        elif name.endswith(('.xlsx', '.xlsb', '.xls', '.csv')):
            excel_like.append((name, path))
        # anything else (images, pdfs) is ignored

    for name, path in excel_like:
        if 'rto' in name and out['rto'] is None:
            out['rto'] = path
    for name, path in excel_like:
        if out['grid'] is None and path != out['rto']:
            out['grid'] = path
    return out


def post(channel, thread_ts, text):
    try:
        slack.chat_postMessage(channel=channel, thread_ts=thread_ts, text=text)
    except SlackApiError as e:
        log.error(f'Slack post failed: {e}')


def post_file(channel, thread_ts, filepath, title, comment):
    try:
        slack.files_upload_v2(
            channel=channel, thread_ts=thread_ts,
            file=filepath, title=title, initial_comment=comment,
        )
    except SlackApiError as e:
        log.error(f'Slack file upload failed: {e}')
        post(channel, thread_ts, f"⚠ Processed successfully but couldn't upload the result file: {e}")


# ══════════════════════════════════════════════════════════════════════════
# Internal helpers for calling the portal Blueprints via the Flask test client
# ══════════════════════════════════════════════════════════════════════════

def _upload_files(client, prefix, grid_path, rto_path):
    data = {'file': (open(grid_path, 'rb'), os.path.basename(grid_path))}
    if rto_path:
        data['rto_file'] = (open(rto_path, 'rb'), os.path.basename(rto_path))
    try:
        resp = client.post(f'{prefix}/api/upload', data=data, content_type='multipart/form-data')
    finally:
        data['file'][0].close()
        if rto_path:
            data['rto_file'][0].close()
    ud = resp.get_json() or {}
    if resp.status_code >= 400 or ud.get('error'):
        raise RuntimeError(ud.get('error', f'Upload failed ({resp.status_code})'))
    return ud


def _download_output(client, prefix, filename):
    resp = client.get(f'{prefix}/api/download/{filename}')
    if resp.status_code >= 400:
        raise RuntimeError(f'Could not download output file ({resp.status_code})')
    local_out = os.path.join(tempfile.gettempdir(), filename)
    with open(local_out, 'wb') as f:
        f.write(resp.data)
    return local_out


# ══════════════════════════════════════════════════════════════════════════
# Portal drivers — replay the same calls the browser UI makes, in-process
# ══════════════════════════════════════════════════════════════════════════

def run_std_portal(client, grid_path, rto_path, cfg: dict, sub_mode: str, merge_detriff: bool = False):
    """Drives the Banca / Standard Grid Processor Portal Blueprint."""
    ud = _upload_files(client, BANCA_PREFIX, grid_path, rto_path)

    if sub_mode == 'sk':
        k = cfg.get('sk', {})
        try:
            out_map = json.loads(k.get('out_map') or '{}')
        except Exception as e:
            raise RuntimeError(f'Config JSON has an invalid sk.out_map field: {e}')
        out_cols = [c.strip() for c in (k.get('out_cols') or '').split('\n') if c.strip()]
        sk_config = {
            'sheet_name': k.get('sheet_name', 'SK Finance'),
            'version_id': k.get('version_id', ''),
            'agent_code_cell': [int(k.get('acc_r', 1) or 1), int(k.get('acc_c', 2) or 2)],
            'relationship_code_cell': [int(k.get('rel_r', 3) or 3), int(k.get('rel_c', 2) or 2)],
            'payment_basis_row': int(k.get('pb_row', 0) or 0),
            'ncb_row': int(k.get('ncb_row', 0) or 0),
            'cc_row': int(k.get('cc_row', 0) or 0),
            'lob_row': int(k.get('lob_row', 0) or 0),
            'data_start_row': int(k.get('ds_row', 0) or 0),
            'state_col': int(k.get('st_col', 0) or 0),
            'cluster_col': int(k.get('cl_col', 0) or 0),
            'data_start_col': int(k.get('dc_col', 0) or 0),
            'valid_payment_basis': [v.strip() for v in (k.get('vpb') or '').split(',') if v.strip()],
            'output_columns': out_cols,
            'output_mapping': out_map,
        }
        payload = {
            'filepath': ud['filepath'], 'rto_filepath': ud.get('rto_filepath'),
            'config': sk_config,
            'transformations': k.get('transformations', []),
            'output_name': k.get('out_name', 'sk_output'),
            'session_id': ud['session_id'],
        }
        endpoint = f'{BANCA_PREFIX}/api/sk_finance/process'
    else:
        s = cfg.get('std', {})
        target_hdrs = [h.strip() for h in (s.get('target_hdrs') or '').split('\n') if h.strip()]
        try:
            column_mapping = json.loads(s.get('map_json') or '{}')
        except Exception as e:
            raise RuntimeError(f'Config JSON has an invalid map_json field: {e}')
        payload = {
            'filepath': ud['filepath'], 'rto_filepath': ud.get('rto_filepath'),
            'sheet_name': s.get('sheet_name', ''),
            'header_row': int(s.get('header_row', 2) or 2),
            'start_col': int(s.get('start_col', 2) or 2),
            'mapping_config': {
                'target_headers': target_hdrs,
                'column_mapping': column_mapping,
                'rto_cluster_source': s.get('rto_src', 'UW Budget Cluster'),
            },
            'transformations': s.get('transformations', []),
            'output_name': s.get('output_name', 'output'),
            'session_id': ud['session_id'],
        }
        endpoint = f'{BANCA_PREFIX}/api/process'

    resp = client.post(endpoint, json=payload)
    pd = resp.get_json() or {}
    if resp.status_code >= 400 or pd.get('error'):
        raise RuntimeError(pd.get('trace') or pd.get('error', f'Processing failed ({resp.status_code})'))

    if merge_detriff and pd.get('output_path'):
        mresp = client.post(f'{BANCA_PREFIX}/api/merge_detriff', json={'output_path': pd['output_path']})
        mpd = mresp.get_json() or {}
        if mresp.status_code >= 400 or mpd.get('error'):
            raise RuntimeError(mpd.get('error', f'Detriff merge failed ({mresp.status_code})'))
        pd['merged_filename'] = mpd['merged_filename']
        pd['merged_rows'] = mpd['merged_rows']
        pd['rows_reduced'] = mpd['rows_reduced']
        local_out = _download_output(client, BANCA_PREFIX, mpd['merged_filename'])
    else:
        local_out = _download_output(client, BANCA_PREFIX, pd['output_filename'])
    return local_out, pd


def run_agency_portal(client, grid_path, rto_path, cfg: dict):
    """Drives the Agency Special Motor Matrix Portal Blueprint."""
    ud = _upload_files(client, AGENCY_PREFIX, grid_path, rto_path)

    meta_col_map = {
        'imd_code':   {'col_idx': int(cfg.get('mc_imd_code', 0) or 0)},
        'imd_name':   {'col_idx': int(cfg.get('mc_imd_name', 0) or 0)},
        'rel_code':   {'col_idx': int(cfg.get('mc_rel', 0) or 0)},
        'imd_type':   {'col_idx': int(cfg.get('mc_imd_type', 0) or 0)},
        'vol_ll':     {'col_idx': int(cfg.get('mc_vol_ll', 0) or 0)},
        'vol_ul':     {'col_idx': int(cfg.get('mc_vol_ul', 0) or 0)},
        'vol_remark': {'col_idx': int(cfg.get('mc_vol_rem', 0) or 0)},
        'uw_cluster': {'col_idx': int(cfg.get('mc_cluster', 0) or 0)},
    }
    extra_meta_cols = [e for e in cfg.get('extraMetaCols', []) if e.get('label') and int(e.get('col_idx', 0) or 0) > 0]

    vol_gwp_map = {r['vol_rem']: {'ll_col': r.get('ll_col', ''), 'ul_col': r.get('ul_col', '')}
                    for r in cfg.get('vgRows', []) if r.get('vol_rem')}
    agent_group_map = {r['imd_type']: r.get('code', '') for r in cfg.get('agentRows', []) if r.get('imd_type')}
    rto_norm_map = {r['from'].upper(): r.get('to', '') for r in cfg.get('normRows', []) if r.get('from')}
    output_static = {r['col']: r.get('val', '') for r in cfg.get('staticFields', []) if r.get('col')}
    column_defaults = {r['col']: r.get('val', '') for r in cfg.get('colDefaults', []) if r.get('col')}
    col_defs = [{'col_idx': c['col_idx'], 'biz_mix_output': c.get('biz_mix_output', ''),
                 'rto_category': c.get('rto_category', ''), 'extra_fields': c.get('extra_fields', {})}
                for c in cfg.get('colCfg', []) if c.get('enabled') and c.get('col_idx')]
    imd_gwp_map = {r['imd_type']: {'ll_col': r.get('ll_col', ''), 'ul_col': r.get('ul_col', '')}
                   for r in cfg.get('imdGwpRows', []) if r.get('imd_type')}

    ignore_values = [v.strip() for v in (cfg.get('ignore_values') or '').split('\n') if v.strip()]
    skip_vol_biz = [v.strip() for v in (cfg.get('skip_vol_biz') or '').split(',') if v.strip()]

    config = {
        'sheet_name': cfg.get('sheet_name', ''),
        'header_rows': [int(cfg.get('header_row1', 2) or 2), int(cfg.get('header_row2', 3) or 3),
                         int(cfg.get('header_row3', 4) or 4)],
        'data_start_row': int(cfg.get('data_start_row', 5) or 5),
        'meta_col_map': meta_col_map, 'col_defs': col_defs,
        'mode': cfg.get('mode', 'special'), 'extra_meta_cols': extra_meta_cols,
        'ignore_values': ignore_values, 'irda_values': ignore_values,
        'irda_prct_value': cfg.get('irda_prct', '-0.1'),
        'irda_outgo_value': cfg.get('irda_outgo', 'IRDA'),
        'normal_outgo_value': cfg.get('norm_outgo', 'GWP'),
        'agent_group_map': agent_group_map, 'vol_gwp_map': vol_gwp_map,
        'std_gwp_ll_col': cfg.get('std_ll', 'Total Gwp Ll*'), 'std_gwp_ul_col': cfg.get('std_ul', 'Total Gwp Ul*'),
        'prime_gwp_ll_col': cfg.get('prime_ll', 'Total Gwp Ll*'), 'prime_gwp_ul_col': cfg.get('prime_ul', 'Total Gwp Ul*'),
        'agency_gwp_ll_col': cfg.get('agency_ll', ''), 'agency_gwp_ul_col': cfg.get('agency_ul', ''),
        'keybrok_gwp_ll_col': cfg.get('keybrok_ll', ''), 'keybrok_gwp_ul_col': cfg.get('keybrok_ul', ''),
        'imd_gwp_map': imd_gwp_map,
        'output_static_fields': output_static, 'column_defaults': column_defaults,
        'rto_norm_map': rto_norm_map, 'skip_if_vol_biz': skip_vol_biz,
        'span_outgo_col': cfg.get('out_span_outgo', 'Span Outgo*'),
        'span_prct_col': cfg.get('out_span_prct', 'Span Prct*'),
        'rto_code_col': cfg.get('out_rto_code', 'Rto Code*'),
        'rto_cluster_col': cfg.get('out_rto_clu', 'Rto Cluster*'),
        'parent_agent_col': cfg.get('out_parent', 'Parent Agent Code*'),
        'primary_agent_col': cfg.get('out_primary', 'Primary Agent Code*'),
        'agent_group_col': cfg.get('out_ag_grp', 'Agent Group Code*'),
        'biz_mix_col': cfg.get('out_biz_mix', 'Biz Mix*'),
    }

    payload = {
        'filepath': ud['filepath'], 'rto_filepath': ud.get('rto_filepath'),
        'session_id': ud['session_id'],
        'output_name': cfg.get('out_fn', 'agency_output'),
        'rto_sheet': cfg.get('rto_sheet', 'RTO Vs Cluster (New)'),
        'rto_header_row': int(cfg.get('rto_hr', 2) or 2),
        'rto_col': cfg.get('rto_col', 'RTO CODE'),
        'rto_cluster_col': cfg.get('rto_clu_col', 'UW CLUSTER (26-27)'),
        'rto_use_cat': bool(cfg.get('rto_use_cat', True)),
        'rto_cat_col': cfg.get('rto_cat_col') if cfg.get('rto_use_cat', True) else None,
        'col_transforms': [r for r in cfg.get('colTransforms', []) if r.get('col') and r.get('op') and r.get('value') != ''],
        'output_format': [r for r in cfg.get('outputFormat', []) if (r.get('col') or '').strip()],
        'config': config,
    }

    resp = client.post(f'{AGENCY_PREFIX}/api/process', json=payload)
    pd = resp.get_json() or {}
    if resp.status_code >= 400 or pd.get('error'):
        raise RuntimeError(pd.get('trace') or pd.get('error', f'Processing failed ({resp.status_code})'))

    local_out = _download_output(client, AGENCY_PREFIX, pd['output_filename'])
    return local_out, pd


# ══════════════════════════════════════════════════════════════════════════
# Mention handling (runs in a background thread so we can ACK Slack in <3s)
# ══════════════════════════════════════════════════════════════════════════

def handle_mention(event: dict, flask_app):
    channel = event['channel']
    thread_ts = event.get('thread_ts') or event['ts']
    text = event.get('text', '')
    files = event.get('files', [])

    portal, sub_mode = detect_target(text)
    merge_detriff = wants_detriff_merge(text)
    if not portal:
        post(channel, thread_ts,
             "I couldn't tell which portal to use. Say *\"use banca portal\"*, "
             "*\"use sk finance\"*, or *\"use agency portal\"* in your message.")
        return
    if not files:
        post(channel, thread_ts,
             "I need at least the grid file and your exported config JSON attached to this message "
             "(and the RTO file too, if that portal needs one).")
        return

    post(channel, thread_ts, f"⏳ Got it — running the *{'Agency' if portal=='agency' else ('SK Finance' if sub_mode=='sk' else 'Banca/Standard')}* "
                              f"portal on {len(files)} file(s)"
                              f"{' with detriff merge' if merge_detriff else ''}…")

    with tempfile.TemporaryDirectory() as tmp:
        try:
            classified = classify_files(files, tmp)
            if not classified['grid']:
                post(channel, thread_ts, "⚠ I couldn't find a grid Excel/CSV file in your attachments.")
                return
            if not classified['config']:
                post(channel, thread_ts, "⚠ I couldn't find a config `.json` file in your attachments — "
                                          "export one from the portal's sidebar (\"Export Config\") and attach it.")
                return
            with open(classified['config']) as f:
                cfg = json.load(f)

            # Run inside an app context so url_for / current_app work correctly,
            # and use the test client to call the sibling portal Blueprints in-process.
            with flask_app.app_context():
                client = flask_app.test_client()
                if portal == 'agency':
                    out_path, pd = run_agency_portal(client, classified['grid'], classified['rto'], cfg)
                else:
                    out_path, pd = run_std_portal(client, classified['grid'], classified['rto'], cfg, sub_mode,
                                                   merge_detriff=merge_detriff)

            summary = f"✅ Done — {pd.get('rows', '?')} rows × {pd.get('cols', '?')} columns."
            if pd.get('skipped_rows') or pd.get('skipped_cells'):
                summary += f" ({pd.get('skipped_rows', 0)} rows / {pd.get('skipped_cells', 0)} cells skipped)"
            if merge_detriff and pd.get('merged_rows') is not None:
                summary += f"\nDetriff merge: {pd['rows']} → {pd['merged_rows']} rows ({pd.get('rows_reduced', 0)} reduced)."
            post_file(channel, thread_ts, out_path, title=pd.get('merged_filename') or pd.get('output_filename', 'output.csv'), comment=summary)

        except requests.exceptions.RequestException as e:
            post(channel, thread_ts, f"⚠ Couldn't download an attached file from Slack: {e}")
        except Exception as e:
            log.exception('Processing failed')
            post(channel, thread_ts, f"⚠ Processing failed: {e}")


# ══════════════════════════════════════════════════════════════════════════
# Flask routes
# ══════════════════════════════════════════════════════════════════════════

@slack_bp.route('/events', methods=['POST'])
def slack_events():
    if not verify_slack_signature(request):
        return jsonify({'error': 'invalid signature'}), 401

    data = request.get_json(silent=True) or {}

    # Slack's one-time URL verification handshake when you first save the Request URL
    if data.get('type') == 'url_verification':
        return jsonify({'challenge': data.get('challenge', '')})

    event_id = data.get('event_id')
    _dedup_prune()
    if event_id:
        if event_id in _seen_event_ids:
            return '', 200  # already processed — Slack retried delivery
        _seen_event_ids[event_id] = time.time()

    event = data.get('event', {})
    if event.get('type') == 'app_mention':
        real_app = current_app._get_current_object()
        threading.Thread(target=handle_mention, args=(event, real_app), daemon=True).start()

    return '', 200  # ACK immediately; real work happens in the background thread


@slack_bp.route('/healthz')
def healthz():
    return jsonify({'status': 'ok'})