# GK_DASHBOARD_CORE_VERSION = 'v8_3_static_read_only_pending_2026_08_05'
import argparse
import base64
import html
import json
import re
import unicodedata
import zipfile
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable, List

import numpy as np
import pandas as pd
from plotly.offline.offline import get_plotlyjs

DEFAULT_INPUT_FILE = Path('data/Submitter_Tracking_Master_With_Supervisor.xlsx')
DEFAULT_OUTPUT_FILE = Path('official.html')
DEFAULT_SHEET_NAME = 0
DEFAULT_LOGO_HTML = Path('data/Milwaukee-logo.html')
DEFAULT_LOGO_PPTX = Path('Slide Template.pptx')

STATUS_COLORS = {
    'In Progress': '#f5a623',
    'Completed': '#2e8b57',
    'Rejected': '#c41230',
}


def strip_accents(text: str) -> str:
    text = str(text).replace('Đ', 'D').replace('đ', 'd')
    text = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in text if not unicodedata.combining(c))


def normalize_text(value) -> str:
    if pd.isna(value):
        return ''
    text = str(value).strip().replace('\u00a0', ' ')
    text = re.sub(r'\s*\([^)]*\)', '', text)
    text = strip_accents(text).lower()
    text = re.sub(r'[^a-z0-9/\-\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def pretty_name(value, blank_value: str = '') -> str:
    norm = normalize_text(value)
    if not norm:
        return blank_value
    norm = re.sub(r'\s+\d+\s*$', '', norm).strip()
    return ' '.join(part.capitalize() for part in norm.split())


def parse_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, dayfirst=True, errors='coerce')


def parse_money(value) -> float:
    if pd.isna(value):
        return 0.0
    s = str(value).strip().replace(',', '')
    if not s:
        return 0.0
    s = re.sub(r'[^0-9.\-]', '', s)
    if s in {'', '.', '-', '-.'}:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def approval_status(value) -> str:
    s = normalize_text(value)
    if 'reject' in s:
        return 'Rejected'
    if 'complete' in s:
        return 'Completed'
    # Dashboard wording request: Submitted should be shown as In Progress.
    return 'In Progress'


def build_week_map_2026():
    start = date(2025, 12, 29)
    weeks = []
    for i in range(53):
        ws = start + timedelta(weeks=i)
        we = ws + timedelta(days=6)
        weeks.append({'label': f'W{i + 1}', 'start': ws, 'end': we})
    return weeks


WEEKS_2026 = build_week_map_2026()


def date_to_week_label(dt: pd.Timestamp) -> str:
    if pd.isna(dt):
        return 'No Week'
    d = dt.date()
    for wk in WEEKS_2026:
        if wk['start'] <= d <= wk['end']:
            return wk['label']
    return 'Out of 2026 Range'


def date_to_month_label(dt: pd.Timestamp) -> str:
    if pd.isna(dt):
        return 'No Month'
    ts = pd.Timestamp(dt)
    return ts.strftime('%Y-%m')


def format_display_date(dt) -> str:
    if pd.isna(dt):
        return ''
    ts = pd.Timestamp(dt)
    return f'{ts.day}/{ts.month}/{ts.year}'


def _ensure_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    required_defaults = {
        'Reference #': '',
        'Improvement Title': '',
        'Before Improvement (By Explanation)': '',
        'After Improvement (By Explanation)': '',
        'Submitter': '',
        'GK Owner Name': '',
        'Submitter Department': '',
        'Approval Status': '',
        'GK Type': '',
        'GK Other': '',
        'Cost Reduction (USD)': 0,
        'Submitted Date': pd.NaT,
        'Completed Date': pd.NaT,
        'Approval Date': pd.NaT,
        'Approved Date': pd.NaT,
        'Supervisor': '',
        'Supervisor Matched Name': '',
        'Supervisor Matched By': '',
        'Supervisor Matched IDL Column': '',
    }
    out = df.copy()
    for col, default in required_defaults.items():
        if col not in out.columns:
            out[col] = default
    return out


def load_data(path: Path, sheet_name=0) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = _ensure_required_columns(df)
    df = df.copy()

    df['Approval Status Clean'] = df['Approval Status'].apply(approval_status)
    df['Submitted Date Parsed'] = parse_date(df['Submitted Date'])
    df['Completed Date Parsed'] = parse_date(df['Completed Date'])
    df['Approval Date Parsed'] = parse_date(df['Approved Date']).combine_first(parse_date(df['Approval Date'])).combine_first(df['Completed Date Parsed'])
    df['Cost Reduction Numeric'] = df['Cost Reduction (USD)'].apply(parse_money)

    # Raw data for detail table: keep exactly as source file.
    df['Submitter Raw'] = df['Submitter'].fillna('').astype(str)
    df['GK Owner Raw'] = df['GK Owner Name'].fillna('').astype(str)
    df['GK Type Display'] = df['GK Type'].fillna('').astype(str).str.strip()
    df['GK Type Filter'] = df['GK Type Display'].where(df['GK Type Display'].str.strip().ne(''), '(Blank)')
    df['GK Other Display'] = df['GK Other'].fillna('').astype(str).str.strip()

    # Dashboard owner for charts/filter:
    # Use mapped Full name from IDL first.
    # This value is generated by Zone3_GK_folder_updater.py.
    mapped_owner = df['Supervisor Matched Name'].fillna('').astype(str)

    df['Owner Display'] = mapped_owner.apply(lambda x: pretty_name(x, blank_value=''))

    # Fallback only for custom dashboard/upload files that do not have IDL mapping.
    fallback_owner = df['GK Owner Name'].apply(lambda x: pretty_name(x, blank_value=''))
    fallback_submitter = df['Submitter'].apply(lambda x: pretty_name(x, blank_value=''))

    df.loc[df['Owner Display'].astype(str).str.strip().eq(''), 'Owner Display'] = fallback_owner
    df.loc[df['Owner Display'].astype(str).str.strip().eq(''), 'Owner Display'] = fallback_submitter

    # Keep Submitter Display as raw text for detail table.
    df['Submitter Display'] = df['Submitter Raw']

    # Supervisor filter/display.
    df['Supervisor Display'] = df['Supervisor'].apply(lambda x: pretty_name(x, blank_value=''))

    # Completed cases are recorded on approval/completed date; all others use submitted date.
    df['Event Date'] = pd.to_datetime(df['Submitted Date Parsed'], errors='coerce')
    completed_with_date = df['Approval Status Clean'].eq('Completed') & df['Approval Date Parsed'].notna()
    df.loc[completed_with_date, 'Event Date'] = df.loc[completed_with_date, 'Approval Date Parsed']
    df['Event Week'] = df['Event Date'].apply(date_to_week_label)
    df['Event Month'] = df['Event Date'].apply(date_to_month_label)
    df['Submitted Week'] = df['Submitted Date Parsed'].apply(date_to_week_label)
    df['Completed Week'] = df['Completed Date Parsed'].apply(date_to_week_label)

    df['Completed Savings USD'] = np.where(
        df['Approval Status Clean'].eq('Completed'),
        df['Cost Reduction Numeric'],
        0.0,
    )

    df['Submitted Date Text'] = df['Submitted Date Parsed'].dt.strftime('%Y-%m-%d').fillna('')
    df['Completed Date Text'] = df['Completed Date Parsed'].dt.strftime('%Y-%m-%d').fillna('')
    df['Event Date Text'] = df['Event Date'].dt.strftime('%Y-%m-%d').fillna('')

    return df


def normalize_reference(value) -> str:
    if pd.isna(value):
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    return s.upper()


def load_many_data(paths: Iterable[Path], sheet_name=0) -> pd.DataFrame:
    frames = []
    for path in paths:
        frames.append(load_data(path, sheet_name=sheet_name))
    if not frames:
        raise ValueError('No input files were provided.')
    df = pd.concat(frames, ignore_index=True, sort=False)
    if 'Reference #' in df.columns:
        df['__ref_key'] = df['Reference #'].apply(normalize_reference)
        with_ref = df[df['__ref_key'] != ''].drop_duplicates(subset='__ref_key', keep='last')
        no_ref = df[df['__ref_key'] == '']
        df = pd.concat([with_ref, no_ref], ignore_index=True, sort=False).drop(columns='__ref_key', errors='ignore')
    return df


def serializable_records(df: pd.DataFrame):
    keep_cols = [
        'Reference #', 'Improvement Title', 'Before Improvement (By Explanation)',
        'After Improvement (By Explanation)', 'Submitter Raw', 'GK Owner Raw',
        'Submitter Display', 'Owner Display', 'Supervisor Display',
        'Submitter Department', 'Approval Status Clean', 'GK Type Display',
        'GK Type Filter', 'GK Other Display', 'Cost Reduction Numeric',
        'Completed Savings USD', 'Submitted Date Text',
        'Completed Date Text', 'Event Date Text', 'Event Week', 'Submitted Week',
        'Completed Week', 'Event Month',
    ]
    work = df.copy()
    for col in keep_cols:
        if col not in work.columns:
            work[col] = ''

    records = []
    for _, row in work[keep_cols].iterrows():
        records.append({
            'reference': '' if pd.isna(row['Reference #']) else str(row['Reference #']),
            'title': '' if pd.isna(row['Improvement Title']) else str(row['Improvement Title']),
            'before': '' if pd.isna(row['Before Improvement (By Explanation)']) else str(row['Before Improvement (By Explanation)']),
            'after': '' if pd.isna(row['After Improvement (By Explanation)']) else str(row['After Improvement (By Explanation)']),
            # owner = mapped Full name for charts/filter
            'owner': row['Owner Display'],

            # raw fields = show in detail table
            'ownerRaw': '' if pd.isna(row['GK Owner Raw']) else str(row['GK Owner Raw']),
            'submitter': '' if pd.isna(row['Submitter Raw']) else str(row['Submitter Raw']),
            'supervisor': row['Supervisor Display'] if str(row['Supervisor Display']).strip() else '',
            'supervisorFilter': row['Supervisor Display'] if str(row['Supervisor Display']).strip() else '(Blank)',
            # Home menu uses Supervisor as the comparison dimension. Keep zone as backward-compatible alias.
            'zone': row['Supervisor Display'] if str(row['Supervisor Display']).strip() else '(Blank)',
            'department': '' if pd.isna(row['Submitter Department']) else str(row['Submitter Department']),
            'status': row['Approval Status Clean'],
            'gkType': '' if pd.isna(row['GK Type Display']) else str(row['GK Type Display']),
            'gkTypeFilter': row['GK Type Filter'] if str(row['GK Type Filter']).strip() else '(Blank)',
            'gkOther': '' if pd.isna(row['GK Other Display']) else str(row['GK Other Display']),
            'costReduction': round(float(row['Cost Reduction Numeric'] or 0), 2),
            'completedSavings': round(float(row['Completed Savings USD'] or 0), 2),
            'submittedDate': row['Submitted Date Text'],
            'completedDate': row['Completed Date Text'],
            'eventDate': row['Event Date Text'],
            'eventWeek': row['Event Week'],
            'eventMonth': row['Event Month'],
            'submittedWeek': row['Submitted Week'],
            'completedWeek': row['Completed Week'],
        })
    return records


def _checkbox_html(filter_name, values):
    items = []
    for val in values:
        safe = html.escape(str(val), quote=True)
        items.append(f'<label class="check-item"><input type="checkbox" data-filter="{filter_name}" value="{safe}" checked> <span>{safe}</span></label>')
    return '\n'.join(items)


def render_dashboard(data_records, source_name: str, logo_data_uri: str = '', latest_update_text: str = '', show_supervisor_filter: bool = True) -> str:
    plotly_js = get_plotlyjs()
    statuses = ['Completed', 'In Progress', 'Rejected']
    gk_types = sorted({r.get('gkTypeFilter', '(Blank)') for r in data_records}) or ['(Blank)']
    owners = sorted({r['owner'] for r in data_records if r.get('owner')}) or ['Unknown']
    supervisors = sorted({r.get('supervisor') for r in data_records if str(r.get('supervisor') or '').strip()})
    all_weeks = [wk['label'] for wk in WEEKS_2026]
    week_set = {r.get('eventWeek') for r in data_records if r.get('eventWeek') in all_weeks}
    weeks = [wk for wk in all_weeks if wk in week_set] or all_weeks
    month_set = {r.get('eventMonth') for r in data_records if r.get('eventMonth')}
    months = sorted(month_set) or ['No Month']

    template = r'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>GK Dashboard</title>
  <style>
    :root {
      --bg:#f3f5f7; --panel:#ffffff; --panel-2:#fbfbfc; --ink:#151515; --muted:#69707d;
      --line:#e3e6ea; --shadow:0 12px 28px rgba(16,24,40,.08); --red:#c41230; --red-2:#92142a;
      --black:#101114; --green:#2e8b57; --amber:#f5a623; --radius:18px; --dashboard-scale:.72; --dashboard-width:138.889vw;
    }
    *{box-sizing:border-box} html,body{margin:0;padding:0;width:100%;min-height:100%;font-family:Inter,Segoe UI,Arial,sans-serif;color:var(--ink);background:var(--bg);overflow-x:hidden}
    .dashboard-viewport{width:100vw;min-height:100vh;overflow-x:hidden;background:var(--bg)}
    .app{width:var(--dashboard-width);min-height:100vh;transform:scale(var(--dashboard-scale));transform-origin:top left}
    .hero{position:relative;overflow:hidden;background:radial-gradient(circle at 100% 0%,rgba(196,18,48,.30),transparent 38%),linear-gradient(120deg,#0d0f13 0%,#151820 58%,#4e0d1d 100%);color:white;padding:22px 30px 20px;border-bottom:4px solid var(--red)}
    .hero:before{content:'';position:absolute;left:30px;bottom:0;width:240px;height:4px;background:linear-gradient(90deg,#ff3959,rgba(255,57,89,0));border-radius:999px;opacity:.9}
    .hero:after{content:'';position:absolute;left:30px;top:18px;width:58px;height:2px;background:rgba(255,255,255,.28);border-radius:999px}
    .hero-top{position:relative;z-index:1;display:flex;align-items:flex-start;justify-content:space-between;gap:24px}
    .hero-kicker{font-size:11px;text-transform:uppercase;letter-spacing:.16em;color:rgba(255,255,255,.44);font-weight:900;margin-bottom:8px}
    .hero h1{margin:0;font-size:34px;line-height:1.05;letter-spacing:-.04em;text-shadow:0 10px 24px rgba(0,0,0,.22)}
    .dashboard-owner-line{margin-top:10px;font-size:13px;font-weight:600;color:rgba(255,255,255,.45);letter-spacing:.01em}.dashboard-owner-line span{color:rgba(255,255,255,.62);font-weight:700}
    .hero-update{margin-top:6px;font-size:12px;font-weight:600;color:rgba(255,255,255,.38);letter-spacing:.01em}
    .hero-update span{color:rgba(255,255,255,.55);font-weight:700}
    .hero-logo-card{position:relative;z-index:1;flex:0 0 auto;width:340px;max-width:340px;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.14);border-radius:18px;padding:14px 20px;box-shadow:0 16px 36px rgba(0,0,0,.18);backdrop-filter:blur(6px)}
    .hero-logo{display:block;width:100%;height:auto;max-height:98px;object-fit:contain}
    .back-home-link{display:inline-flex;align-items:center;gap:6px;color:#fff;text-decoration:none;background:rgba(255,255,255,.11);border:1px solid rgba(255,255,255,.22);border-radius:12px;padding:8px 12px;font-weight:900;font-size:13px;margin-bottom:12px;box-shadow:0 10px 22px rgba(0,0,0,.16)}.back-home-link:hover{background:rgba(255,255,255,.18)}
    .hero-watermark{position:absolute;right:26px;bottom:10px;opacity:.055;pointer-events:none}.hero-watermark-img{width:220px;height:auto;display:block}
    .page{padding:18px}.filter-panel{position:sticky;top:0;z-index:15;margin-top:-28px;background:rgba(255,255,255,.94);backdrop-filter:blur(14px);border:1px solid rgba(255,255,255,.9);box-shadow:var(--shadow);border-radius:20px;padding:14px;display:grid;gap:12px}
    .filter-head{display:flex;justify-content:space-between;align-items:center;gap:12px}.filter-title{font-size:18px;font-weight:800}.filter-actions{display:flex;gap:10px;align-items:center}
    .btn{border:0;border-radius:14px;padding:11px 15px;font-weight:800;cursor:pointer;transition:.18s ease;font-size:14px}.btn-primary{background:var(--red);color:#fff}.btn-primary:hover{background:var(--red-2)}.btn-ghost{background:#eef1f4;color:#222}
    .export-menu{position:relative}.export-menu summary{list-style:none}.export-menu summary::-webkit-details-marker{display:none}.export-menu-panel{position:absolute;right:0;top:calc(100% + 8px);min-width:170px;z-index:45;background:#fff;border:1px solid var(--line);border-radius:16px;box-shadow:var(--shadow);padding:8px;display:grid;gap:6px}.export-option{width:100%;text-align:left;background:#fff}.export-option:hover{background:#f4f6f8}
    .filter-grid{display:grid;grid-template-columns:1fr 1fr .78fr .72fr .72fr .72fr 1.45fr;gap:10px;align-items:start}.filter-grid.no-supervisor{grid-template-columns:1.2fr .85fr .8fr .8fr .8fr 1.45fr}.filter-block{background:var(--panel-2);border:1px solid var(--line);border-radius:16px;padding:9px 11px;min-height:72px}.filter-block label.title{display:block;font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:7px;font-weight:800}
    .filter-block input[type="date"],.filter-block input[type="text"]{width:100%;border:1px solid var(--line);background:#fff;border-radius:12px;padding:9px 10px;font-size:13px;outline:none}.date-range{display:grid;grid-template-columns:1fr 1fr;gap:8px}
    details.multi{position:relative}details.multi summary{list-style:none;cursor:pointer;border:1px solid var(--line);background:#fff;border-radius:12px;padding:10px 12px;display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:14px;color:#1a1a1a}details.multi summary::-webkit-details-marker{display:none}.summary-count{color:var(--muted);font-weight:600}
    .option-panel{position:absolute;left:0;right:0;top:calc(100% + 8px);z-index:80;background:#fff;border:1px solid var(--line);border-radius:16px;box-shadow:var(--shadow);padding:0;max-height:330px;overflow:auto}
    .option-tools{display:grid;grid-template-columns:1fr auto auto auto;gap:8px;align-items:center;background:#fff;padding:10px 10px 9px 10px;margin-bottom:6px;border-bottom:1px solid #edf0f2;position:sticky;top:0;z-index:3}
    .option-tools.no-search{grid-template-columns:auto auto auto;justify-content:start}
    .option-back{margin-left:0}
    .filter-search{width:100%;border:1px solid var(--line);background:#fff;border-radius:10px;padding:7px 9px;font-size:12px;outline:none}
    .tiny-btn{border:1px solid var(--line);background:#fff;border-radius:10px;padding:6px 8px;font-size:12px;cursor:pointer;white-space:nowrap}
    .check-list{display:grid;gap:6px;padding:0 10px 10px 10px}
    .check-item{display:flex;gap:9px;align-items:center;padding:6px 4px;border-radius:10px;font-size:13px;line-height:1.25}
    .check-item:hover{background:#f4f6f8}
    .active-filters{display:flex;flex-wrap:wrap;gap:8px}.chip{padding:8px 10px;border-radius:999px;background:#fff1f4;color:var(--red);border:1px solid #ffd0da;font-size:12px;font-weight:700}
    .kpi-grid{margin-top:14px;display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:14px}.kpi-card{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);padding:15px;box-shadow:var(--shadow);position:relative;overflow:hidden}.kpi-card:before{content:'';position:absolute;inset:0 auto 0 0;width:5px;background:var(--red)}.kpi-card.primary{background:linear-gradient(135deg,#14171d 0%,#1b2027 100%);color:#fff;border-color:#1e2024}.kpi-card.primary:before{background:linear-gradient(180deg,#ff3f60,#c41230)}.kpi-label{font-size:13px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);font-weight:800}.kpi-card.primary .kpi-label{color:rgba(255,255,255,.62)}.kpi-value{margin-top:8px;font-size:32px;font-weight:900;letter-spacing:-.03em}
    .main-grid{margin-top:14px;display:grid;grid-template-columns:2fr 1fr;gap:16px}.wide-grid{display:grid;grid-template-columns:1.7fr .9fr;gap:14px;margin-top:14px;align-items:start}.bottom-grid{display:grid;grid-template-columns:1.2fr .8fr;gap:14px;margin-top:14px;align-items:stretch}.card{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:18px}.card-header{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:12px}.card-title{font-size:22px;font-weight:900;letter-spacing:-.02em}.card-badge{padding:9px 12px;border-radius:999px;background:#f4f6f8;color:#26303d;font-size:14px;font-weight:800;white-space:nowrap}.plot{width:100%;height:360px}.plot.small{height:340px}.plot.owner-main{height:390px}.plot.savings-main{height:390px}
    .credit-list{display:grid;gap:12px}.credit-item{background:#fafbfc;border:1px solid var(--line);border-radius:14px;padding:12px 13px}.credit-label{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);font-weight:800;margin-bottom:5px}.credit-name{font-size:15px;font-weight:800;color:#111317;line-height:1.45}.muted{color:var(--muted)}
    .table-card{margin-top:16px}.table-tools{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:12px;flex-wrap:wrap}.top-scrollbar{overflow-x:auto;overflow-y:hidden;height:14px;background:#fff;border:1px solid var(--line);border-bottom:none;border-radius:18px 18px 0 0}.top-scrollbar.hidden{display:none}.top-scrollbar-inner{height:1px}.table-scroll{overflow:auto;border:1px solid var(--line);border-radius:0 0 18px 18px;background:#fff}.table-scroll.standalone{border-radius:18px}table{border-collapse:separate;border-spacing:0;min-width:1680px;width:max-content;background:#fff;table-layout:fixed}thead th{position:sticky;top:0;z-index:1;background:#111317;color:#fff;font-size:11px;text-transform:uppercase;letter-spacing:.045em;padding:10px 7px;text-align:left;white-space:nowrap}tbody td{padding:10px 7px;border-bottom:1px solid #edf0f2;vertical-align:top;font-size:12px;line-height:1.35;overflow:hidden}tbody tr:nth-child(even){background:#fafbfc}.status-pill{display:inline-flex;align-items:center;gap:4px;padding:5px 8px;border-radius:999px;font-size:11px;font-weight:800}.status-InProgress{background:#fff3da;color:#9d5b00}.status-Completed{background:#ddf5e8;color:#13623b}.status-Rejected{background:#ffe3e9;color:#9b1331}.title-cell{font-weight:800}.desc-cell{color:#4e5561}.clamp{display:-webkit-box;-webkit-box-orient:vertical;overflow:hidden}.clamp-title{-webkit-line-clamp:3}.clamp-desc{-webkit-line-clamp:4}.tracking-col{min-width:160px}.table-note{font-size:11px;color:#6b7280;font-weight:700;margin:-3px 0 8px}.pending-table{min-width:1120px}#followTable{min-width:1450px}#detailTable{min-width:2200px}


    /* v8.1 compact tracking table: fixed widths + multiline clamp */
    #detailTable{min-width:1980px;table-layout:fixed}
    #followTable{min-width:1420px;table-layout:fixed}
    #pendingTable{min-width:1260px;table-layout:fixed}
    .col-ref{width:110px;min-width:110px;max-width:110px}
    .col-title{width:220px;min-width:220px;max-width:220px}
    .col-supervisor{width:135px;min-width:135px;max-width:135px}
    .col-owner{width:150px;min-width:150px;max-width:150px}
    .col-submitter{width:160px;min-width:160px;max-width:160px}
    .col-status{width:110px;min-width:110px;max-width:110px}
    .col-follow-status{width:110px;min-width:110px;max-width:110px}
    .col-deadline{width:120px;min-width:120px;max-width:120px}
    .col-comment{width:210px;min-width:210px;max-width:210px}
    .col-updated{width:140px;min-width:140px;max-width:140px}
    .col-type{width:105px;min-width:105px;max-width:105px}
    .col-other{width:140px;min-width:140px;max-width:140px}
    .col-month,.col-week,.col-date{width:100px;min-width:100px;max-width:100px}
    .col-savings{width:90px;min-width:90px;max-width:90px}
    .col-before,.col-after{width:260px;min-width:260px;max-width:260px}
    .col-pending{width:105px;min-width:105px;max-width:105px}
    .cell-clip{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:2;overflow:hidden;white-space:normal;word-break:break-word;overflow-wrap:anywhere;max-height:42px}
    .cell-clip.title-text{font-weight:900;line-height:1.25;max-height:44px}
    .cell-clip.desc-text{color:#4e5561;line-height:1.32;max-height:42px}
    .cell-clip.comment-text{color:#4e5561;line-height:1.32;max-height:42px}
    #detailTable tbody td,#followTable tbody td,#pendingTable tbody td{height:54px}
    #detailTable tbody tr,#followTable tbody tr,#pendingTable tbody tr{height:54px}
    .table-scroll{max-height:540px}
    .toast-modal-backdrop{position:fixed;inset:0;background:rgba(16,17,20,.32);display:none;align-items:center;justify-content:center;z-index:1000000}
    .toast-modal-card{width:min(420px,92vw);background:#fff;border:1px solid var(--line);border-radius:22px;box-shadow:0 22px 70px rgba(0,0,0,.25);padding:24px;text-align:center}
    .toast-icon{width:52px;height:52px;border-radius:999px;background:#ddf5e8;color:#13623b;display:inline-flex;align-items:center;justify-content:center;font-size:28px;font-weight:900;margin-bottom:12px}
    .toast-title{font-size:21px;font-weight:950;margin-bottom:6px}
    .toast-msg{font-size:13px;color:#69707d;font-weight:700;margin-bottom:16px}
    .toast-close{border:0;background:#111317;color:#fff;border-radius:12px;padding:10px 18px;font-weight:900;cursor:pointer}
    #trackingFollowStatus:disabled{background:#f3f5f7;color:#69707d;cursor:not-allowed}

    /* v8.2 fit Follow-up and Pending tabs into the card */
    #followTable{width:100%!important;min-width:0!important;table-layout:fixed}
    #pendingTable{width:100%!important;min-width:0!important;table-layout:fixed}
    #followTable .col-ref,#followTable .col-date{width:105px;min-width:105px;max-width:105px}
    #followTable .col-title{width:230px;min-width:230px;max-width:230px}
    #followTable .col-supervisor{width:120px;min-width:120px;max-width:120px}
    #followTable .col-owner{width:145px;min-width:145px;max-width:145px}
    #followTable .col-status{width:105px;min-width:105px;max-width:105px}
    #followTable .col-type{width:90px;min-width:90px;max-width:90px}
    #followTable .col-deadline{width:110px;min-width:110px;max-width:110px}
    #followTable .col-follow-status{width:95px;min-width:95px;max-width:95px}
    #followTable .col-comment{width:260px;min-width:260px;max-width:260px}
    #followTable .col-updated{width:130px;min-width:130px;max-width:130px}
    #pendingTable .col-ref{width:105px;min-width:105px;max-width:105px}
    #pendingTable .col-title{width:210px;min-width:210px;max-width:210px}
    #pendingTable .col-supervisor{width:118px;min-width:118px;max-width:118px}
    #pendingTable .col-owner{width:135px;min-width:135px;max-width:135px}
    #pendingTable .col-status{width:95px;min-width:95px;max-width:95px}
    #pendingTable .col-pending{width:90px;min-width:90px;max-width:90px}
    #pendingTable .col-deadline{width:105px;min-width:105px;max-width:105px}
    #pendingTable .col-comment{width:180px;min-width:180px;max-width:180px}
    .pending-layout{grid-template-columns:minmax(0,.95fr) minmax(0,1.05fr)!important;gap:12px!important;align-items:stretch}
    .pending-layout>div,.pending-chart-card,.pending-list{min-width:0}
    .pending-chart-card{overflow:hidden}
    .pending-chart{height:360px!important}
    #pendingTableWrap{max-height:420px}
    #followTableWrap{max-height:420px}

    .detail-tabs{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:10px}.detail-tab-btn{border:1px solid var(--line);background:#f6f8fa;color:#4b5563;border-radius:999px;padding:9px 12px;font-weight:900;font-size:13px;cursor:pointer;transition:.18s ease}.detail-tab-btn:hover{background:#eef1f4;transform:translateY(-1px)}.detail-tab-btn.active{background:#111317;color:#fff;border-color:#111317;box-shadow:0 10px 24px rgba(16,24,40,.16)}.detail-tab-panel{display:none}.detail-tab-panel.active{display:block}.table-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.clickable-row{cursor:pointer;transition:background .14s ease,box-shadow .14s ease}.clickable-row:hover{background:#fff8f1!important;box-shadow:inset 4px 0 0 #c41230}.clickable-row.follow-row{background:#fffdf2}.follow-star{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:999px;background:#fff2b8;color:#9a6700;font-weight:900;margin-right:5px}.tracking-mini{display:block;margin-top:5px;font-size:10.5px;font-weight:800;color:#69707d}.tracking-mini.overdue{color:#c41230}.pending-layout{display:grid;grid-template-columns:1.05fr 1.35fr;gap:14px;align-items:stretch}.pending-list{min-height:390px}.pending-chart-card{border:1px solid var(--line);border-radius:18px;padding:12px;background:#fbfcfd}.pending-chart{height:420px}.empty-state{display:grid;place-items:center;min-height:180px;color:#69707d;font-weight:800;text-align:center;border:1px dashed #d9dee5;border-radius:16px;background:#fafbfc}.case-modal-backdrop{background:rgba(12,14,18,.62);backdrop-filter:blur(7px)}.case-modal{width:min(1120px,94vw);padding:0;overflow:hidden}.case-modal .modal-head{background:linear-gradient(120deg,#101114 0%,#181b22 58%,#5d0e1f 100%);color:#fff;padding:18px 20px;margin:0}.case-modal .modal-head .card-title{font-size:22px;color:#fff}.case-modal-body{padding:18px 20px 20px;display:grid;grid-template-columns:1.18fr .82fr;gap:18px;align-items:start}.case-detail-grid{display:grid;grid-template-columns:160px 1fr;gap:10px 14px;border:1px solid var(--line);border-radius:18px;padding:16px;background:#fbfcfd}.case-label{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#69707d;font-weight:900}.case-value{font-size:13px;line-height:1.42;color:#161a22;font-weight:650}.case-value.long{font-weight:500;white-space:pre-wrap}.tracking-form{border:1px solid var(--line);border-radius:18px;padding:16px;background:#fff;box-shadow:0 12px 28px rgba(16,24,40,.06)}.tracking-form h3{margin:0 0 12px;font-size:18px}.form-row{display:grid;gap:7px;margin-bottom:12px}.form-row label{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:#69707d;font-weight:900}.form-row input,.form-row textarea,.form-row select{width:100%;border:1px solid #dfe4ea;border-radius:12px;padding:10px 11px;font:inherit;font-size:13px;outline:none;background:#fff}.form-row textarea{min-height:110px;resize:vertical}.inline-check{display:flex;gap:10px;align-items:center;background:#fff8f1;border:1px solid #ffdfb8;border-radius:14px;padding:10px 12px;margin-bottom:12px;font-weight:900;color:#5b3b00}.inline-check input{width:18px;height:18px}.tracking-status-msg{font-size:12px;font-weight:800;color:#69707d;min-height:18px}.tracking-status-msg.ok{color:#0b7a3b}.tracking-status-msg.err{color:#c41230}.modal-close-btn{border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.10);color:#fff;border-radius:12px;padding:8px 12px;font-weight:900;cursor:pointer}.modal-close-btn:hover{background:rgba(255,255,255,.18)}.tracking-disabled-note{font-size:12px;color:#8a4b00;background:#fff8e1;border:1px solid #ffe1a6;border-radius:12px;padding:9px 10px;margin-bottom:12px}.deadline-pill{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:4px 8px;background:#eef1f4;color:#26303d;font-weight:900;font-size:11px}.deadline-pill.overdue{background:#ffe3e9;color:#9b1331}.deadline-pill.soon{background:#fff3da;color:#9d5b00}@media(max-width:900px){.case-modal-body,.pending-layout{grid-template-columns:1fr}.case-detail-grid{grid-template-columns:1fr}.detail-tabs{overflow:auto;flex-wrap:nowrap}.detail-tab-btn{white-space:nowrap}}

    .modal-backdrop{position:fixed;inset:0;background:rgba(16,17,20,.45);display:none;align-items:center;justify-content:center;padding:24px;z-index:999999}.modal-card{width:min(1240px,96vw);max-height:88vh;overflow:auto;background:#fff;border:1px solid var(--line);border-radius:22px;box-shadow:0 20px 50px rgba(0,0,0,.18);padding:18px}.modal-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:14px}.full-chart-wrap{border:1px solid var(--line);border-radius:16px;padding:10px;margin-bottom:14px}.full-chart-plot{width:100%;height:520px}.summary-table{min-width:0!important;width:100%;table-layout:fixed}.summary-table thead th{position:static;padding:10px 8px;font-size:11px}.summary-table tbody td{padding:10px 8px;font-size:12px}.count-badge{display:inline-flex;min-width:34px;justify-content:center;align-items:center;padding:4px 8px;border-radius:999px;font-size:12px;font-weight:800}.count-inprogress{background:#fff3da;color:#9d5b00}.count-completed{background:#ddf5e8;color:#13623b}.count-rejected{background:#ffe3e9;color:#9b1331}
    .pdf-mode .filter-panel,.pdf-mode .filter-actions,.pdf-mode .table-card,.pdf-mode .modal-backdrop,.pdf-mode .active-filters,.pdf-mode .no-print{display:none!important}
    @media (max-width:900px){.filter-grid,.kpi-grid,.main-grid,.wide-grid,.bottom-grid{grid-template-columns:1fr}.page{padding:14px}.hero{padding:24px 18px 18px}.hero-top{flex-direction:column}.hero-logo{width:170px}}
  </style>
</head>
<body>
  <div class="dashboard-viewport"><div class="app">
    <header class="hero"><div class="hero-top"><div><a id="officialBackHomeButton" class="back-home-link no-print" href="index.html">← Back to menu</a><div class="hero-kicker">Gemba Kaizen Dashboard</div><h1>GK Performance Overview</h1><div class="dashboard-owner-line">Dashboard owner <span>Nguyen Viet Bao</span></div>__LATEST_UPDATE_LINE__</div>__HERO_LOGO__</div>__HERO_WATERMARK__</header>
    <main class="page"><div id="reportForPdf">
      <section class="filter-panel">
        <div class="filter-head"><div class="filter-title">Dashboard Filters</div><div class="filter-actions no-print"><button class="btn btn-ghost" id="resetFilters">Reset filters</button><details class="export-menu" id="exportMenu"><summary class="btn btn-ghost">Export dashboard</summary><div class="export-menu-panel"><button class="tiny-btn export-option" id="exportPdfOption" type="button">Export as PDF</button><button class="tiny-btn export-option" id="exportPngOption" type="button">Export as PNG</button></div></details><button class="btn btn-primary" id="exportCsv">Export filtered CSV</button></div></div>
        <div class="filter-grid __FILTER_GRID_CLASS__">
__SUPERVISOR_FILTER_BLOCK__
          <div class="filter-block"><label class="title">GK Owner</label><details class="multi"><summary><span>Select owner(s)</span><span class="summary-count" id="ownerSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search owner..." class="filter-search" data-target="ownerOptions"><button class="tiny-btn" data-action="all" data-filter="owner">All</button><button class="tiny-btn" data-action="none" data-filter="owner">None</button><button class="tiny-btn option-back" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="ownerOptions">__OWNER_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Status</label><details class="multi"><summary><span>Select status</span><span class="summary-count" id="statusSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="status">All</button><button class="tiny-btn" data-action="none" data-filter="status">None</button><button class="tiny-btn option-back" type="button" data-close-filter="true">Back</button></div><div class="check-list">__STATUS_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">GK Type</label><details class="multi"><summary><span>Select type</span><span class="summary-count" id="gkTypeSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search type..." class="filter-search" data-target="gkTypeOptions"><button class="tiny-btn" data-action="all" data-filter="gkType">All</button><button class="tiny-btn" data-action="none" data-filter="gkType">None</button><button class="tiny-btn option-back" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="gkTypeOptions">__GK_TYPE_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Month</label><details class="multi"><summary><span>Select month</span><span class="summary-count" id="monthSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="month">All</button><button class="tiny-btn" data-action="none" data-filter="month">None</button><button class="tiny-btn option-back" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="monthOptions">__MONTH_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Week</label><details class="multi"><summary><span>Select week(s)</span><span class="summary-count" id="weekSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search week..." class="filter-search" data-target="weekOptions"><button class="tiny-btn" data-action="all" data-filter="week">All</button><button class="tiny-btn" data-action="none" data-filter="week">None</button><button class="tiny-btn option-back" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="weekOptions">__WEEK_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Date Range</label><div class="date-range"><input type="date" id="dateFrom"/><input type="date" id="dateTo"/></div></div>
        </div><div class="active-filters" id="activeFilters"></div>
      </section>
      <section class="kpi-grid"><div class="kpi-card primary"><div class="kpi-label">Total GK</div><div class="kpi-value" id="kpiTotal">0</div></div><div class="kpi-card"><div class="kpi-label">In Progress</div><div class="kpi-value" id="kpiSubmitted">0</div></div><div class="kpi-card"><div class="kpi-label">Completed</div><div class="kpi-value" id="kpiCompleted">0</div></div><div class="kpi-card"><div class="kpi-label">Rejected</div><div class="kpi-value" id="kpiRejected">0</div></div><div class="kpi-card"><div class="kpi-label">Completed Savings</div><div class="kpi-value" id="kpiSavings">$0</div></div></section>
      <section class="main-grid"><div class="card"><div class="card-header"><div class="card-title">Weekly GK Activity</div><div class="card-badge" id="weeklyBadge">0 cases</div></div><div id="weeklyChart" class="plot"></div></div><div class="card"><div class="card-header"><div class="card-title">Status Mix</div></div><div id="statusChart" class="plot small"></div></div></section>
      <section class="wide-grid"><div class="card"><div class="card-header"><div class="card-title">Top GK Owners</div><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap" class="no-print"><div class="card-badge">Top 10</div><button class="btn btn-ghost" id="viewAllOwnersChartBtn" style="padding:8px 12px">Show all chart</button></div></div><div id="ownerChart" class="plot owner-main"></div></div><div class="card"><div class="card-header"><div class="card-title">Completed Savings</div><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap" class="no-print"><div class="card-badge" id="savingsBadge">$0</div><button class="btn btn-ghost" id="viewAllSavingsChartBtn" style="padding:8px 12px">Show all chart</button></div></div><div id="savingsChart" class="plot savings-main"></div></div></section>
    </div>
      <section class="card table-card">
        <div class="table-tools">
          <div>
            <div class="card-title">Detailed GK List</div>
            <div class="detail-tabs no-print" role="tablist" aria-label="Detailed GK tabs">
              <button class="detail-tab-btn active" data-detail-tab="all" type="button">All GK Records</button>
              <button class="detail-tab-btn" data-detail-tab="pending" type="button">Pending Timeline</button>
            </div>
          </div>
          <div class="table-actions">
            <div class="card-badge" id="tableBadge">0 rows</div>
          </div>
        </div>
        <div id="tabAll" class="detail-tab-panel active">
          <div class="table-note"></div><div class="top-scrollbar" id="detailTopScroll"><div class="top-scrollbar-inner" id="detailTopScrollInner"></div></div><div class="table-scroll" id="detailTableWrap"><table id="detailTable"><thead><tr><th class="col-ref">Reference #</th><th class="col-title">Title</th><th class="col-supervisor">Supervisor</th><th class="col-owner">GK Owner</th><th class="col-submitter">Submitter</th><th class="col-status">Status</th><th class="col-type">GK Type</th><th class="col-other">GK Other</th><th class="col-month">Month</th><th class="col-week">Week</th><th class="col-date">Submit Date</th><th class="col-date">Done Date</th><th class="col-savings">Savings</th><th class="col-before">Before</th><th class="col-after">After</th></tr></thead><tbody id="tableBody"></tbody></table></div>
        </div>
        <div id="tabPending" class="detail-tab-panel">
          <div class="pending-layout">
            <div><div class="top-scrollbar" id="pendingTopScroll"><div class="top-scrollbar-inner" id="pendingTopScrollInner"></div></div><div class="pending-list table-scroll" id="pendingTableWrap"><table id="pendingTable" class="pending-table"><thead><tr><th class="col-ref">Reference #</th><th class="col-title">Title</th><th class="col-supervisor">Supervisor</th><th class="col-owner">GK Owner</th><th class="col-status">Status</th><th class="col-pending">Pending Days</th></tr></thead><tbody id="pendingTableBody"></tbody></table></div></div>
            <div class="pending-chart-card"><div class="card-header"><div class="card-title">Pending Duration</div><div class="card-badge" id="pendingBadge">0 cases</div></div><div id="pendingDurationChart" class="pending-chart"></div></div>
          </div>
        </div>
      </section>
      <div class="modal-backdrop case-modal-backdrop" id="gkDetailModal" style="display:none"><div class="modal-card case-modal"><div class="modal-head"><div><div class="hero-kicker" style="margin-bottom:6px">GK Case Detail</div><div class="card-title" id="caseModalTitle">GK Detail</div></div><button class="modal-close-btn" id="closeGkDetailModal" type="button">Close</button></div><div class="case-modal-body" style="grid-template-columns:1fr"><div><div id="caseDetailFields" class="case-detail-grid"></div></div></div></div></div>
      <div class="modal-backdrop" id="ownerChartModal" style="display:none"><div class="modal-card"><div class="modal-head"><div class="card-title">All GK Owners - Full Chart</div><button class="btn btn-ghost" id="closeOwnerChartModal">Close</button></div><div class="full-chart-wrap"><div id="ownerFullChart" class="full-chart-plot"></div></div><div class="table-scroll"><table class="summary-table"><thead><tr><th>GK Owner</th><th>Supervisor</th><th>Total GK</th><th>In Progress</th><th>Completed</th><th>Rejected</th><th>Completed Savings</th></tr></thead><tbody id="ownerChartModalTableBody"></tbody></table></div></div></div>
      <div class="modal-backdrop" id="savingsChartModal" style="display:none"><div class="modal-card"><div class="modal-head"><div class="card-title">Completed Savings - Full Chart</div><button class="btn btn-ghost" id="closeSavingsChartModal">Close</button></div><div class="full-chart-wrap"><div id="savingsFullChart" class="full-chart-plot"></div></div><div class="table-scroll"><table class="summary-table"><thead><tr><th>GK Owner</th><th>Supervisor</th><th>Total GK</th><th>Completed</th><th>Completed Savings</th><th>In Progress</th><th>Rejected</th></tr></thead><tbody id="savingsChartModalTableBody"></tbody></table></div></div></div>
    </main></div></div>
  <script>__PLOTLY_JS__</script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script>
    const rawData = __RAW_DATA__;
    const statusColors = __STATUS_COLORS__;
    const weekOrder = __WEEK_ORDER__;
    const allOwners = __OWNERS_JSON__;
    const allSupervisors = __SUPERVISORS_JSON__;
    const showSupervisorFilter = __SHOW_SUPERVISOR_FILTER__;
    const allMonths = __MONTHS_JSON__;
    const allGkTypes = __GK_TYPES_JSON__;
    const statuses = ['Completed','In Progress','Rejected'];
    const state = {owners:new Set(allOwners),supervisors:new Set(allSupervisors),statuses:new Set(statuses),gkTypes:new Set(allGkTypes),months:new Set(allMonths),weeks:new Set(weekOrder),dateFrom:'',dateTo:''};
    const fmtInt = new Intl.NumberFormat('en-US',{maximumFractionDigits:0});
    const fmtMoney = new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0});
    const supervisorPalette=['#c41230','#2f80ed','#229954','#f2994a','#9b51e0','#00a39a','#eb5757','#2f4858','#f2c94c','#56ccf2'];
    const supervisorColors={};allSupervisors.forEach((s,i)=>{supervisorColors[s]=supervisorPalette[i%supervisorPalette.length]});
    function shortSupervisorName(name){const parts=String(name||'').trim().split(/\s+/).filter(Boolean);return parts.length?parts[parts.length-1]:String(name||'')}
    function escapeHtml(str){return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;')}
    function statusClass(status){return 'status-' + String(status).replace(/\s+/g,'')}
    function updateSummaries(){
      if(showSupervisorFilter){document.getElementById('supervisorSummary').textContent = state.supervisors.size===allSupervisors.length?'All':`${state.supervisors.size} selected`;}
      document.getElementById('ownerSummary').textContent = state.owners.size===allOwners.length?'All':`${state.owners.size} selected`;
      document.getElementById('statusSummary').textContent = state.statuses.size===statuses.length?'All':`${state.statuses.size} selected`;
      document.getElementById('gkTypeSummary').textContent = state.gkTypes.size===allGkTypes.length?'All':`${state.gkTypes.size} selected`;
      document.getElementById('monthSummary').textContent = state.months.size===allMonths.length?'All':`${state.months.size} selected`;
      document.getElementById('weekSummary').textContent = state.weeks.size===weekOrder.length?'All':`${state.weeks.size} selected`;
    }
    function getFilteredData(){return rawData.filter(d=>{
      if(showSupervisorFilter && state.supervisors.size!==allSupervisors.length && !state.supervisors.has(d.supervisor)) return false; if(!state.owners.has(d.owner)) return false; if(!state.statuses.has(d.status)) return false; if(!state.gkTypes.has(d.gkTypeFilter||'(Blank)')) return false; if(!state.months.has(d.eventMonth)) return false; if(!state.weeks.has(d.eventWeek)) return false; if(state.dateFrom&&(!d.eventDate||d.eventDate<state.dateFrom)) return false; if(state.dateTo&&(!d.eventDate||d.eventDate>state.dateTo)) return false; return true;});}
    function setMetric(id,value){document.getElementById(id).textContent=value}
    function updateKpis(data){const total=data.length, inprog=data.filter(d=>d.status==='In Progress').length, completed=data.filter(d=>d.status==='Completed').length, rejected=data.filter(d=>d.status==='Rejected').length, savings=data.reduce((s,d)=>s+(Number(d.completedSavings)||0),0); setMetric('kpiTotal',fmtInt.format(total));setMetric('kpiSubmitted',fmtInt.format(inprog));setMetric('kpiCompleted',fmtInt.format(completed));setMetric('kpiRejected',fmtInt.format(rejected));setMetric('kpiSavings',fmtMoney.format(savings));document.getElementById('weeklyBadge').textContent=`${fmtInt.format(total)} cases`;document.getElementById('tableBadge').textContent=`${fmtInt.format(total)} rows`;document.getElementById('savingsBadge').textContent=fmtMoney.format(savings);}
    function baseLayout(){return {paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',margin:{l:50,r:20,t:10,b:50},font:{family:'Inter, Segoe UI, Arial, sans-serif',color:'#1a1a1a',size:14},xaxis:{gridcolor:'#eef1f4',zeroline:false,title:''},yaxis:{gridcolor:'#eef1f4',zeroline:false,title:''},legend:{orientation:'h',y:-.18}}}
    function renderWeeklyChartTo(data,targetId){
      const by={};
      weekOrder.forEach(w=>by[w]={'In Progress':0,Completed:0,Rejected:0});
      data.forEach(d=>{
        if(!by[d.eventWeek])by[d.eventWeek]={'In Progress':0,Completed:0,Rejected:0};
        by[d.eventWeek][d.status]+=1;
      });
      const x=weekOrder.filter(w=>statuses.some(s=>(by[w]?.[s]||0)>0));
      const totals=x.map(w=>statuses.reduce((sum,s)=>sum+(by[w]?.[s]||0),0));
      const maxTotal=Math.max(...totals,1);
      const minLabelValue=Math.max(6,Math.ceil(maxTotal*.07));
      const traces=statuses.map(s=>({
        type:'bar',
        name:s,
        x,
        y:x.map(w=>by[w][s]||0),
        marker:{color:statusColors[s]},
        hovertemplate:'%{x}<br>'+s+': %{y}<extra></extra>'
      }));
      traces.push({
        type:'scatter',
        mode:'text',
        name:'Total',
        x,
        y:totals.map(v=>v+Math.max(1.5,maxTotal*.025)),
        text:totals.map(v=>v>=minLabelValue?fmtInt.format(v):''),
        textposition:'top center',
        textfont:{size:12,color:'#222831',family:'Inter, Segoe UI, Arial, sans-serif'},
        hoverinfo:'skip',
        showlegend:false,
        cliponaxis:false
      });
      const layout=baseLayout();
      layout.barmode='stack';
      layout.margin.b=70;
      layout.margin.t=18;
      layout.yaxis.range=[0,maxTotal*1.18];
      layout.yaxis.nticks=6;
      Plotly.react(targetId,traces,layout,{displayModeBar:false,responsive:true});
    }
    function renderStatusChartTo(data,targetId){const counts={'In Progress':0,Completed:0,Rejected:0};data.forEach(d=>counts[d.status]+=1);const labels=statuses, values=labels.map(k=>counts[k]);const trace={type:'pie',labels,values,hole:.66,domain:{x:[.1,.9],y:[.08,.92]},textinfo:'label+percent',textfont:{size:13},marker:{colors:labels.map(k=>statusColors[k])},hovertemplate:'%{label}: %{value}<extra></extra>',sort:false,showlegend:false,automargin:true};const layout={paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',margin:{l:22,r:22,t:24,b:22},showlegend:false};Plotly.react(targetId,[trace],layout,{displayModeBar:false,responsive:true});}
    function buildOwnerSummary(data){const map={};data.forEach(d=>{if(!map[d.owner])map[d.owner]={owner:d.owner,supervisor:d.supervisor||'',total:0,inprogress:0,completed:0,rejected:0,savings:0};map[d.owner].total++;if(d.status==='In Progress')map[d.owner].inprogress++;if(d.status==='Completed')map[d.owner].completed++;if(d.status==='Rejected')map[d.owner].rejected++;map[d.owner].savings+=Number(d.completedSavings)||0;});return Object.values(map).sort((a,b)=>b.total-a.total||a.owner.localeCompare(b.owner));}

    function renderOwnerStackedChart(data,targetId,options={}){
        const rows=buildOwnerSummary(data);
        const showAll=Boolean(options.showAll);
        const visible=(showAll?rows:rows.slice(0,10)).slice().reverse();

        const y=visible.map(r=>r.owner);
        const maxTotal=visible.length?Math.max(...visible.map(r=>r.total)):1;

        const el=document.getElementById(targetId);
        if(el){
            el.style.height=showAll?`${Math.max(500,visible.length*34+150)}px`:'390px';
        }

        const traces=[
            {
            type:'bar',
            orientation:'h',
            name:'Completed',
            x:visible.map(r=>r.completed),
            y,
            text:visible.map(r=>r.completed>0?String(r.completed):''),
            textposition:'inside',
            textangle:0,
            insidetextanchor:'middle',
            constraintext:'none',
            marker:{color:statusColors.Completed},
            hovertemplate:'%{y}<br>Completed: %{x}<extra></extra>'
            },
            {
            type:'bar',
            orientation:'h',
            name:'In Progress',
            x:visible.map(r=>r.inprogress),
            y,
            text:visible.map(r=>r.inprogress>0?String(r.inprogress):''),
            textposition:'inside',
            textangle:0,
            insidetextanchor:'middle',
            constraintext:'none',
            marker:{color:statusColors['In Progress']},
            hovertemplate:'%{y}<br>In Progress: %{x}<extra></extra>'
            },
            {
            type:'bar',
            orientation:'h',
            name:'Rejected',
            x:visible.map(r=>r.rejected),
            y,
            text:visible.map(r=>r.rejected>0?String(r.rejected):''),
            textposition:'inside',
            textangle:0,
            insidetextanchor:'middle',
            constraintext:'none',
            marker:{color:statusColors.Rejected},
            hovertemplate:'%{y}<br>Rejected: %{x}<extra></extra>'
            },
            {
            type:'scatter',
            mode:'text',
            x:visible.map(r=>r.total+Math.max(.2,maxTotal*.025)),
            y,
            text:visible.map(r=>String(r.total)),
            textposition:'middle left',
            textfont:{size:12,color:'#111317'},
            hoverinfo:'skip',
            showlegend:false
            }
        ];

        const layout=baseLayout();
        layout.barmode='stack';
        layout.margin.l=showAll?210:185;
        layout.margin.r=60;
        layout.xaxis.range=[0,maxTotal*1.18];
        layout.legend={orientation:'h',y:1.08,x:0,bgcolor:'rgba(0,0,0,0)'};

        Plotly.react(targetId,traces,layout,{displayModeBar:false,responsive:true});
        }
    function renderSavingsChartTo(data,targetId,options={}){const rows=buildOwnerSummary(data).sort((a,b)=>b.savings-a.savings||b.completed-a.completed||b.total-a.total||a.owner.localeCompare(b.owner));const visible=(options.showAll?rows:rows.slice(0,10)).slice().reverse();const actual=visible.map(r=>r.savings), display=visible.map(r=>r.savings>0?r.savings:.0001), y=visible.map(r=>r.owner), maxVal=actual.length?Math.max(...actual,0):0;const el=document.getElementById(targetId);if(el)el.style.height=options.showAll?`${Math.max(500,visible.length*34+150)}px`:'390px';const trace={type:'bar',orientation:'h',name:'Completed Savings',x:display,y,customdata:actual,marker:{color:'#2f4858'},text:actual.map(v=>fmtMoney.format(v)),textposition:'outside',cliponaxis:false,hovertemplate:'%{y}<br>Completed Savings: $%{customdata:,.0f}<extra></extra>'};const layout=baseLayout();layout.margin.l=options.showAll?210:185;layout.margin.r=80;layout.xaxis.range=[0,maxVal>0?maxVal*1.22:1];layout.xaxis.tickprefix='$';layout.legend={orientation:'h',y:1.08,x:0,bgcolor:'rgba(0,0,0,0)'};Plotly.react(targetId,[trace],layout,{displayModeBar:false,responsive:true});}
    function fillOwnerSummaryTable(targetId,data){const rows=buildOwnerSummary(data);document.getElementById(targetId).innerHTML=rows.map(r=>`<tr><td><strong>${escapeHtml(r.owner)}</strong></td><td>${escapeHtml(r.supervisor)}</td><td><strong>${fmtInt.format(r.total)}</strong></td><td><span class="count-badge count-inprogress">${fmtInt.format(r.inprogress)}</span></td><td><span class="count-badge count-completed">${fmtInt.format(r.completed)}</span></td><td><span class="count-badge count-rejected">${fmtInt.format(r.rejected)}</span></td><td style="text-align:right;font-weight:700">${fmtMoney.format(r.savings)}</td></tr>`).join('')}
    function fillSavingsSummaryTable(targetId,data){const rows=buildOwnerSummary(data).sort((a,b)=>b.savings-a.savings||b.completed-a.completed||b.total-a.total||a.owner.localeCompare(b.owner));document.getElementById(targetId).innerHTML=rows.map(r=>`<tr><td><strong>${escapeHtml(r.owner)}</strong></td><td>${escapeHtml(r.supervisor)}</td><td><strong>${fmtInt.format(r.total)}</strong></td><td><span class="count-badge count-completed">${fmtInt.format(r.completed)}</span></td><td style="text-align:right;font-weight:700">${fmtMoney.format(r.savings)}</td><td><span class="count-badge count-inprogress">${fmtInt.format(r.inprogress)}</span></td><td><span class="count-badge count-rejected">${fmtInt.format(r.rejected)}</span></td></tr>`).join('')}
    function moveModalToBody(modal){if(modal && modal.parentElement!==document.body){document.body.appendChild(modal)}}
    function setupTopScrollbar(barId,wrapId,innerId){const bar=document.getElementById(barId),wrap=document.getElementById(wrapId),inner=document.getElementById(innerId);if(!bar||!wrap||!inner)return;const table=wrap.querySelector('table');const syncWidth=()=>{const total=table?table.scrollWidth:0;const visible=wrap.clientWidth;inner.style.width=`${Math.max(total,visible)}px`;bar.classList.toggle('hidden',total<=visible+8);};let lock=false;bar.onscroll=()=>{if(lock)return;lock=true;wrap.scrollLeft=bar.scrollLeft;requestAnimationFrame(()=>lock=false)};wrap.onscroll=()=>{if(lock)return;lock=true;bar.scrollLeft=wrap.scrollLeft;requestAnimationFrame(()=>lock=false)};syncWidth();window.addEventListener('resize',syncWidth);setTimeout(syncWidth,60);setTimeout(syncWidth,320);}
    function setupAllTopScrollbars(){setupTopScrollbar('detailTopScroll','detailTableWrap','detailTopScrollInner');setupTopScrollbar('pendingTopScroll','pendingTableWrap','pendingTopScrollInner');}
    function openOwnerChartModal(data){const modal=document.getElementById('ownerChartModal');moveModalToBody(modal);fillOwnerSummaryTable('ownerChartModalTableBody',data);modal.style.display='flex';setTimeout(()=>{renderOwnerStackedChart(data,'ownerFullChart',{showAll:true,showLegend:true});Plotly.Plots.resize(document.getElementById('ownerFullChart'));},180)}
    function closeOwnerChartModal(){document.getElementById('ownerChartModal').style.display='none'}
    function openSavingsChartModal(data){const modal=document.getElementById('savingsChartModal');moveModalToBody(modal);fillSavingsSummaryTable('savingsChartModalTableBody',data);modal.style.display='flex';setTimeout(()=>{renderSavingsChartTo(data,'savingsFullChart',{showAll:true});Plotly.Plots.resize(document.getElementById('savingsFullChart'));},180)}
    function closeSavingsChartModal(){document.getElementById('savingsChartModal').style.display='none'}
    let activeDetailTab='all';
    function normalizeRef(value){return String(value||'').trim();}
    function todayLocalDate(){const d=new Date();return new Date(d.getFullYear(),d.getMonth(),d.getDate());}
    function parseDateSafe(value){if(!value)return null;const d=new Date(String(value)+'T00:00:00');return isNaN(d.getTime())?null:d;}
    function daysPending(d){const start=parseDateSafe(d.submittedDate||d.eventDate);if(!start)return 0;return Math.max(0,Math.floor((todayLocalDate()-start)/(24*60*60*1000)));}
    function shortText(value,max=46){const text=String(value||'');return text.length>max?text.slice(0,max-1)+'…':text;}
    function baseDetailRowHtml(d){
      const otherText=String(d.gkType||'').trim().toLowerCase()==='other'?(d.gkOther||''):'';
      const ref=escapeHtml(d.reference);
      return `<tr class="clickable-row" data-reference="${ref}">
        <td class="col-ref">${escapeHtml(d.reference)}</td>
        <td class="title-cell col-title"><div class="cell-clip title-text" title="${escapeHtml(d.title)}">${escapeHtml(d.title)}</div></td>
        <td class="col-supervisor"><div class="cell-clip">${escapeHtml(d.supervisor)}</div></td>
        <td class="col-owner"><div class="cell-clip">${escapeHtml(d.ownerRaw)}</div></td>
        <td class="col-submitter"><div class="cell-clip">${escapeHtml(d.submitter)}</div></td>
        <td class="col-status"><span class="status-pill ${statusClass(d.status)}">${escapeHtml(d.status)}</span></td>
        <td class="col-type"><div class="cell-clip">${escapeHtml(d.gkType)}</div></td>
        <td class="col-other"><div class="cell-clip">${escapeHtml(otherText)}</div></td>
        <td class="col-month">${escapeHtml(d.eventMonth)}</td>
        <td class="col-week">${escapeHtml(d.eventWeek)}</td>
        <td class="col-date">${escapeHtml(d.submittedDate)}</td>
        <td class="col-date">${escapeHtml(d.completedDate)}</td>
        <td class="col-savings">${fmtMoney.format(Number(d.completedSavings)||0)}</td>
        <td class="col-before"><div class="cell-clip" title="${escapeHtml(d.before)}">${escapeHtml(shortText(d.before,70))}</div></td>
        <td class="col-after"><div class="cell-clip" title="${escapeHtml(d.after)}">${escapeHtml(shortText(d.after,70))}</div></td>
      </tr>`;
    }
    function bindDetailRows(targetId){document.querySelectorAll(`#${targetId} tr[data-reference]`).forEach(row=>row.addEventListener('click',()=>openGkDetailByRef(row.dataset.reference)));}
    function renderTable(data){
      const rows=data.slice().sort((a,b)=>(b.eventDate||'').localeCompare(a.eventDate||'')||(b.reference||'').localeCompare(a.reference||''));
      const body=document.getElementById('tableBody');
      body.innerHTML=rows.map(baseDetailRowHtml).join('');
      bindDetailRows('tableBody');
    }
    function renderPendingTab(data){
      const rows=data.filter(d=>d.status==='In Progress').map(d=>({...d,pendingDays:daysPending(d)})).sort((a,b)=>b.pendingDays-a.pendingDays||(b.submittedDate||'').localeCompare(a.submittedDate||''));
      const body=document.getElementById('pendingTableBody');
      document.getElementById('pendingBadge').textContent=`${fmtInt.format(rows.length)} cases`;
      if(!rows.length){
        body.innerHTML='<tr><td colspan="6"><div class="empty-state">No pending GK in the current filter.</div></td></tr>';
        Plotly.purge('pendingDurationChart');
        return;
      }
      body.innerHTML=rows.map(d=>`<tr class="clickable-row" data-reference="${escapeHtml(d.reference)}"><td class="col-ref">${escapeHtml(d.reference)}</td><td class="title-cell col-title"><div class="cell-clip title-text" title="${escapeHtml(d.title)}">${escapeHtml(d.title)}</div></td><td class="col-supervisor">${escapeHtml(d.supervisor)}</td><td class="col-owner">${escapeHtml(d.ownerRaw)}</td><td class="col-status"><span class="status-pill ${statusClass(d.status)}">${escapeHtml(d.status)}</span></td><td class="col-pending"><strong>${fmtInt.format(d.pendingDays)}</strong></td></tr>`).join('');
      bindDetailRows('pendingTableBody');
      const top=rows.slice(0,12).reverse();
      const x=top.map(d=>d.pendingDays);
      const y=top.map(d=>`${shortSupervisorName(d.supervisor)} · ${d.reference}`);
      const colors=top.map(d=>supervisorColors[d.supervisor]||'#2f4858');
      const customdata=top.map(d=>[d.supervisor,d.reference,d.title,d.ownerRaw]);
      const trace={type:'bar',orientation:'h',x,y,marker:{color:colors,line:{color:'#ffffff',width:1.2}},text:x.map(v=>`${fmtInt.format(v)}d`),textposition:'outside',cliponaxis:false,hovertemplate:'<b>%{customdata[0]}</b><br>Ref: %{customdata[1]}<br>Title: %{customdata[2]}<br>GK Owner: %{customdata[3]}<br>Pending: %{x} days<extra></extra>',customdata};
      const layout=baseLayout();layout.margin={l:170,r:78,t:4,b:36};layout.xaxis.title='Pending days';layout.xaxis.gridcolor='#edf1f5';layout.xaxis.zeroline=false;layout.xaxis.range=[0,Math.max(5,...x)*1.18];layout.yaxis.automargin=true;layout.yaxis.gridcolor='rgba(0,0,0,0)';layout.height=360;layout.showlegend=false;
      Plotly.react('pendingDurationChart',[trace],layout,{displayModeBar:false,responsive:true});
    }
    function renderDetailTabs(data){renderTable(data);renderPendingTab(data);}
    function setDetailTab(tab){activeDetailTab=tab==='pending'?'pending':'all';document.querySelectorAll('.detail-tab-btn').forEach(btn=>btn.classList.toggle('active',btn.dataset.detailTab===activeDetailTab));document.querySelectorAll('.detail-tab-panel').forEach(panel=>panel.classList.remove('active'));const target=document.getElementById(activeDetailTab==='pending'?'tabPending':'tabAll');if(target)target.classList.add('active');setTimeout(()=>{try{Plotly.Plots.resize(document.getElementById('pendingDurationChart'));}catch(e){}},80);}
    function detailField(label,value,isLong=false){return `<div class="case-label">${escapeHtml(label)}</div><div class="case-value ${isLong?'long':''}">${escapeHtml(value||'')}</div>`;}
    function openGkDetailByRef(ref){
      const target=normalizeRef(ref);const d=rawData.find(r=>normalizeRef(r.reference)===target);if(!d)return;
      const modal=document.getElementById('gkDetailModal');moveModalToBody(modal);
      document.getElementById('caseModalTitle').textContent=`${d.reference||''} - ${shortText(d.title,70)}`;
      const otherText=String(d.gkType||'').trim().toLowerCase()==='other'?(d.gkOther||''):'';
      document.getElementById('caseDetailFields').innerHTML=[
        detailField('Reference #',d.reference),detailField('Title',d.title,true),detailField('Supervisor',d.supervisor),detailField('GK Owner',d.ownerRaw),detailField('Submitter',d.submitter),detailField('Status',d.status),detailField('GK Type',d.gkType),detailField('GK Other',otherText,true),detailField('Submitted Date',d.submittedDate),detailField('Completed Date',d.completedDate),detailField('Event Date',d.eventDate),detailField('Savings',fmtMoney.format(Number(d.completedSavings)||0)),detailField('Before',d.before,true),detailField('After',d.after,true)
      ].join('');
      modal.style.display='flex';
    }
    function closeGkDetailModal(){document.getElementById('gkDetailModal').style.display='none';}
    function updateActiveFilters(){const chips=[];if(showSupervisorFilter && state.supervisors.size!==allSupervisors.length)chips.push(`<span class="chip">Sup: ${state.supervisors.size} selected</span>`);if(state.owners.size!==allOwners.length)chips.push(`<span class="chip">Owner: ${state.owners.size} selected</span>`);if(state.statuses.size!==statuses.length)chips.push(`<span class="chip">Status: ${state.statuses.size} selected</span>`);if(state.gkTypes.size!==allGkTypes.length)chips.push(`<span class="chip">GK Type: ${state.gkTypes.size} selected</span>`);if(state.months.size!==allMonths.length)chips.push(`<span class="chip">Month: ${state.months.size} selected</span>`);if(state.weeks.size!==weekOrder.length)chips.push(`<span class="chip">Week: ${state.weeks.size} selected</span>`);if(state.dateFrom||state.dateTo)chips.push(`<span class="chip">Date: ${state.dateFrom||'...'} → ${state.dateTo||'...'}</span>`);document.getElementById('activeFilters').innerHTML=chips.join('')}
    function refresh(){const data=getFilteredData();updateSummaries();updateKpis(data);updateActiveFilters();renderWeeklyChartTo(data,'weeklyChart');renderStatusChartTo(data,'statusChart');renderOwnerStackedChart(data,'ownerChart');renderSavingsChartTo(data,'savingsChart');renderDetailTabs(data);setTimeout(setupAllTopScrollbars,30)}
    function getCheckedValues(filter){
        return new Set([...document.querySelectorAll(`input[data-filter="${filter}"]:checked`)].map(el=>el.value));
        }

        function setLinkedOptionVisibility(filter,allowedSet){
        document.querySelectorAll(`input[data-filter="${filter}"]`).forEach(cb=>{
            const item=cb.closest('.check-item');
            if(!item)return;
            const visible=allowedSet.has(cb.value);
            item.dataset.linkVisible=visible?'1':'0';

            const panel=item.closest('.option-panel');
            const search=panel?panel.querySelector('.filter-search'):null;
            const q=search?search.value.trim().toLowerCase():'';
            const matched=!q||item.textContent.toLowerCase().includes(q);

            item.style.display=(visible&&matched)?'flex':'none';
        });
        }

        function syncMonthWeekOptionVisibility(changedFilter=''){
        const selectedMonths=getCheckedValues('month');
        const selectedWeeks=getCheckedValues('week');

        const monthIsAll=selectedMonths.size===allMonths.length;
        const weekIsAll=selectedWeeks.size===weekOrder.length;

        const allowedWeeks=new Set();
        const allowedMonths=new Set();

        rawData.forEach(d=>{
            const m=d.eventMonth||'';
            const w=d.eventWeek||'';
            if(!m||!w)return;

            if(monthIsAll||selectedMonths.has(m))allowedWeeks.add(w);
            if(weekIsAll||selectedWeeks.has(w))allowedMonths.add(m);
        });

        setLinkedOptionVisibility('week',allowedWeeks);
        setLinkedOptionVisibility('month',allowedMonths);
        }

        function syncSupervisorOwnerOptionVisibility(){
        if(!showSupervisorFilter)return;
        const selectedSupervisors=getCheckedValues('supervisor');
        const selectedOwners=getCheckedValues('owner');
        const supervisorIsAll=selectedSupervisors.size===allSupervisors.length;
        const ownerIsAll=selectedOwners.size===allOwners.length;
        const allowedOwners=new Set();
        const allowedSupervisors=new Set();

        rawData.forEach(d=>{
            const sup=d.supervisor||'';
            const owner=d.owner||'';
            if(owner&&(supervisorIsAll||selectedSupervisors.has(sup)))allowedOwners.add(owner);
            if(sup&&(ownerIsAll||selectedOwners.has(owner)))allowedSupervisors.add(sup);
        });

        setLinkedOptionVisibility('owner',allowedOwners);
        setLinkedOptionVisibility('supervisor',allowedSupervisors);
        }
    function syncStateFromInputs(changedFilter=''){
        if(showSupervisorFilter){
            state.supervisors=new Set([...document.querySelectorAll('input[data-filter="supervisor"]:checked')].map(el=>el.value));
        }
        state.owners=new Set([...document.querySelectorAll('input[data-filter="owner"]:checked')].map(el=>el.value));
        state.statuses=new Set([...document.querySelectorAll('input[data-filter="status"]:checked')].map(el=>el.value));
        state.gkTypes=new Set([...document.querySelectorAll('input[data-filter="gkType"]:checked')].map(el=>el.value));
        state.months=new Set([...document.querySelectorAll('input[data-filter="month"]:checked')].map(el=>el.value));
        state.weeks=new Set([...document.querySelectorAll('input[data-filter="week"]:checked')].map(el=>el.value));
        state.dateFrom=document.getElementById('dateFrom').value;
        state.dateTo=document.getElementById('dateTo').value;

        syncMonthWeekOptionVisibility(changedFilter);
        syncSupervisorOwnerOptionVisibility();

        refresh();
        }
    function toggleFilterSet(filter,checked){
        document.querySelectorAll(`input[data-filter="${filter}"]`).forEach(el=>{
            if(el.closest('.check-item')?.dataset.linkVisible==='0')return;
            el.checked=checked;
        });
        syncStateFromInputs(filter);
        }
    function resetFilters(){
        document.querySelectorAll('input[data-filter]').forEach(el=>el.checked=true);
        document.querySelectorAll('.check-item').forEach(item=>{
            item.dataset.linkVisible='1';
            item.style.display='flex';
        });
        document.querySelectorAll('.filter-search').forEach(input=>input.value='');
        document.getElementById('dateFrom').value='';
        document.getElementById('dateTo').value='';
        syncStateFromInputs();
        }
    function closeOpenFilters(except=null){document.querySelectorAll('details.multi[open]').forEach(d=>{if(d!==except)d.open=false})}
    function initDropdownSearch(){
    document.querySelectorAll('.filter-search').forEach(input=>{
        input.addEventListener('input',()=>{
        const target=document.getElementById(input.dataset.target);
        if(!target)return;
        const q=input.value.trim().toLowerCase();
        target.querySelectorAll('.check-item').forEach(item=>{
            const linkVisible=item.dataset.linkVisible!=='0';
            const matched=item.textContent.toLowerCase().includes(q);
            item.style.display=(linkVisible&&matched)?'flex':'none';
        });
        });
    });
    }
    function exportFilteredCsv(){const rows=getFilteredData();const headers=['reference','title','supervisor','owner','submitter','status','gkType','gkOther','eventMonth','eventWeek','submittedDate','completedDate','completedSavings','before','after'];const csv=[headers.join(',')].concat(rows.map(r=>headers.map(h=>'"'+String(r[h]??'').replace(/"/g,'""')+'"').join(','))).join('\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='gk_filtered_export.csv';a.click();URL.revokeObjectURL(url)}
    function closeAllDropdownsBeforeExport(){
        document.querySelectorAll('details[open]').forEach(d=>d.removeAttribute('open'));
        const exportMenu=document.getElementById('exportMenu');
        if(exportMenu)exportMenu.open=false;
    }

    async function exportDashboard(format){
        closeAllDropdownsBeforeExport();

        const node=document.getElementById('reportForPdf')||document.body;
        const originalScrollY=window.scrollY;
        const app=document.querySelector('.app');
        const oldTransform=app?app.style.transform:'';
        const oldWidth=app?app.style.width:'';
        const oldMinHeight=app?app.style.minHeight:'';

        if(typeof html2canvas==='undefined'){
            alert('Export library is not loaded. Please wait and try again.');
            return;
        }

        window.scrollTo(0,0);
        document.body.classList.add('pdf-mode');

        if(app){
            app.style.transform='none';
            app.style.width='100vw';
            app.style.minHeight='auto';
        }

        await new Promise(resolve=>setTimeout(resolve,650));

        try{
            await Promise.all(
                Array.from(node.querySelectorAll('.js-plotly-plot')).map(plot=>{
                    try{return Plotly.Plots.resize(plot);}catch(e){return null;}
                })
            );

            const canvas=await html2canvas(node,{
                scale:2,
                useCORS:true,
                allowTaint:true,
                backgroundColor:'#f3f5f7',
                windowWidth:Math.max(node.scrollWidth, document.documentElement.clientWidth),
                windowHeight:Math.max(node.scrollHeight, document.documentElement.clientHeight),
                scrollX:0,
                scrollY:0,
                ignoreElements:(el)=>{
                    return el.classList&&(
                        el.classList.contains('modal-backdrop')||
                        el.classList.contains('option-panel')||
                        el.classList.contains('export-menu-panel')
                    );
                }
            });

            const imgData=canvas.toDataURL('image/png');

            if(format==='png'){
                const a=document.createElement('a');
                a.href=imgData;
                a.download='GK_Dashboard_Report.png';
                document.body.appendChild(a);
                a.click();
                a.remove();
                return;
            }

            if(format==='pdf'){
                if(!window.jspdf||!window.jspdf.jsPDF){
                    window.print();
                    return;
                }

                const {jsPDF}=window.jspdf;
                const pdf=new jsPDF('l','mm','a4');

                const pageWidth=pdf.internal.pageSize.getWidth();
                const pageHeight=pdf.internal.pageSize.getHeight();

                const imgWidth=pageWidth;
                const imgHeight=canvas.height*imgWidth/canvas.width;

                let heightLeft=imgHeight;
                let position=0;

                pdf.addImage(imgData,'PNG',0,position,imgWidth,imgHeight);
                heightLeft-=pageHeight;

                while(heightLeft>0){
                    position=heightLeft-imgHeight;
                    pdf.addPage();
                    pdf.addImage(imgData,'PNG',0,position,imgWidth,imgHeight);
                    heightLeft-=pageHeight;
                }

                pdf.save('GK_Dashboard_Report.pdf');
            }
        }catch(err){
            console.error(err);
            if(format==='pdf'){
                window.print();
            }else{
                alert('PNG export failed in this browser.');
            }
        }finally{
            document.body.classList.remove('pdf-mode');
            if(app){
                app.style.transform=oldTransform;
                app.style.width=oldWidth;
                app.style.minHeight=oldMinHeight;
            }
            window.scrollTo(0,originalScrollY);
            setTimeout(()=>{
                document.querySelectorAll('.js-plotly-plot').forEach(plot=>{
                    try{Plotly.Plots.resize(plot);}catch(e){}
                });
            },100);
        }
    }
    function bindIfExists(id,eventName,handler){const el=document.getElementById(id);if(el)el.addEventListener(eventName,handler)}
    document.querySelectorAll('input[data-filter]').forEach(el=>el.addEventListener('change',()=>syncStateFromInputs(el.dataset.filter)));document.getElementById('dateFrom').addEventListener('change',syncStateFromInputs);document.getElementById('dateTo').addEventListener('change',syncStateFromInputs);document.getElementById('resetFilters').addEventListener('click',resetFilters);bindIfExists('exportPdfOption','click',()=>exportDashboard('pdf'));bindIfExists('exportPngOption','click',()=>exportDashboard('png'));document.getElementById('exportCsv').addEventListener('click',exportFilteredCsv);bindIfExists('viewAllOwnersChartBtn','click',()=>{closeOpenFilters();openOwnerChartModal(getFilteredData())});bindIfExists('viewAllSavingsChartBtn','click',()=>{closeOpenFilters();openSavingsChartModal(getFilteredData())});bindIfExists('closeOwnerChartModal','click',closeOwnerChartModal);bindIfExists('closeSavingsChartModal','click',closeSavingsChartModal);bindIfExists('ownerChartModal','click',e=>{if(e.target.id==='ownerChartModal')closeOwnerChartModal()});bindIfExists('savingsChartModal','click',e=>{if(e.target.id==='savingsChartModal')closeSavingsChartModal()});
    document.querySelectorAll('.tiny-btn[data-action]').forEach(btn=>{btn.addEventListener('click',e=>{e.preventDefault();toggleFilterSet(btn.getAttribute('data-filter'),btn.getAttribute('data-action')==='all')})});document.querySelectorAll('[data-close-filter="true"]').forEach(btn=>{btn.addEventListener('click',e=>{e.preventDefault();const details=btn.closest('details.multi');if(details)details.open=false})});document.querySelectorAll('details.multi').forEach(details=>{details.addEventListener('toggle',()=>{if(details.open)closeOpenFilters(details)})});document.addEventListener('click',e=>{if(!e.target.closest('details.multi'))closeOpenFilters();if(!e.target.closest('#exportMenu')){const exportMenu=document.getElementById('exportMenu');if(exportMenu)exportMenu.open=false}});
    document.querySelectorAll('.detail-tab-btn').forEach(btn=>btn.addEventListener('click',()=>setDetailTab(btn.dataset.detailTab||'all')));
    bindIfExists('closeGkDetailModal','click',closeGkDetailModal);
    bindIfExists('gkDetailModal','click',e=>{if(e.target.id==='gkDetailModal')closeGkDetailModal()});
    document.addEventListener('keydown',e=>{if(e.key==='Escape'&&document.getElementById('gkDetailModal')?.style.display==='flex')closeGkDetailModal()});
    initDropdownSearch();
    resetFilters();

  </script>
</body>
</html>'''

    hero_logo = ''
    hero_watermark = ''
    if logo_data_uri:
        hero_logo = f'<div class="hero-logo-card"><img src="{logo_data_uri}" alt="Milwaukee logo" class="hero-logo" /></div>'
        hero_watermark = f'<div class="hero-watermark"><img src="{logo_data_uri}" alt="Milwaukee logo watermark" class="hero-watermark-img" /></div>'

    replacements = {
        '__PLOTLY_JS__': plotly_js,
        '__RAW_DATA__': json.dumps(data_records, ensure_ascii=False),
        '__STATUS_COLORS__': json.dumps(STATUS_COLORS),
        '__WEEK_ORDER__': json.dumps(weeks),
        '__OWNERS_JSON__': json.dumps(owners, ensure_ascii=False),
        '__SUPERVISORS_JSON__': json.dumps(supervisors, ensure_ascii=False),
        '__SHOW_SUPERVISOR_FILTER__': 'true' if show_supervisor_filter and bool(supervisors) else 'false',
        '__FILTER_GRID_CLASS__': '' if show_supervisor_filter and supervisors else 'no-supervisor',
        '__SUPERVISOR_FILTER_BLOCK__': (f'<div class="filter-block"><label class="title">Supervisor</label><details class="multi"><summary><span>Select Sup</span><span class="summary-count" id="supervisorSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="supervisor">All</button><button class="tiny-btn" data-action="none" data-filter="supervisor">None</button><button class="tiny-btn option-back" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="supervisorOptions">{_checkbox_html("supervisor", supervisors)}</div></div></details></div>' if show_supervisor_filter and supervisors else ''),
        '__MONTHS_JSON__': json.dumps(months, ensure_ascii=False),
        '__GK_TYPES_JSON__': json.dumps(gk_types, ensure_ascii=False),
        '__OWNER_CHECKBOXES__': _checkbox_html('owner', owners),
        '__SUPERVISOR_CHECKBOXES__': _checkbox_html('supervisor', supervisors),
        '__STATUS_CHECKBOXES__': _checkbox_html('status', statuses),
        '__GK_TYPE_CHECKBOXES__': _checkbox_html('gkType', gk_types),
        '__MONTH_CHECKBOXES__': _checkbox_html('month', months),
        '__WEEK_CHECKBOXES__': _checkbox_html('week', weeks),
        '__HERO_LOGO__': hero_logo,
        '__HERO_WATERMARK__': hero_watermark,
        '__LATEST_UPDATE_LINE__': f'<div class="hero-update">Latest update: <span>{html.escape(latest_update_text)}</span></div>' if latest_update_text else '',
    }
    for key, value in replacements.items():
        template = template.replace(key, value)
    return template



def render_home_dashboard(data_records, source_name: str, logo_data_uri: str = '', latest_update_text: str = '') -> str:
    plotly_js = get_plotlyjs()
    statuses = ['Completed', 'In Progress', 'Rejected']
    gk_types = sorted({r.get('gkTypeFilter', '(Blank)') for r in data_records}) or ['(Blank)']
    supervisors = sorted({(r.get('supervisor') or '').strip() for r in data_records if str(r.get('supervisor') or '').strip()}) or ['(Blank)']
    all_weeks = [wk['label'] for wk in WEEKS_2026]
    week_set = {r.get('eventWeek') for r in data_records if r.get('eventWeek') in all_weeks}
    weeks = [wk for wk in all_weeks if wk in week_set] or all_weeks
    month_set = {r.get('eventMonth') for r in data_records if r.get('eventMonth') and r.get('eventMonth') != 'No Month'}
    months = sorted(month_set) or ['No Month']

    template = r'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>GK Dashboard App Home</title>
  <style>
    :root{
      --bg:#eef1f4;--panel:#fff;--panel-2:#fbfcfd;--ink:#111317;--muted:#66707d;--muted-2:#8b95a1;
      --line:#e1e6ec;--line-2:#edf1f5;--shadow:0 14px 34px rgba(16,24,40,.085);--shadow-hover:0 18px 46px rgba(16,24,40,.135);
      --red:#c41230;--red-2:#9f102a;--black:#0f1319;--green:#229954;--amber:#f5a623;--blue:#2f80ed;--teal:#00a39a;--purple:#9b51e0;
      --radius:20px;--home-scale:.82;--home-width:121.951vw;
    }
    *{box-sizing:border-box}
    html,body{margin:0;padding:0;width:100%;min-height:100%;font-family:Inter,Segoe UI,Arial,sans-serif;color:var(--ink);background:var(--bg);overflow-x:hidden;scroll-behavior:smooth}
    .home-viewport{width:100vw;min-height:100vh;overflow-x:hidden;background:linear-gradient(180deg,#f4f6f8 0%,#eef1f4 100%)}
    .home-app{width:var(--home-width);min-height:100vh;transform:scale(var(--home-scale));transform-origin:top left;animation:pageEnter .42s ease both}
    body.page-leave .home-app{animation:pageLeave .16s ease both}
    @keyframes pageEnter{from{opacity:0;transform:scale(var(--home-scale)) translateY(12px)}to{opacity:1;transform:scale(var(--home-scale)) translateY(0)}}
    @keyframes pageLeave{to{opacity:0;transform:scale(var(--home-scale)) translateY(-8px)}}

    .hero{position:relative;overflow:hidden;background:radial-gradient(circle at 80% -20%,rgba(196,18,48,.34),transparent 34%),linear-gradient(120deg,#070b10 0%,#111923 46%,#4d0a1b 100%);color:#fff;padding:24px 34px 30px;border-bottom:5px solid var(--red)}
    .hero:before{content:'';position:absolute;inset:auto auto 0 34px;width:280px;height:4px;background:linear-gradient(90deg,#ff3959,rgba(255,57,89,0));border-radius:999px}
    .hero:after{content:'';position:absolute;right:190px;top:-110px;width:360px;height:360px;border-radius:50%;background:radial-gradient(circle,rgba(196,18,48,.17),transparent 62%);pointer-events:none}
    .hero-grid{position:absolute;right:0;bottom:0;width:520px;height:150px;opacity:.13;background-image:radial-gradient(rgba(255,255,255,.45) 1px,transparent 1px);background-size:10px 10px;mask-image:linear-gradient(90deg,transparent,#000 24%,#000)}
    .hero-top{position:relative;z-index:1;display:flex;justify-content:space-between;gap:24px;align-items:flex-start}
    .hero-kicker{font-size:11px;text-transform:uppercase;letter-spacing:.20em;color:rgba(255,255,255,.62);font-weight:950;margin-bottom:8px}
    .hero h1{margin:0;font-size:36px;line-height:1.04;letter-spacing:-.045em;text-shadow:0 10px 24px rgba(0,0,0,.24)}
    .hero-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}
    .hero-logo-card{flex:0 0 auto;width:350px;max-width:350px;background:rgba(255,255,255,.075);border:1px solid rgba(255,255,255,.16);border-radius:18px;padding:14px 22px;box-shadow:0 18px 42px rgba(0,0,0,.20);backdrop-filter:blur(8px)}
    .hero-logo{display:block;width:100%;height:auto;max-height:100px;object-fit:contain}

    .page{padding:18px 22px 26px;max-width:1720px;margin:0 auto}
    .btn{border:0;border-radius:13px;padding:11px 16px;font-weight:950;cursor:pointer;transition:transform .18s ease,box-shadow .18s ease,background .18s ease,text-shadow .18s ease;text-decoration:none;display:inline-flex;align-items:center;justify-content:center;gap:8px;font-size:14px;letter-spacing:.01em}
    .btn:hover{transform:translateY(-2px);box-shadow:0 12px 26px rgba(16,24,40,.18)}
    .btn-primary{background:linear-gradient(180deg,#d71337,#b70f2d);color:#fff}.btn-primary:hover{background:linear-gradient(180deg,#e0193e,#9f102a)}
    .btn-dark{background:rgba(255,255,255,.08);color:#fff;border:1px solid rgba(255,255,255,.32)}.btn-dark:hover{background:rgba(255,255,255,.14)}
    .btn-ghost{background:#f2f4f7;color:#202631;border:1px solid #e6ebf0}.btn-ghost:hover{background:#fff}

    .filter-panel{position:sticky;top:0;z-index:20;background:rgba(255,255,255,.94);backdrop-filter:blur(14px);border:1px solid rgba(255,255,255,.98);box-shadow:var(--shadow);border-radius:22px;padding:15px 16px;display:grid;gap:13px;margin-top:-10px}
    .filter-head{display:flex;justify-content:space-between;gap:12px;align-items:center}.filter-title{font-size:19px;font-weight:950;letter-spacing:-.02em}.filter-grid{display:grid;grid-template-columns:1.18fr .82fr .92fr .82fr .82fr 1.35fr;gap:10px;align-items:start}
    .filter-block{background:linear-gradient(180deg,#fff,#fbfcfd);border:1px solid var(--line);border-radius:16px;padding:9px 11px;min-height:72px}
    .filter-block label.title{display:block;font-size:12px;text-transform:uppercase;letter-spacing:.075em;color:#596372;margin-bottom:7px;font-weight:950}
    .filter-block input[type="date"],.filter-block input[type="text"]{width:100%;border:1px solid var(--line);background:#fff;border-radius:12px;padding:9px 10px;font-size:13px;outline:none}.date-range{display:grid;grid-template-columns:1fr 1fr;gap:8px}
    details.multi{position:relative}details.multi summary{list-style:none;cursor:pointer;border:1px solid var(--line);background:#fff;border-radius:12px;padding:10px 12px;display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:14px;color:#1a1a1a}details.multi summary::-webkit-details-marker{display:none}.summary-count{color:var(--muted);font-weight:800}
    .option-panel{position:absolute;left:0;right:0;top:calc(100% + 8px);z-index:80;background:#fff;border:1px solid var(--line);border-radius:16px;box-shadow:var(--shadow);padding:0;max-height:330px;overflow:auto}.option-tools{display:grid;grid-template-columns:1fr auto auto auto;gap:8px;align-items:center;background:#fff;padding:10px 10px 9px;border-bottom:1px solid #edf0f2;position:sticky;top:0;z-index:3}.option-tools.no-search{grid-template-columns:auto auto auto;justify-content:start}.filter-search{width:100%;border:1px solid var(--line);background:#fff;border-radius:10px;padding:7px 9px;font-size:12px;outline:none}.tiny-btn{border:1px solid var(--line);background:#fff;border-radius:10px;padding:6px 8px;font-size:12px;cursor:pointer;white-space:nowrap}.check-list{display:grid;gap:6px;padding:10px}.check-item{display:flex;gap:9px;align-items:center;padding:6px 4px;border-radius:10px;font-size:13px;line-height:1.25}.check-item:hover{background:#f4f6f8}
    .active-filters{display:flex;flex-wrap:wrap;gap:8px}.chip{padding:8px 10px;border-radius:999px;background:#fff1f4;color:var(--red);border:1px solid #ffd0da;font-size:12px;font-weight:900}

    .kpi-grid{margin-top:14px;display:grid;grid-template-columns:1.05fr .95fr .95fr .95fr 1.1fr;gap:14px}
    .kpi-card{background:var(--panel);border:1px solid var(--line);border-radius:19px;padding:15px 17px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:transform .18s ease,box-shadow .18s ease;display:flex;align-items:center;gap:15px;min-height:88px}.kpi-card:hover{transform:translateY(-2px);box-shadow:var(--shadow-hover)}
    .kpi-card.primary{background:linear-gradient(135deg,#0e141c 0%,#131922 100%);color:#fff;border-color:#1d2631}.kpi-card.primary:before{content:'';position:absolute;inset:0 auto 0 0;width:5px;background:linear-gradient(180deg,#ff3f60,#c41230)}
    .kpi-icon{width:44px;height:44px;border-radius:16px;display:flex;align-items:center;justify-content:center;font-size:24px;font-weight:950;flex:0 0 auto;border:2px solid rgba(196,18,48,.65);color:var(--red);background:#fff}.kpi-card.primary .kpi-icon{background:rgba(196,18,48,.12);border-color:rgba(255,57,89,.75);color:#ff3f60}.kpi-icon.green{border-color:rgba(46,139,87,.65);color:#218a4e}.kpi-icon.amber{border-color:rgba(245,166,35,.75);color:#c97b00}.kpi-icon.red{border-color:rgba(196,18,48,.75);color:#c41230}
    .kpi-label{font-size:12px;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);font-weight:950}.kpi-card.primary .kpi-label{color:rgba(255,255,255,.62)}.kpi-value{margin-top:5px;font-size:30px;font-weight:950;letter-spacing:-.035em}.kpi-card.primary .kpi-value{color:#fff}

    .chart-row{margin-top:14px;display:grid;grid-template-columns:minmax(0,3.05fr) minmax(300px,.95fr);gap:14px}.chart-row-single{margin-top:14px;display:grid;grid-template-columns:1fr;gap:14px}.chart-grid{margin-top:14px;display:grid;grid-template-columns:1fr 1fr;gap:14px}
    .card{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:18px 18px 14px;transition:transform .18s ease,box-shadow .18s ease}.card:hover{transform:translateY(-2px);box-shadow:var(--shadow-hover)}
    .card-header{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:8px}.card-title{font-size:20px;font-weight:950;letter-spacing:-.02em}.card-badge{padding:9px 12px;border-radius:999px;background:#f4f6f8;color:#26303d;font-size:13px;font-weight:950;white-space:nowrap}.plot{width:100%;height:390px}.plot.large{height:430px}.plot.compact{height:430px}.plot.mid{height:400px}
    .fade-in{animation:sectionIn .52s ease both}.fade-in:nth-of-type(2){animation-delay:.04s}.fade-in:nth-of-type(3){animation-delay:.08s}.fade-in:nth-of-type(4){animation-delay:.12s}@keyframes sectionIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
    @media(max-width:1100px){:root{--home-scale:1;--home-width:100vw}.home-app{width:100vw;transform:none}.filter-grid,.kpi-grid,.chart-row,.chart-grid{grid-template-columns:1fr}.hero-top{flex-direction:column}.hero-logo-card{width:260px}.page{padding:14px}.filter-panel{position:relative}.hero{padding:24px 20px 30px}}@media(max-width:700px){.date-range{grid-template-columns:1fr}.hero h1{font-size:28px}.kpi-value{font-size:28px}.kpi-card{align-items:flex-start}}
  </style>
</head>
<body>
  <div class="home-viewport"><div class="home-app">
    <header class="hero"><div class="hero-grid"></div>
      <div class="hero-top">
        <div>
          <div class="hero-kicker">Gemba Kaizen Dashboard</div>
          <h1>GK Supervisor Performance Overview</h1>
          <div class="hero-actions"><a class="btn btn-primary page-link" href="official.html">Official Dashboard</a></div>
        </div>
        __HERO_LOGO__
      </div>
    </header>
    <main class="page">
      <section class="filter-panel fade-in">
        <div class="filter-head"><div class="filter-title">Filters</div><button class="btn btn-ghost" id="resetFilters">↻ Reset filters</button></div>
        <div class="filter-grid">
          <div class="filter-block"><label class="title">Supervisor</label><details class="multi"><summary><span>👤 Select supervisor(s)</span><span class="summary-count" id="supervisorSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search supervisor..." class="filter-search" data-target="supervisorOptions"><button class="tiny-btn" data-action="all" data-filter="supervisor">All</button><button class="tiny-btn" data-action="none" data-filter="supervisor">None</button><button class="tiny-btn" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="supervisorOptions">__SUPERVISOR_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Status</label><details class="multi"><summary><span>Select status</span><span class="summary-count" id="statusSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="status">All</button><button class="tiny-btn" data-action="none" data-filter="status">None</button><button class="tiny-btn" type="button" data-close-filter="true">Back</button></div><div class="check-list">__STATUS_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">GK Type</label><details class="multi"><summary><span>Select type</span><span class="summary-count" id="gkTypeSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search type..." class="filter-search" data-target="gkTypeOptions"><button class="tiny-btn" data-action="all" data-filter="gkType">All</button><button class="tiny-btn" data-action="none" data-filter="gkType">None</button><button class="tiny-btn" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="gkTypeOptions">__GK_TYPE_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Month</label><details class="multi"><summary><span>Select month</span><span class="summary-count" id="monthSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="month">All</button><button class="tiny-btn" data-action="none" data-filter="month">None</button><button class="tiny-btn" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="monthOptions">__MONTH_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Week</label><details class="multi"><summary><span>Select week(s)</span><span class="summary-count" id="weekSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search week..." class="filter-search" data-target="weekOptions"><button class="tiny-btn" data-action="all" data-filter="week">All</button><button class="tiny-btn" data-action="none" data-filter="week">None</button><button class="tiny-btn" type="button" data-close-filter="true">Back</button></div><div class="check-list" id="weekOptions">__WEEK_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Date Range</label><div class="date-range"><input type="date" id="dateFrom"/><input type="date" id="dateTo"/></div></div>
        </div><div class="active-filters" id="activeFilters"></div>
      </section>
      <section class="kpi-grid fade-in">
        <div class="kpi-card primary"><div class="kpi-icon">✓</div><div><div class="kpi-label">Total GK</div><div class="kpi-value" id="kpiTotal">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon amber">◌</div><div><div class="kpi-label">In Progress</div><div class="kpi-value" id="kpiInProgress">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon green">✓</div><div><div class="kpi-label">Completed</div><div class="kpi-value" id="kpiCompleted">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon red">×</div><div><div class="kpi-label">Rejected</div><div class="kpi-value" id="kpiRejected">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon red">$</div><div><div class="kpi-label">Completed Savings</div><div class="kpi-value" id="kpiSavings">$0</div></div></div>
      </section>
      <section class="chart-row fade-in"><div class="card"><div class="card-header"><div class="card-title">Monthly Savings by Supervisor</div><div class="card-badge" id="savingsBadge">$0</div></div><div id="monthlySavingsChart" class="plot large"></div></div><div class="card"><div class="card-header"><div class="card-title">Total Savings</div></div><div id="totalSavingsBySupervisorChart" class="plot compact"></div></div></section>
      <section class="chart-row-single fade-in"><div class="card"><div class="card-header"><div class="card-title">Monthly Total GK by Supervisor</div><div class="card-badge" id="totalBadge">0 cases</div></div><div id="monthlyTotalChart" class="plot mid"></div></div></section>
      <section class="chart-grid fade-in"><div class="card"><div class="card-header"><div class="card-title">Status Breakdown by Supervisor</div></div><div id="statusBySupervisorChart" class="plot mid"></div></div><div class="card"><div class="card-header"><div class="card-title">Completion Rate by Supervisor</div><div class="card-badge" id="completionBadge">0%</div></div><div id="completionRateChart" class="plot mid"></div></div></section>
    </main>
  </div></div>
  <script>__PLOTLY_JS__</script>
  <script>
    const rawData=__RAW_DATA__;
    const statusColors=__STATUS_COLORS__;
    const weekOrder=__WEEK_ORDER__;
    const allSupervisors=__SUPERVISORS_JSON__;
    const allMonths=__MONTHS_JSON__;
    const allGkTypes=__GK_TYPES_JSON__;
    const statuses=['Completed','In Progress','Rejected'];
    const state={supervisors:new Set(allSupervisors),statuses:new Set(statuses),gkTypes:new Set(allGkTypes),months:new Set(allMonths),weeks:new Set(weekOrder),dateFrom:'',dateTo:''};
    const fmtInt=new Intl.NumberFormat('en-US',{maximumFractionDigits:0});
    const fmtMoney=new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0});
    const fmtPct=new Intl.NumberFormat('en-US',{maximumFractionDigits:1});
    const supervisorPalette=['#c41230','#2f80ed','#229954','#f2994a','#9b51e0','#00a39a','#eb5757','#2f4858','#f2c94c','#56ccf2'];
    const supervisorColors={};allSupervisors.forEach((s,i)=>{supervisorColors[s]=supervisorPalette[i%supervisorPalette.length]});
    function shortName(name){const parts=String(name||'').trim().split(/\s+/).filter(Boolean);return parts.length?parts[parts.length-1]:''}
    function formatMonth(m){if(!m||m==='No Month')return m;const [y,mo]=String(m).split('-');const names=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];return `${names[(+mo||1)-1]} ${y}`}
    function moneyShort(v){v=Number(v)||0;if(Math.abs(v)>=1000000)return '$'+(v/1000000).toFixed(v%1000000?1:0)+'M';if(Math.abs(v)>=1000)return '$'+Math.round(v/1000)+'k';return '$'+fmtInt.format(v)}
    function updateSummaries(){document.getElementById('supervisorSummary').textContent=state.supervisors.size===allSupervisors.length?'All':`${state.supervisors.size} selected`;document.getElementById('statusSummary').textContent=state.statuses.size===statuses.length?'All':`${state.statuses.size} selected`;document.getElementById('gkTypeSummary').textContent=state.gkTypes.size===allGkTypes.length?'All':`${state.gkTypes.size} selected`;document.getElementById('monthSummary').textContent=state.months.size===allMonths.length?'All':`${state.months.size} selected`;document.getElementById('weekSummary').textContent=state.weeks.size===weekOrder.length?'All':`${state.weeks.size} selected`}
    function getFilteredData(){return rawData.filter(d=>{if(!state.supervisors.has(d.supervisor))return false;if(!state.statuses.has(d.status))return false;if(!state.gkTypes.has(d.gkTypeFilter||'(Blank)'))return false;if(!state.months.has(d.eventMonth))return false;if(!state.weeks.has(d.eventWeek))return false;if(state.dateFrom&&(!d.eventDate||d.eventDate<state.dateFrom))return false;if(state.dateTo&&(!d.eventDate||d.eventDate>state.dateTo))return false;return true})}
    function baseLayout(extra={}){return Object.assign({paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',margin:{l:60,r:30,t:10,b:64},font:{family:'Inter, Segoe UI, Arial, sans-serif',color:'#202631',size:13},xaxis:{gridcolor:'#edf1f5',zeroline:false,title:'',tickfont:{size:12,color:'#3d4652'}},yaxis:{gridcolor:'#edf1f5',zeroline:false,title:'',tickfont:{size:12,color:'#3d4652'}},legend:{orientation:'h',y:-.18,x:0,bgcolor:'rgba(0,0,0,0)',font:{size:12}},hovermode:'x unified'},extra)}
    function buildSupervisorSummary(data){const map={};allSupervisors.forEach(s=>map[s]={supervisor:s,total:0,inprogress:0,completed:0,rejected:0,savings:0,rate:0});data.forEach(d=>{const s=d.supervisor||'(Blank)';if(!map[s])map[s]={supervisor:s,total:0,inprogress:0,completed:0,rejected:0,savings:0,rate:0};map[s].total++;if(d.status==='In Progress')map[s].inprogress++;if(d.status==='Completed')map[s].completed++;if(d.status==='Rejected')map[s].rejected++;map[s].savings+=Number(d.completedSavings)||0});Object.values(map).forEach(r=>r.rate=r.total?100*r.completed/r.total:0);return Object.values(map).filter(r=>r.total>0||r.savings>0)}
    function buildMonthly(data,metric){const months=allMonths.filter(m=>m&&m!=='No Month');const map={};allSupervisors.forEach(s=>{map[s]={};months.forEach(m=>map[s][m]=0)});data.forEach(d=>{const s=d.supervisor||'(Blank)',m=d.eventMonth;if(!map[s]){map[s]={};months.forEach(mm=>map[s][mm]=0)}if(!map[s][m])map[s][m]=0;map[s][m]+=metric==='savings'?(Number(d.completedSavings)||0):1});return {months,map}}
    function renderMonthlyLine(data,targetId,metric){const {months,map}=buildMonthly(data,metric);const labels=months.map(formatMonth);const traces=allSupervisors.map(s=>{const y=months.map(m=>map[s]?.[m]||0);const any=y.some(v=>v>0);if(!any)return null;const max=Math.max(...y,0);let text=y.map(v=>'');if(metric==='savings'){text=y.map(v=>v>=1000?moneyShort(v):'')}else{const lastIdx=y.map((v,i)=>v>0?i:-1).filter(i=>i>=0).pop();if(lastIdx!==undefined)text[lastIdx]=fmtInt.format(y[lastIdx]);}
      return {type:'scatter',mode:metric==='savings'?'lines+markers+text':'lines+markers',name:s,x:labels,y,line:{width:2.8,color:supervisorColors[s],shape:'linear'},marker:{size:7,color:supervisorColors[s],line:{width:1.5,color:'#fff'}},text,textposition:'top center',textfont:{size:12,color:'#2b333e',family:'Inter, Segoe UI, Arial, sans-serif'},hovertemplate:`${s}<br>%{x}<br>${metric==='savings'?'Savings: $%{y:,.0f}':'Cases: %{y:,.0f}'}<extra></extra>`,connectgaps:false};}).filter(Boolean);
      const maxY=Math.max(1,...traces.flatMap(t=>t.y));const layout=baseLayout();layout.margin.r=metric==='total'?92:34;layout.yaxis.gridcolor='#e8edf3';layout.xaxis.gridcolor='#eef2f6';if(metric==='savings'){layout.yaxis.tickprefix='$';layout.yaxis.tickformat='~s';layout.yaxis.range=[0,maxY*1.22];layout.margin.t=24;}else{layout.yaxis.title={text:'Cases',standoff:6};layout.yaxis.range=[0,Math.max(10,Math.ceil(maxY*1.2/10)*10)];layout.yaxis.nticks=7;layout.hovermode='x unified';const annotations=[];traces.forEach(t=>{let idx=t.y.length-1;while(idx>=0 && !t.y[idx])idx--;if(idx>=0){annotations.push({x:t.x[idx],y:t.y[idx],xref:'x',yref:'y',text:fmtInt.format(t.y[idx]),showarrow:false,xanchor:'left',xshift:10,font:{size:12,color:supervisorColors[t.name],family:'Inter, Segoe UI, Arial, sans-serif'}})}});layout.annotations=annotations;}Plotly.react(targetId,traces,layout,{displayModeBar:false,responsive:true})}
    function renderTotalSavingsSide(summary){const rows=summary.filter(r=>r.savings>0).sort((a,b)=>b.savings-a.savings).slice(0,8).reverse();const max=Math.max(1,...rows.map(r=>r.savings));const layout=baseLayout({margin:{l:70,r:82,t:10,b:62},showlegend:false,hovermode:'closest'});layout.xaxis.tickprefix='$';layout.xaxis.tickformat='~s';layout.xaxis.range=[0,max*1.26];layout.yaxis.gridcolor='rgba(0,0,0,0)';Plotly.react('totalSavingsBySupervisorChart',[{type:'bar',orientation:'h',x:rows.map(r=>r.savings),y:rows.map(r=>shortName(r.supervisor)),marker:{color:rows.map(r=>supervisorColors[r.supervisor])},text:rows.map(r=>moneyShort(r.savings)),textposition:'outside',cliponaxis:false,hovertemplate:'%{customdata}<br>Total Savings: $%{x:,.0f}<extra></extra>',customdata:rows.map(r=>r.supervisor)}],layout,{displayModeBar:false,responsive:true})}
    function renderStatus(summary){const rows=summary.sort((a,b)=>b.total-a.total).slice().reverse();const y=rows.map(r=>shortName(r.supervisor));const max=Math.max(1,...rows.map(r=>r.total));function show(v,total){return v>0&&v>=Math.max(3,total*.06)?fmtInt.format(v):''}const traces=[{name:'Completed',key:'completed',color:statusColors.Completed},{name:'In Progress',key:'inprogress',color:statusColors['In Progress']},{name:'Rejected',key:'rejected',color:statusColors.Rejected}].map(s=>({type:'bar',orientation:'h',name:s.name,x:rows.map(r=>r[s.key]),y,marker:{color:s.color},text:rows.map(r=>show(r[s.key],r.total)),textposition:'inside',insidetextanchor:'middle',textfont:{size:12,color:'#fff'},hovertemplate:'%{customdata}<br>'+s.name+': %{x}<extra></extra>',customdata:rows.map(r=>r.supervisor)}));traces.push({type:'scatter',mode:'text',x:rows.map(r=>r.total+max*.035),y,text:rows.map(r=>fmtInt.format(r.total)),textposition:'middle left',textfont:{size:12,color:'#202631'},hoverinfo:'skip',showlegend:false});const layout=baseLayout({barmode:'stack',margin:{l:72,r:64,t:12,b:60}});layout.xaxis.range=[0,max*1.18];layout.legend={orientation:'h',y:1.12,x:.18,bgcolor:'rgba(0,0,0,0)'};Plotly.react('statusBySupervisorChart',traces,layout,{displayModeBar:false,responsive:true})}
    function renderCompletion(summary){const rows=summary.filter(r=>r.total>0).sort((a,b)=>b.rate-a.rate).slice().reverse();const layout=baseLayout({margin:{l:70,r:75,t:12,b:60},showlegend:false,hovermode:'closest'});layout.xaxis.range=[0,100];layout.xaxis.ticksuffix='%';layout.xaxis.dtick=20;layout.yaxis.gridcolor='rgba(0,0,0,0)';Plotly.react('completionRateChart',[{type:'bar',orientation:'h',x:rows.map(r=>r.rate),y:rows.map(r=>shortName(r.supervisor)),marker:{color:rows.map(r=>supervisorColors[r.supervisor])},text:rows.map(r=>`${fmtPct.format(r.rate)}%`),textposition:'outside',cliponaxis:false,customdata:rows.map(r=>[r.completed,r.total,r.supervisor]),hovertemplate:'%{customdata[2]}<br>Completion: %{x:.1f}%<br>Completed: %{customdata[0]} / %{customdata[1]}<extra></extra>'}],layout,{displayModeBar:false,responsive:true})}
    function setMetric(id,value){document.getElementById(id).textContent=value}
    function updateKpis(data,summary){const total=data.length,inprogress=data.filter(d=>d.status==='In Progress').length,completed=data.filter(d=>d.status==='Completed').length,rejected=data.filter(d=>d.status==='Rejected').length,savings=data.reduce((s,d)=>s+(Number(d.completedSavings)||0),0);setMetric('kpiTotal',fmtInt.format(total));setMetric('kpiInProgress',fmtInt.format(inprogress));setMetric('kpiCompleted',fmtInt.format(completed));setMetric('kpiRejected',fmtInt.format(rejected));setMetric('kpiSavings',fmtMoney.format(savings));setMetric('savingsBadge',fmtMoney.format(savings));setMetric('totalBadge',`${fmtInt.format(total)} cases`);setMetric('completionBadge',total?`${fmtPct.format(100*completed/total)}%`:'0%')}
    function updateActiveFilters(){const chips=[];if(state.supervisors.size!==allSupervisors.length)chips.push(`<span class="chip">Supervisor: ${state.supervisors.size} selected</span>`);if(state.statuses.size!==statuses.length)chips.push(`<span class="chip">Status: ${state.statuses.size} selected</span>`);if(state.gkTypes.size!==allGkTypes.length)chips.push(`<span class="chip">GK Type: ${state.gkTypes.size} selected</span>`);if(state.months.size!==allMonths.length)chips.push(`<span class="chip">Month: ${state.months.size} selected</span>`);if(state.weeks.size!==weekOrder.length)chips.push(`<span class="chip">Week: ${state.weeks.size} selected</span>`);if(state.dateFrom||state.dateTo)chips.push(`<span class="chip">Date: ${state.dateFrom||'...'} → ${state.dateTo||'...'}</span>`);document.getElementById('activeFilters').innerHTML=chips.join('')}
    function refresh(){const data=getFilteredData();const summary=buildSupervisorSummary(data);updateSummaries();updateKpis(data,summary);updateActiveFilters();renderMonthlyLine(data,'monthlySavingsChart','savings');renderTotalSavingsSide(summary);renderMonthlyLine(data,'monthlyTotalChart','total');renderStatus(summary);renderCompletion(summary)}
    function syncStateFromInputs(){state.supervisors=new Set([...document.querySelectorAll('input[data-filter="supervisor"]:checked')].map(el=>el.value));state.statuses=new Set([...document.querySelectorAll('input[data-filter="status"]:checked')].map(el=>el.value));state.gkTypes=new Set([...document.querySelectorAll('input[data-filter="gkType"]:checked')].map(el=>el.value));state.months=new Set([...document.querySelectorAll('input[data-filter="month"]:checked')].map(el=>el.value));state.weeks=new Set([...document.querySelectorAll('input[data-filter="week"]:checked')].map(el=>el.value));state.dateFrom=document.getElementById('dateFrom').value;state.dateTo=document.getElementById('dateTo').value;refresh()}
    function toggleFilterSet(filter,checked){document.querySelectorAll(`input[data-filter="${filter}"]`).forEach(el=>{if(el.closest('.check-item')?.style.display==='none')return;el.checked=checked});syncStateFromInputs()}
    function resetFilters(){document.querySelectorAll('input[data-filter]').forEach(el=>el.checked=true);document.querySelectorAll('.filter-search').forEach(input=>input.value='');document.querySelectorAll('.check-item').forEach(item=>item.style.display='flex');document.getElementById('dateFrom').value='';document.getElementById('dateTo').value='';syncStateFromInputs()}
    function closeOpenFilters(except=null){document.querySelectorAll('details.multi[open]').forEach(d=>{if(d!==except)d.open=false})}
    function initDropdownSearch(){document.querySelectorAll('.filter-search').forEach(input=>{input.addEventListener('input',()=>{const q=input.value.trim().toLowerCase();const target=document.getElementById(input.dataset.target);if(!target)return;target.querySelectorAll('.check-item').forEach(item=>{item.style.display=item.textContent.toLowerCase().includes(q)?'flex':'none'})})})}
    document.querySelectorAll('input[data-filter]').forEach(el=>el.addEventListener('change',syncStateFromInputs));document.getElementById('dateFrom').addEventListener('change',syncStateFromInputs);document.getElementById('dateTo').addEventListener('change',syncStateFromInputs);document.getElementById('resetFilters').addEventListener('click',resetFilters);document.querySelectorAll('.tiny-btn[data-action]').forEach(btn=>btn.addEventListener('click',e=>{e.preventDefault();toggleFilterSet(btn.dataset.filter,btn.dataset.action==='all')}));document.querySelectorAll('[data-close-filter="true"]').forEach(btn=>btn.addEventListener('click',e=>{e.preventDefault();const details=btn.closest('details.multi');if(details)details.open=false}));document.querySelectorAll('details.multi').forEach(details=>details.addEventListener('toggle',()=>{if(details.open)closeOpenFilters(details)}));document.addEventListener('click',e=>{if(!e.target.closest('details.multi'))closeOpenFilters()});document.querySelectorAll('a.page-link').forEach(a=>{a.addEventListener('click',e=>{if(e.ctrlKey||e.metaKey||e.shiftKey||e.altKey)return;e.preventDefault();document.body.classList.add('page-leave');setTimeout(()=>{window.location.href=a.href},150)})});initDropdownSearch();resetFilters();
  </script>
</body>
</html>'''

    hero_logo = ''
    if logo_data_uri:
        hero_logo = f'<div class="hero-logo-card"><img src="{logo_data_uri}" alt="Milwaukee logo" class="hero-logo" /></div>'

    replacements = {
        '__PLOTLY_JS__': plotly_js,
        '__RAW_DATA__': json.dumps(data_records, ensure_ascii=False),
        '__STATUS_COLORS__': json.dumps(STATUS_COLORS),
        '__WEEK_ORDER__': json.dumps(weeks),
        '__SUPERVISORS_JSON__': json.dumps(supervisors, ensure_ascii=False),
        '__MONTHS_JSON__': json.dumps(months, ensure_ascii=False),
        '__GK_TYPES_JSON__': json.dumps(gk_types, ensure_ascii=False),
        '__SUPERVISOR_CHECKBOXES__': _checkbox_html('supervisor', supervisors),
        '__STATUS_CHECKBOXES__': _checkbox_html('status', statuses),
        '__GK_TYPE_CHECKBOXES__': _checkbox_html('gkType', gk_types),
        '__MONTH_CHECKBOXES__': _checkbox_html('month', months),
        '__WEEK_CHECKBOXES__': _checkbox_html('week', weeks),
        '__HERO_LOGO__': hero_logo,
    }
    for key, value in replacements.items():
        template = template.replace(key, value)
    return template

def image_path_to_data_uri(path: Path) -> str:
    suffix = path.suffix.lower()
    mime = 'image/png' if suffix == '.png' else 'image/jpeg' if suffix in {'.jpg', '.jpeg'} else 'application/octet-stream'
    return f'data:{mime};base64,' + base64.b64encode(path.read_bytes()).decode('ascii')


def logo_html_to_data_uri(path: Path) -> str:
    text = path.read_text(encoding='utf-8', errors='ignore')
    patterns = [r'url\((data:image/[^)]+)\)', r'src=[\"\'](data:image/[^\"\']+)']
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1).strip().strip('\"\'')
    return ''


def extract_logo_data_uri(logo_html: Path | None = None, logo_image: Path | None = None, logo_pptx: Path | None = None) -> str:
    if logo_html and logo_html.exists():
        data_uri = logo_html_to_data_uri(logo_html)
        if data_uri:
            return data_uri
    if logo_image and logo_image.exists():
        return image_path_to_data_uri(logo_image)
    if logo_pptx and logo_pptx.exists():
        with zipfile.ZipFile(logo_pptx, 'r') as zf:
            candidates = [n for n in zf.namelist() if n.startswith('ppt/media/') and n.lower().endswith(('.png', '.jpg', '.jpeg'))]
            preferred = None
            for pref in ['image3.png', 'image2.png', 'image1.png']:
                for name in candidates:
                    if name.lower().endswith(pref):
                        preferred = name
                        break
                if preferred:
                    break
            if preferred is None and candidates:
                preferred = candidates[0]
            if preferred:
                data = zf.read(preferred)
                lower = preferred.lower()
                mime = 'image/png' if lower.endswith('.png') else 'image/jpeg'
                return f'data:{mime};base64,' + base64.b64encode(data).decode('ascii')
    return ''


def main() -> None:
    parser = argparse.ArgumentParser(description='Build static GK dashboard HTML.')
    parser.add_argument('--input', type=Path, default=DEFAULT_INPUT_FILE)
    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT_FILE)
    parser.add_argument('--sheet-name', default=DEFAULT_SHEET_NAME)
    parser.add_argument('--logo-html', type=Path, default=DEFAULT_LOGO_HTML)
    parser.add_argument('--logo-pptx', type=Path, default=DEFAULT_LOGO_PPTX)
    parser.add_argument('--drop-unmapped-supervisor', action='store_true', help='Drop rows with blank Supervisor Display before rendering.')
    args = parser.parse_args()
    sheet_name = int(args.sheet_name) if str(args.sheet_name).isdigit() else args.sheet_name
    df = load_data(args.input, sheet_name=sheet_name)
    if args.drop_unmapped_supervisor and 'Supervisor Display' in df.columns:
        df = df[df['Supervisor Display'].astype(str).str.strip() != ''].copy()
    records = serializable_records(df)
    latest_dates = pd.concat([df['Submitted Date Parsed'], df['Completed Date Parsed']], ignore_index=True).dropna()
    latest_update = format_display_date(latest_dates.max()) if not latest_dates.empty else ''
    html_text = render_dashboard(records, args.input.name, logo_data_uri=extract_logo_data_uri(logo_html=args.logo_html, logo_pptx=args.logo_pptx), latest_update_text=latest_update)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(html_text, encoding='utf-8')
    print(f'Dashboard written to: {args.output}')


if __name__ == '__main__':
    main()
