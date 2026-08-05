# GK_DASHBOARD_CORE_VERSION = 'v9_static_read_only_2026_08_05'
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
DEFAULT_OUTPUT_FILE = Path('index.html')
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
    # Prefer an explicit approval date when supplied; current exports use Completed Date.
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

    # Timeline rule:
    # - Completed cases are recorded on the approval date.
    # - Other cases remain recorded on the submitted date.
    # - Approved Date / Approval Date is preferred; current exports fall back to Completed Date.
    # - If none of those dates exists, fall back to Submitted Date.
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
    """Return only the fields required by the read-only charts.

    Keeping the static payload minimal avoids embedding descriptions, comments,
    employee IDs, and other source columns that are not used by the dashboard.
    """
    keep_cols = [
        'Supervisor Display', 'Approval Status Clean', 'GK Type Filter',
        'Completed Savings USD', 'Event Date Text', 'Event Week', 'Event Month',
    ]
    work = df.copy()
    for col in keep_cols:
        if col not in work.columns:
            work[col] = ''

    records = []
    for _, row in work[keep_cols].iterrows():
        supervisor = row['Supervisor Display'] if str(row['Supervisor Display']).strip() else '(Blank)'
        records.append({
            'supervisor': supervisor,
            'status': row['Approval Status Clean'],
            'gkTypeFilter': row['GK Type Filter'] if str(row['GK Type Filter']).strip() else '(Blank)',
            'completedSavings': round(float(row['Completed Savings USD'] or 0), 2),
            'eventDate': row['Event Date Text'],
            'eventWeek': row['Event Week'],
            'eventMonth': row['Event Month'],
        })
    return records


def _checkbox_html(filter_name, values):
    items = []
    for val in values:
        safe = html.escape(str(val), quote=True)
        items.append(f'<label class="check-item"><input type="checkbox" data-filter="{filter_name}" value="{safe}" checked> <span>{safe}</span></label>')
    return '\n'.join(items)


def render_dashboard(data_records, source_name: str, logo_data_uri: str = '', latest_update_text: str = '') -> str:
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
  <title>GK Performance Dashboard</title>
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
          __LATEST_UPDATE_LINE__
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
    document.querySelectorAll('input[data-filter]').forEach(el=>el.addEventListener('change',syncStateFromInputs));document.getElementById('dateFrom').addEventListener('change',syncStateFromInputs);document.getElementById('dateTo').addEventListener('change',syncStateFromInputs);document.getElementById('resetFilters').addEventListener('click',resetFilters);document.querySelectorAll('.tiny-btn[data-action]').forEach(btn=>btn.addEventListener('click',e=>{e.preventDefault();toggleFilterSet(btn.dataset.filter,btn.dataset.action==='all')}));document.querySelectorAll('[data-close-filter="true"]').forEach(btn=>btn.addEventListener('click',e=>{e.preventDefault();const details=btn.closest('details.multi');if(details)details.open=false}));document.querySelectorAll('details.multi').forEach(details=>details.addEventListener('toggle',()=>{if(details.open)closeOpenFilters(details)}));document.addEventListener('click',e=>{if(!e.target.closest('details.multi'))closeOpenFilters()});initDropdownSearch();resetFilters();
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
        '__LATEST_UPDATE_LINE__': f'<div style="margin-top:10px;font-size:13px;color:rgba(255,255,255,.62);font-weight:700">Latest data date: <span style="color:#fff">{html.escape(latest_update_text)}</span></div>' if latest_update_text else '',
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
    parser = argparse.ArgumentParser(description='Build a self-contained static GK dashboard for GitHub Pages.')
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
    latest_dates = df['Event Date'].dropna()
    latest_update = format_display_date(latest_dates.max()) if not latest_dates.empty else ''
    html_text = render_dashboard(
        records,
        args.input.name,
        logo_data_uri=extract_logo_data_uri(logo_html=args.logo_html, logo_pptx=args.logo_pptx),
        latest_update_text=latest_update,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(html_text, encoding='utf-8')
    print(f'Dashboard written to: {args.output}')
    print(f'Rows embedded: {len(records)}')


if __name__ == '__main__':
    main()
