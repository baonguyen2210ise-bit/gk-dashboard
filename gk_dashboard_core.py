import argparse
import base64
import html
import json
import re
import unicodedata
import zipfile
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable

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

    df['Submitter Raw'] = df['Submitter'].fillna('').astype(str)
    df['GK Owner Raw'] = df['GK Owner Name'].fillna('').astype(str)
    df['GK Type Display'] = df['GK Type'].fillna('').astype(str).str.strip()
    df['GK Type Filter'] = df['GK Type Display'].where(df['GK Type Display'].str.strip().ne(''), '(Blank)')
    df['GK Other Display'] = df['GK Other'].fillna('').astype(str).str.strip()

    mapped_owner = df['Supervisor Matched Name'].fillna('').astype(str)
    df['Owner Display'] = mapped_owner.apply(lambda x: pretty_name(x, blank_value=''))
    fallback_owner = df['GK Owner Name'].apply(lambda x: pretty_name(x, blank_value=''))
    fallback_submitter = df['Submitter'].apply(lambda x: pretty_name(x, blank_value=''))
    df.loc[df['Owner Display'].astype(str).str.strip().eq(''), 'Owner Display'] = fallback_owner
    df.loc[df['Owner Display'].astype(str).str.strip().eq(''), 'Owner Display'] = fallback_submitter

    df['Submitter Display'] = df['Submitter Raw']
    df['Supervisor Display'] = df['Supervisor'].apply(lambda x: pretty_name(x, blank_value=''))

    df['Event Date'] = pd.to_datetime(df['Submitted Date Parsed'], errors='coerce')
    completed_with_date = df['Approval Status Clean'].eq('Completed') & df['Approval Date Parsed'].notna()
    df.loc[completed_with_date, 'Event Date'] = df.loc[completed_with_date, 'Approval Date Parsed']
    df['Event Week'] = df['Event Date'].apply(date_to_week_label)
    df['Event Month'] = df['Event Date'].apply(date_to_month_label)

    df['Completed Savings USD'] = np.where(
        df['Approval Status Clean'].eq('Completed'),
        df['Cost Reduction Numeric'],
        0.0,
    )

    df['Submitted Date Text'] = df['Submitted Date Parsed'].dt.strftime('%Y-%m-%d').fillna('')
    df['Completed Date Text'] = df['Completed Date Parsed'].dt.strftime('%Y-%m-%d').fillna('')
    df['Approval Date Text'] = df['Approval Date Parsed'].dt.strftime('%Y-%m-%d').fillna('')
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
        'Reference #', 'Improvement Title', 'Submitter Display', 'Owner Display', 'Supervisor Display',
        'Submitter Department', 'Approval Status Clean', 'GK Type Filter', 'GK Other Display',
        'Completed Savings USD', 'Event Date Text', 'EventWeek', 'EventMonth',
        'Event Week', 'Event Month', 'Submitted Date Text', 'Completed Date Text', 'Approval Date Text'
    ]
    work = df.copy()
    for col in keep_cols:
        if col not in work.columns:
            work[col] = ''

    records = []
    for _, row in work.iterrows():
        supervisor = row['Supervisor Display'] if str(row['Supervisor Display']).strip() else '(Blank)'
        owner = row['Owner Display'] if str(row['Owner Display']).strip() else '(Blank)'
        submitter = row['Submitter Display'] if str(row['Submitter Display']).strip() else '(Blank)'
        department = row['Submitter Department'] if str(row['Submitter Department']).strip() else '(Blank)'
        gk_type = row['GK Type Filter'] if str(row['GK Type Filter']).strip() else '(Blank)'
        title = str(row['Improvement Title']).strip()
        reference = normalize_reference(row['Reference #'])
        records.append({
            'reference': reference,
            'title': title,
            'submitter': submitter,
            'owner': owner,
            'supervisor': supervisor,
            'department': department,
            'status': row['Approval Status Clean'],
            'gkTypeFilter': gk_type,
            'gkOther': row['GK Other Display'] if str(row['GK Other Display']).strip() else '',
            'completedSavings': round(float(row['Completed Savings USD'] or 0), 2),
            'eventDate': row['Event Date Text'],
            'eventWeek': row['Event Week'],
            'eventMonth': row['Event Month'],
            'submittedDate': row['Submitted Date Text'],
            'completedDate': row['Completed Date Text'],
            'approvalDate': row['Approval Date Text'],
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
    owners = sorted({(r.get('owner') or '').strip() for r in data_records if str(r.get('owner') or '').strip()}) or ['(Blank)']
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
<title>GK Dashboard</title>
<style>
:root{
  --bg:#edf1f5;--surface:#ffffff;--surface-2:#f8fafc;--line:#dbe3ea;--line-2:#e7edf2;--text:#0f1721;--muted:#647082;
  --accent:#c41230;--accent-dark:#8f1025;--nav:#0d131c;--nav-2:#131c27;--nav-line:rgba(255,255,255,.12);
  --green:#2e8b57;--amber:#f5a623;--red:#c41230;--blue:#2563eb;--shadow:0 18px 40px rgba(15,23,33,.08);--radius:20px;
  --sidebar-w:280px;--sidebar-collapsed:88px;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0;font-family:Inter,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text)}
body{overflow-x:hidden}
.app-shell{display:flex;min-height:100vh}
.sidebar{position:fixed;inset:0 auto 0 0;width:var(--sidebar-w);background:linear-gradient(180deg,var(--nav),#101722 100%);color:#fff;padding:18px 16px 18px;display:flex;flex-direction:column;gap:18px;z-index:50;border-right:1px solid rgba(255,255,255,.05);transition:width .22s ease,transform .22s ease}
body.sidebar-collapsed .sidebar{width:var(--sidebar-collapsed)}
.sidebar-top{display:flex;align-items:center;gap:12px;min-height:44px}
.logo-dot{width:42px;height:42px;border-radius:14px;background:linear-gradient(180deg,#e11d48,#b60f2d);display:flex;align-items:center;justify-content:center;font-weight:900;box-shadow:0 12px 24px rgba(196,18,48,.35)}
.brand-block{min-width:0;transition:opacity .18s ease,transform .18s ease}
.brand-title{font-size:16px;font-weight:900;letter-spacing:.02em}.brand-sub{font-size:12px;color:rgba(255,255,255,.6);margin-top:2px}
body.sidebar-collapsed .brand-block, body.sidebar-collapsed .menu-section-title, body.sidebar-collapsed .nav-label, body.sidebar-collapsed .sidebar-footer{opacity:0;pointer-events:none;transform:translateX(-4px);width:0;overflow:hidden}
.sidebar-toggle{margin-left:auto;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.06);color:#fff;width:38px;height:38px;border-radius:12px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:18px}
.menu-section-title{font-size:11px;text-transform:uppercase;letter-spacing:.16em;color:rgba(255,255,255,.46);padding:0 10px}
.nav-menu{display:grid;gap:8px}
.nav-item{display:flex;align-items:center;gap:12px;border:1px solid transparent;background:transparent;color:#fff;padding:12px 13px;border-radius:16px;cursor:pointer;text-align:left;transition:background .18s ease,border-color .18s ease,transform .18s ease}
.nav-item:hover{background:rgba(255,255,255,.07);transform:translateY(-1px)}
.nav-item.active{background:linear-gradient(180deg,rgba(196,18,48,.22),rgba(196,18,48,.12));border-color:rgba(255,255,255,.12);box-shadow:inset 0 0 0 1px rgba(255,255,255,.05)}
.nav-icon{flex:0 0 auto;width:42px;height:42px;border-radius:14px;background:rgba(255,255,255,.08);display:flex;align-items:center;justify-content:center;font-size:18px}
.nav-label{min-width:0}.nav-title{font-weight:850;font-size:15px}.nav-sub{font-size:12px;color:rgba(255,255,255,.6);margin-top:2px}
.sidebar-footer{margin-top:auto;padding:14px;border:1px solid var(--nav-line);border-radius:18px;background:rgba(255,255,255,.04);font-size:12px;color:rgba(255,255,255,.68);line-height:1.45}
body.sidebar-collapsed .nav-item{justify-content:center;padding:12px 0}body.sidebar-collapsed .sidebar-toggle{margin-left:0}
.main{flex:1;min-width:0;margin-left:var(--sidebar-w);transition:margin-left .22s ease}
body.sidebar-collapsed .main{margin-left:var(--sidebar-collapsed)}
.topbar{position:sticky;top:0;z-index:35;display:flex;align-items:center;justify-content:space-between;gap:16px;padding:16px 26px;background:rgba(237,241,245,.88);backdrop-filter:blur(12px);border-bottom:1px solid rgba(15,23,33,.06)}
.topbar-title{font-size:14px;color:var(--muted);font-weight:700}.topbar-strong{display:block;font-size:24px;color:var(--text);font-weight:950;letter-spacing:-.03em;margin-top:4px}
.topbar-actions{display:flex;align-items:center;gap:10px;flex-wrap:wrap}.pill{display:inline-flex;align-items:center;gap:8px;padding:10px 14px;border-radius:999px;background:#fff;border:1px solid var(--line);box-shadow:0 10px 22px rgba(15,23,33,.05);font-size:13px;font-weight:800;color:var(--muted)}
.page{padding:0 26px 28px}
.hero{margin-top:10px;position:relative;overflow:hidden;border-radius:28px;background:radial-gradient(circle at 85% 0,rgba(196,18,48,.36),transparent 30%),linear-gradient(115deg,#070c12 0%,#101824 44%,#5f1025 100%);color:#fff;padding:28px 30px 26px;border:1px solid rgba(255,255,255,.06);box-shadow:var(--shadow)}
.hero:before{content:'';position:absolute;inset:auto 0 0 0;height:4px;background:linear-gradient(90deg,#ff5f79,#c41230)}
.hero-grid{position:absolute;right:14px;bottom:12px;width:300px;height:130px;opacity:.12;background-image:radial-gradient(rgba(255,255,255,.55) 1px,transparent 1px);background-size:10px 10px;mask-image:linear-gradient(90deg,transparent,#000 24%,#000)}
.hero-content{position:relative;z-index:1;display:flex;justify-content:space-between;gap:24px;align-items:flex-start}
.hero-kicker{font-size:11px;text-transform:uppercase;letter-spacing:.18em;color:rgba(255,255,255,.65);font-weight:900;margin-bottom:10px}
.hero h1{margin:0;font-size:44px;line-height:1.05;letter-spacing:-.045em;max-width:840px}
.hero-sub{margin-top:10px;color:rgba(255,255,255,.72);font-size:15px;max-width:760px;line-height:1.55}
.hero-meta{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}
.meta-chip{padding:10px 14px;border-radius:999px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.08);font-size:13px;font-weight:850;color:#fff}
.logo-card{flex:0 0 auto;width:250px;max-width:100%;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:20px;padding:12px 18px;display:flex;align-items:center;justify-content:center;min-height:110px;backdrop-filter:blur(6px)}
.logo-card img{display:block;max-width:190px;max-height:68px;width:100%;height:auto;object-fit:contain}
.filters-panel{margin-top:18px;background:rgba(255,255,255,.86);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.9);border-radius:24px;padding:18px;box-shadow:var(--shadow)}
.filter-head{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap}.filter-title{font-size:28px;font-weight:950;letter-spacing:-.03em}.filter-actions{display:flex;gap:10px;flex-wrap:wrap}
.btn{border:0;border-radius:14px;padding:11px 15px;font-weight:900;cursor:pointer;font-size:14px;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:transform .18s ease, box-shadow .18s ease, background .18s ease}
.btn:hover{transform:translateY(-1px);box-shadow:0 10px 20px rgba(15,23,33,.12)}
.btn-primary{background:linear-gradient(180deg,#d91539,#b0102d);color:#fff}.btn-secondary{background:#fff;color:var(--text);border:1px solid var(--line)}.btn-ghost{background:#f5f7fa;color:#334155;border:1px solid #e4e9ef}
.filter-grid{margin-top:16px;display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:12px}
.filter-grid.detail{grid-template-columns:repeat(7,minmax(0,1fr))}
.filter-block{background:linear-gradient(180deg,#fff,#fbfdff);border:1px solid var(--line);border-radius:18px;padding:12px;min-height:84px}
.filter-block label.title{display:block;font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#687588;font-weight:950;margin-bottom:8px}
.filter-block input[type="date"], .filter-search, .table-search{width:100%;border:1px solid var(--line);background:#fff;border-radius:12px;padding:10px 11px;font-size:13px;outline:none}
.date-range{display:grid;grid-template-columns:1fr 1fr;gap:8px}
details.multi{position:relative}details.multi summary{list-style:none;cursor:pointer;border:1px solid var(--line);background:#fff;border-radius:13px;padding:10px 12px;display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:14px;color:#0f1721}details.multi summary::-webkit-details-marker{display:none}.summary-count{color:var(--muted);font-weight:900}
.option-panel{position:absolute;left:0;right:0;top:calc(100% + 8px);z-index:90;background:#fff;border:1px solid var(--line);border-radius:18px;box-shadow:0 24px 48px rgba(15,23,33,.16);padding:0;max-height:330px;overflow:auto}.option-tools{display:grid;grid-template-columns:1fr auto auto auto;gap:8px;align-items:center;background:#fff;padding:10px 10px 9px;border-bottom:1px solid #edf0f2;position:sticky;top:0;z-index:3}.option-tools.no-search{grid-template-columns:auto auto auto;justify-content:start}.tiny-btn{border:1px solid var(--line);background:#fff;border-radius:10px;padding:7px 9px;font-size:12px;cursor:pointer;white-space:nowrap}.check-list{display:grid;gap:5px;padding:10px}.check-item{display:flex;gap:9px;align-items:center;padding:6px 5px;border-radius:10px;font-size:13px}.check-item:hover{background:#f5f7fb}
.active-filters{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}.chip{padding:8px 10px;border-radius:999px;background:#fff2f5;color:var(--accent);border:1px solid #ffd5de;font-size:12px;font-weight:900}
.kpi-grid{margin-top:16px;display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:14px}
.kpi-card{background:var(--surface);border:1px solid var(--line);border-radius:22px;padding:18px;box-shadow:var(--shadow);display:flex;align-items:center;gap:16px;min-height:100px;position:relative;overflow:hidden}.kpi-card.primary{background:linear-gradient(135deg,#0c131c 0%,#121c28 100%);color:#fff;border-color:#172231}.kpi-card.primary:before{content:'';position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,#ff5f79,#c41230)}
.kpi-icon{width:52px;height:52px;border-radius:18px;background:#fff;border:2px solid rgba(196,18,48,.55);color:var(--accent);display:flex;align-items:center;justify-content:center;font-size:28px;font-weight:900;flex:0 0 auto}.kpi-card.primary .kpi-icon{background:rgba(255,255,255,.08);border-color:rgba(255,255,255,.24);color:#ff5f79}.kpi-icon.green{border-color:rgba(46,139,87,.5);color:var(--green)}.kpi-icon.amber{border-color:rgba(245,166,35,.58);color:var(--amber)}.kpi-icon.red{border-color:rgba(196,18,48,.58);color:var(--red)}
.kpi-label{font-size:12px;text-transform:uppercase;letter-spacing:.09em;color:#6c788a;font-weight:950}.kpi-card.primary .kpi-label{color:rgba(255,255,255,.62)}.kpi-value{font-size:42px;font-weight:950;letter-spacing:-.05em;line-height:1;margin-top:6px}
.view{display:none}.view.active{display:block;animation:fadeIn .22s ease both}@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
.chart-grid-2{margin-top:16px;display:grid;grid-template-columns:minmax(0,1.55fr) minmax(0,1fr);gap:16px}.chart-grid-3{margin-top:16px;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px}.chart-grid-main{margin-top:16px;display:grid;grid-template-columns:minmax(0,1.8fr) minmax(320px,1fr);gap:16px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:24px;padding:18px;box-shadow:var(--shadow);min-width:0}.card-head{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;margin-bottom:10px}.card-title{font-size:20px;font-weight:950;letter-spacing:-.03em}.card-sub{color:var(--muted);font-size:13px;margin-top:4px}.badge{padding:10px 14px;border-radius:999px;background:#f4f7fb;color:#243143;font-size:13px;font-weight:900;white-space:nowrap}
.chart{width:100%;height:380px}.chart.small{height:320px}.chart.tall{height:430px}
.table-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:12px}.table-tools{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:18px;background:#fff}.data-table{width:100%;border-collapse:collapse;min-width:1100px}.data-table th,.data-table td{padding:12px 14px;border-bottom:1px solid #edf1f5;text-align:left;font-size:13px;vertical-align:top}.data-table th{position:sticky;top:0;background:#f8fafc;font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#5f6c7d;z-index:1}.data-table tr:hover td{background:#fafcff}.table-status{display:inline-flex;align-items:center;padding:6px 9px;border-radius:999px;font-size:12px;font-weight:900}.table-status.Completed{background:#eaf7ef;color:#17653c}.table-status.Rejected{background:#fff0f2;color:#9f102a}.table-status.InProgress{background:#fff7e8;color:#a15f00}
.table-empty{padding:22px;color:var(--muted);font-weight:700}
@media (max-width: 1380px){.filter-grid,.filter-grid.detail,.kpi-grid,.chart-grid-3{grid-template-columns:repeat(2,minmax(0,1fr))}.chart-grid-2,.chart-grid-main{grid-template-columns:1fr}.hero-content{flex-direction:column}.logo-card{width:220px}}
@media (max-width: 960px){.sidebar{transform:translateX(-100%)}body.mobile-sidebar-open .sidebar{transform:translateX(0);width:var(--sidebar-w)}.main,body.sidebar-collapsed .main{margin-left:0}.topbar{padding:14px 18px}.page{padding:0 18px 24px}.hero h1{font-size:34px}.filter-grid,.filter-grid.detail,.kpi-grid,.chart-grid-3{grid-template-columns:1fr}.date-range{grid-template-columns:1fr}.topbar-strong{font-size:20px}}
</style>
</head>
<body>
<div class="app-shell">
  <aside class="sidebar" id="sidebar">
    <div class="sidebar-top">
      <div class="logo-dot">GK</div>
      <div class="brand-block">
        <div class="brand-title">GK Dashboard</div>
        <div class="brand-sub">Static analytics workspace</div>
      </div>
      <button class="sidebar-toggle" id="sidebarToggle" title="Collapse / expand navigation">☰</button>
    </div>

    <div class="menu-section-title">Navigation</div>
    <nav class="nav-menu">
      <button class="nav-item active" data-view-btn="supervisor">
        <div class="nav-icon">📊</div>
        <div class="nav-label">
          <div class="nav-title">Official Dashboard</div>
          <div class="nav-sub">Supervisor overview</div>
        </div>
      </button>
      <button class="nav-item" data-view-btn="detail">
        <div class="nav-icon">📋</div>
        <div class="nav-label">
          <div class="nav-title">Detail Dashboard</div>
          <div class="nav-sub">Detail analytics & records</div>
        </div>
      </button>
    </nav>

    <div class="sidebar-footer">
      <div><strong>Latest data:</strong> __LATEST_UPDATE_TEXT__</div>
      <div style="margin-top:6px"><strong>Source:</strong> __SOURCE_NAME__</div>
      <div style="margin-top:10px">Use the navigation on the left to switch between the supervisor dashboard and the detail dashboard.</div>
    </div>
  </aside>

  <main class="main">
    <div class="topbar">
      <div>
        <div class="topbar-title">Gemba Kaizen Dashboard</div>
        <span class="topbar-strong" id="topbarTitle">Official Dashboard · Supervisor overview</span>
      </div>
      <div class="topbar-actions">
        <div class="pill">Updated: __LATEST_UPDATE_TEXT__</div>
        <button class="btn btn-secondary" id="mobileNavToggle">☰ Menu</button>
      </div>
    </div>

    <div class="page">
      <section class="hero">
        <div class="hero-grid"></div>
        <div class="hero-content">
          <div>
            <div class="hero-kicker">Gemba Kaizen Performance</div>
            <h1 id="heroTitle">GK Supervisor Performance Overview</h1>
            <div class="hero-sub" id="heroSubtitle">A clean read-only workspace with one navigation panel and two dashboard views: the supervisor overview and the detailed analytics dashboard.</div>
            <div class="hero-meta">
              <span class="meta-chip">Official dashboard</span>
              <span class="meta-chip">Detail dashboard</span>
              <span class="meta-chip">Completed cases use approval / completed date</span>
            </div>
          </div>
          <div class="logo-card">__HERO_LOGO__</div>
        </div>
      </section>

      <section class="filters-panel">
        <div class="filter-head">
          <div>
            <div class="filter-title">Filters</div>
            <div class="card-sub">Shared filters apply to both dashboards. The detail dashboard has extra tools below for table search and CSV export.</div>
          </div>
          <div class="filter-actions">
            <button class="btn btn-ghost" id="toggleSidebarBtn">⇤ Toggle navigation</button>
            <button class="btn btn-primary" id="resetFilters">↻ Reset filters</button>
          </div>
        </div>

        <div class="filter-grid detail">
          <div class="filter-block"><label class="title">Supervisor</label><details class="multi"><summary><span>Select supervisor(s)</span><span class="summary-count" id="supervisorSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search supervisor..." class="filter-search" data-target="supervisorOptions"><button class="tiny-btn" data-action="all" data-filter="supervisor">All</button><button class="tiny-btn" data-action="none" data-filter="supervisor">None</button><button class="tiny-btn" type="button" data-close-filter="true">Close</button></div><div class="check-list" id="supervisorOptions">__SUPERVISOR_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">GK Owner</label><details class="multi"><summary><span>Select owner(s)</span><span class="summary-count" id="ownerSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search owner..." class="filter-search" data-target="ownerOptions"><button class="tiny-btn" data-action="all" data-filter="owner">All</button><button class="tiny-btn" data-action="none" data-filter="owner">None</button><button class="tiny-btn" type="button" data-close-filter="true">Close</button></div><div class="check-list" id="ownerOptions">__OWNER_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Status</label><details class="multi"><summary><span>Select status</span><span class="summary-count" id="statusSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="status">All</button><button class="tiny-btn" data-action="none" data-filter="status">None</button><button class="tiny-btn" type="button" data-close-filter="true">Close</button></div><div class="check-list">__STATUS_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">GK Type</label><details class="multi"><summary><span>Select type</span><span class="summary-count" id="gkTypeSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search type..." class="filter-search" data-target="gkTypeOptions"><button class="tiny-btn" data-action="all" data-filter="gkType">All</button><button class="tiny-btn" data-action="none" data-filter="gkType">None</button><button class="tiny-btn" type="button" data-close-filter="true">Close</button></div><div class="check-list" id="gkTypeOptions">__GK_TYPE_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Month</label><details class="multi"><summary><span>Select month</span><span class="summary-count" id="monthSummary">All</span></summary><div class="option-panel"><div class="option-tools no-search"><button class="tiny-btn" data-action="all" data-filter="month">All</button><button class="tiny-btn" data-action="none" data-filter="month">None</button><button class="tiny-btn" type="button" data-close-filter="true">Close</button></div><div class="check-list" id="monthOptions">__MONTH_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Week</label><details class="multi"><summary><span>Select week(s)</span><span class="summary-count" id="weekSummary">All</span></summary><div class="option-panel"><div class="option-tools"><input type="text" placeholder="Search week..." class="filter-search" data-target="weekOptions"><button class="tiny-btn" data-action="all" data-filter="week">All</button><button class="tiny-btn" data-action="none" data-filter="week">None</button><button class="tiny-btn" type="button" data-close-filter="true">Close</button></div><div class="check-list" id="weekOptions">__WEEK_CHECKBOXES__</div></div></details></div>
          <div class="filter-block"><label class="title">Date range</label><div class="date-range"><input type="date" id="dateFrom"><input type="date" id="dateTo"></div></div>
        </div>
        <div class="active-filters" id="activeFilters"></div>
      </section>

      <section class="kpi-grid">
        <div class="kpi-card primary"><div class="kpi-icon">✓</div><div><div class="kpi-label">Total GK</div><div class="kpi-value" id="kpiTotal">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon amber">◌</div><div><div class="kpi-label">In Progress</div><div class="kpi-value" id="kpiInProgress">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon green">✓</div><div><div class="kpi-label">Completed</div><div class="kpi-value" id="kpiCompleted">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon red">×</div><div><div class="kpi-label">Rejected</div><div class="kpi-value" id="kpiRejected">0</div></div></div>
        <div class="kpi-card"><div class="kpi-icon">$</div><div><div class="kpi-label">Completed Savings</div><div class="kpi-value" id="kpiSavings">$0</div></div></div>
      </section>

      <section class="view active" id="view-supervisor">
        <div class="chart-grid-main">
          <div class="card"><div class="card-head"><div><div class="card-title">Monthly GK Activity</div><div class="card-sub">Monthly case trend based on event date. Completed cases follow approval / completed date.</div></div><div class="badge" id="superMonthlyBadge">0 cases</div></div><div id="supMonthlyChart" class="chart"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Status Mix</div><div class="card-sub">Share of the currently filtered cases</div></div><div class="badge" id="statusBadge">0%</div></div><div id="supStatusChart" class="chart small"></div></div>
        </div>
        <div class="chart-grid-2">
          <div class="card"><div class="card-head"><div><div class="card-title">Completed Savings by Supervisor</div><div class="card-sub">Top supervisors ranked by completed savings</div></div><div class="badge" id="savingsBadge">$0</div></div><div id="supSavingsChart" class="chart"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Completion Rate by Supervisor</div><div class="card-sub">Completed ÷ total cases</div></div><div class="badge" id="completionBadge">0%</div></div><div id="supCompletionChart" class="chart"></div></div>
        </div>
        <div class="chart-grid-2">
          <div class="card"><div class="card-head"><div><div class="card-title">Total Cases by Supervisor</div><div class="card-sub">Top supervisors ranked by total volume</div></div><div class="badge" id="volumeBadge">0 cases</div></div><div id="supCasesChart" class="chart"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">GK Type Mix</div><div class="card-sub">Distribution of GK types after filtering</div></div><div class="badge" id="typeBadge">0 types</div></div><div id="supTypeChart" class="chart"></div></div>
        </div>
      </section>

      <section class="view" id="view-detail">
        <div class="chart-grid-main">
          <div class="card"><div class="card-head"><div><div class="card-title">Weekly GK Activity</div><div class="card-sub">Event-week activity by status</div></div><div class="badge" id="detailWeeklyBadge">0 weeks</div></div><div id="detailWeeklyChart" class="chart"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Status Mix</div><div class="card-sub">Detail dashboard status split</div></div><div class="badge" id="detailStatusBadge">0%</div></div><div id="detailStatusChart" class="chart small"></div></div>
        </div>
        <div class="chart-grid-3">
          <div class="card"><div class="card-head"><div><div class="card-title">Top Owners by Savings</div><div class="card-sub">Highest completed savings</div></div></div><div id="detailOwnerSavingsChart" class="chart small"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Top Owners by Cases</div><div class="card-sub">Highest case volume</div></div></div><div id="detailOwnerCasesChart" class="chart small"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Top Departments</div><div class="card-sub">Most active submitter departments</div></div></div><div id="detailDepartmentChart" class="chart small"></div></div>
        </div>
        <div class="card" style="margin-top:16px">
          <div class="table-head">
            <div>
              <div class="card-title">Detail Records</div>
              <div class="card-sub">Search, review, and export the filtered records.</div>
            </div>
            <div class="table-tools">
              <input class="table-search" id="tableSearch" placeholder="Search reference, title, submitter, owner..." style="width:320px;max-width:100%">
              <button class="btn btn-secondary" id="exportCsvBtn">Export filtered CSV</button>
            </div>
          </div>
          <div class="table-wrap" id="detailTableWrap"></div>
        </div>
      </section>
    </div>
  </main>
</div>
<script>__PLOTLY_JS__</script>
<script>
const rawData = __RAW_DATA__;
const statusColors = __STATUS_COLORS__;
const weekOrder = __WEEK_ORDER__;
const allSupervisors = __SUPERVISORS_JSON__;
const allOwners = __OWNERS_JSON__;
const allMonths = __MONTHS_JSON__;
const allGkTypes = __GK_TYPES_JSON__;
const statuses = ['Completed','In Progress','Rejected'];
const fmtInt = new Intl.NumberFormat('en-US');
const fmtMoney = new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0});
const fmtPct = new Intl.NumberFormat('en-US',{maximumFractionDigits:1});
const state = {
  activeView:'supervisor',
  supervisors:new Set(allSupervisors),
  owners:new Set(allOwners),
  statuses:new Set(statuses),
  gkTypes:new Set(allGkTypes),
  months:new Set(allMonths),
  weeks:new Set(weekOrder),
  dateFrom:'',
  dateTo:'',
  tableSearch:''
};

function getSummaryText(selectedCount,totalCount){return selectedCount===totalCount?'All':String(selectedCount)}
function setMetric(id,v){const el=document.getElementById(id);if(el)el.textContent=v}
function safeText(v){return (v ?? '').toString()}
function statusClass(v){return v==='Completed'?'Completed':(v==='Rejected'?'Rejected':'InProgress')}
function datePasses(eventDate){if(!state.dateFrom && !state.dateTo) return true; if(!eventDate) return false; if(state.dateFrom && eventDate < state.dateFrom) return false; if(state.dateTo && eventDate > state.dateTo) return false; return true;}
function getFilteredData(){return rawData.filter(r=> state.supervisors.has(r.supervisor) && state.owners.has(r.owner) && state.statuses.has(r.status) && state.gkTypes.has(r.gkTypeFilter) && state.months.has(r.eventMonth) && state.weeks.has(r.eventWeek) && datePasses(r.eventDate));}
function getTableData(data){const q=state.tableSearch.trim().toLowerCase(); if(!q) return data; return data.filter(r=>[r.reference,r.title,r.submitter,r.owner,r.supervisor,r.department,r.status,r.gkTypeFilter].join(' ').toLowerCase().includes(q));}
function aggregateByField(data,field){const m=new Map();data.forEach(r=>{const k=(r[field]||'(Blank)').toString().trim()||'(Blank)'; if(!m.has(k)) m.set(k,[]); m.get(k).push(r);}); return m;}
function topGroups(data,field,metric='count',limit=10){const groups=aggregateByField(data,field); const rows=[]; for(const [name,items] of groups.entries()){ const total=items.length; const completed=items.filter(d=>d.status==='Completed').length; const savings=items.reduce((s,d)=>s+(Number(d.completedSavings)||0),0); rows.push({name,total,completed,savings,completionRate:total?completed/total:0}); } rows.sort((a,b)=> (metric==='savings' ? b.savings-a.savings : metric==='completionRate' ? b.completionRate-a.completionRate : b.total-a.total)); return rows.slice(0,limit); }
function monthBuckets(data){const buckets=new Map(); allMonths.forEach(m=>buckets.set(m,{month:m,'Completed':0,'In Progress':0,'Rejected':0})); data.forEach(r=>{const key=r.eventMonth || 'No Month'; if(!buckets.has(key)) buckets.set(key,{month:key,'Completed':0,'In Progress':0,'Rejected':0}); buckets.get(key)[r.status]=(buckets.get(key)[r.status]||0)+1;}); return Array.from(buckets.values()).filter(x=>x.month!=='No Month' && (x['Completed']||x['In Progress']||x['Rejected'])); }
function weekBuckets(data){const buckets=new Map(); weekOrder.forEach(w=>buckets.set(w,{week:w,'Completed':0,'In Progress':0,'Rejected':0})); data.forEach(r=>{const key=r.eventWeek || 'No Week'; if(!buckets.has(key)) buckets.set(key,{week:key,'Completed':0,'In Progress':0,'Rejected':0}); buckets.get(key)[r.status]=(buckets.get(key)[r.status]||0)+1;}); return Array.from(buckets.values()).filter(x=>(x['Completed']||x['In Progress']||x['Rejected'])); }
function typeBuckets(data){const groups=aggregateByField(data,'gkTypeFilter'); const rows=[]; for(const [name,items] of groups.entries()){ rows.push({name,total:items.length}); } rows.sort((a,b)=>b.total-a.total); return rows.slice(0,10); }
function departmentBuckets(data){const groups=aggregateByField(data,'department'); const rows=[]; for(const [name,items] of groups.entries()){ rows.push({name,total:items.length}); } rows.sort((a,b)=>b.total-a.total); return rows.slice(0,10); }
function statusCounts(data){ const counts={'Completed':0,'In Progress':0,'Rejected':0}; data.forEach(r=> counts[r.status]=(counts[r.status]||0)+1); return counts; }
function updateKpis(data){const counts=statusCounts(data); const total=data.length; const savings=data.reduce((s,d)=>s+(Number(d.completedSavings)||0),0); setMetric('kpiTotal',fmtInt.format(total)); setMetric('kpiInProgress',fmtInt.format(counts['In Progress']||0)); setMetric('kpiCompleted',fmtInt.format(counts['Completed']||0)); setMetric('kpiRejected',fmtInt.format(counts['Rejected']||0)); setMetric('kpiSavings',fmtMoney.format(savings)); setMetric('savingsBadge',fmtMoney.format(savings)); setMetric('volumeBadge',fmtInt.format(total)+' cases'); setMetric('superMonthlyBadge',fmtInt.format(total)+' cases'); const compPct = total ? (counts['Completed']||0)*100/total : 0; setMetric('completionBadge',fmtPct.format(compPct)+'%'); setMetric('statusBadge',fmtPct.format(compPct)+'% completed'); setMetric('detailStatusBadge',fmtPct.format(compPct)+'% completed'); const weeklyRows = weekBuckets(data); setMetric('detailWeeklyBadge',fmtInt.format(weeklyRows.length)+' weeks'); const uniqueTypes=new Set(data.map(d=>d.gkTypeFilter)).size; setMetric('typeBadge',fmtInt.format(uniqueTypes)+' types'); }
function updateSummaries(){ setMetric('supervisorSummary',getSummaryText(state.supervisors.size,allSupervisors.length)); setMetric('ownerSummary',getSummaryText(state.owners.size,allOwners.length)); setMetric('statusSummary',getSummaryText(state.statuses.size,statuses.length)); setMetric('gkTypeSummary',getSummaryText(state.gkTypes.size,allGkTypes.length)); setMetric('monthSummary',getSummaryText(state.months.size,allMonths.length)); setMetric('weekSummary',getSummaryText(state.weeks.size,weekOrder.length)); }
function updateActiveFilters(){ const chips=[]; if(state.supervisors.size!==allSupervisors.length) chips.push(`<span class="chip">Supervisor: ${state.supervisors.size}</span>`); if(state.owners.size!==allOwners.length) chips.push(`<span class="chip">Owner: ${state.owners.size}</span>`); if(state.statuses.size!==statuses.length) chips.push(`<span class="chip">Status: ${state.statuses.size}</span>`); if(state.gkTypes.size!==allGkTypes.length) chips.push(`<span class="chip">GK Type: ${state.gkTypes.size}</span>`); if(state.months.size!==allMonths.length) chips.push(`<span class="chip">Month: ${state.months.size}</span>`); if(state.weeks.size!==weekOrder.length) chips.push(`<span class="chip">Week: ${state.weeks.size}</span>`); if(state.dateFrom || state.dateTo) chips.push(`<span class="chip">Date: ${state.dateFrom||'...'} → ${state.dateTo||'...'}</span>`); document.getElementById('activeFilters').innerHTML = chips.join(''); }
function baseLayout(height=360){ return {paper_bgcolor:'#ffffff',plot_bgcolor:'#ffffff',margin:{l:48,r:20,t:8,b:48},height,legend:{orientation:'h',y:1.12,x:0,font:{size:12}},font:{family:'Inter, Segoe UI, Arial, sans-serif',color:'#0f1721'},xaxis:{gridcolor:'#e9eef4',linecolor:'#cfd8e3',tickfont:{size:12}},yaxis:{gridcolor:'#e9eef4',zerolinecolor:'#e9eef4',tickfont:{size:12}},showlegend:true}; }
function renderMonthlyActivity(data){ const rows=monthBuckets(data); const x=rows.map(r=>r.month); const traces=statuses.map(st=>({type:'bar',name:st,x,y:rows.map(r=>r[st]||0),marker:{color:statusColors[st]}})); Plotly.react('supMonthlyChart', traces, {...baseLayout(380), barmode:'stack'}, {displayModeBar:false,responsive:true}); }
function renderStatusDonut(data,target){ const counts=statusCounts(data); const labels=statuses.filter(s=>counts[s]>0); const values=labels.map(s=>counts[s]); Plotly.react(target,[{type:'pie',hole:.62,labels,values,sort:false,marker:{colors:labels.map(l=>statusColors[l])},textinfo:'label+percent',hovertemplate:'%{label}: %{value} cases<extra></extra>'}],{paper_bgcolor:'#fff',margin:{l:10,r:10,t:10,b:10},showlegend:false,height:320,font:{family:'Inter, Segoe UI, Arial, sans-serif'}},{displayModeBar:false,responsive:true}); }
function renderHorizontalBar(target,rows,labelField,valueField,titleType='currency'){ const labels=rows.map(r=>r.name).reverse(); const values=rows.map(r=>r[valueField]).reverse(); Plotly.react(target,[{type:'bar',orientation:'h',y:labels,x:values,marker:{color:'#2563eb'},text:values.map(v=>titleType==='currency'?fmtMoney.format(v):titleType==='pct'?fmtPct.format(v*100)+'%':fmtInt.format(v)),textposition:'outside',cliponaxis:false,hovertemplate:'%{y}<br>%{x}<extra></extra>'}],{...baseLayout(360),height:360,margin:{l:150,r:40,t:8,b:30},showlegend:false},{displayModeBar:false,responsive:true}); }
function renderGroupedStatusBar(data,target,field,label){ const groups=aggregateByField(data,field); const rows=[]; for(const [name,items] of groups.entries()){ rows.push({name,'Completed':items.filter(d=>d.status==='Completed').length,'In Progress':items.filter(d=>d.status==='In Progress').length,'Rejected':items.filter(d=>d.status==='Rejected').length,total:items.length}); } rows.sort((a,b)=>b.total-a.total); const top=rows.slice(0,10); const x=top.map(r=>r.name); const traces=statuses.map(st=>({type:'bar',name:st,x,y:top.map(r=>r[st]),marker:{color:statusColors[st]}})); Plotly.react(target,traces,{...baseLayout(320),barmode:'stack',margin:{l:40,r:20,t:8,b:110},xaxis:{tickangle:-30,gridcolor:'#e9eef4',linecolor:'#cfd8e3'}},{displayModeBar:false,responsive:true}); }
function renderSupervisorCharts(data){ renderMonthlyActivity(data); renderStatusDonut(data,'supStatusChart'); renderHorizontalBar('supSavingsChart', topGroups(data,'supervisor','savings',10), 'supervisor','savings','currency'); renderHorizontalBar('supCompletionChart', topGroups(data,'supervisor','completionRate',10), 'supervisor','completionRate','pct'); renderHorizontalBar('supCasesChart', topGroups(data,'supervisor','count',10), 'supervisor','total','count'); const typeRows=typeBuckets(data); Plotly.react('supTypeChart',[{type:'bar',x:typeRows.map(r=>r.name),y:typeRows.map(r=>r.total),marker:{color:'#7c3aed'},text:typeRows.map(r=>fmtInt.format(r.total)),textposition:'outside'}],{...baseLayout(360),showlegend:false,margin:{l:44,r:20,t:8,b:90},xaxis:{tickangle:-28,gridcolor:'#e9eef4',linecolor:'#cfd8e3'}},{displayModeBar:false,responsive:true}); }
function renderDetailCharts(data){ const rows=weekBuckets(data); const x=rows.map(r=>r.week); const traces=statuses.map(st=>({type:'bar',name:st,x,y:rows.map(r=>r[st]||0),marker:{color:statusColors[st]}})); Plotly.react('detailWeeklyChart', traces, {...baseLayout(380), barmode:'stack', xaxis:{tickangle:-35,gridcolor:'#e9eef4',linecolor:'#cfd8e3'}}, {displayModeBar:false,responsive:true}); renderStatusDonut(data,'detailStatusChart'); renderHorizontalBar('detailOwnerSavingsChart', topGroups(data,'owner','savings',8), 'owner','savings','currency'); renderHorizontalBar('detailOwnerCasesChart', topGroups(data,'owner','count',8), 'owner','total','count'); const deptRows=departmentBuckets(data); Plotly.react('detailDepartmentChart',[{type:'bar',orientation:'h',y:deptRows.map(r=>r.name).reverse(),x:deptRows.map(r=>r.total).reverse(),marker:{color:'#0ea5a4'},text:deptRows.map(r=>fmtInt.format(r.total)).reverse(),textposition:'outside',cliponaxis:false}],{...baseLayout(320),height:320,margin:{l:150,r:40,t:8,b:30},showlegend:false},{displayModeBar:false,responsive:true}); }
function renderTable(data){ const tableData=getTableData(data); const wrap=document.getElementById('detailTableWrap'); if(!tableData.length){ wrap.innerHTML='<div class="table-empty">No records match the current filters.</div>'; return; } const rows=tableData.slice(0,500).map(r=>`<tr><td>${safeText(r.reference)}</td><td>${safeText(r.title)}</td><td>${safeText(r.submitter)}</td><td>${safeText(r.owner)}</td><td>${safeText(r.supervisor)}</td><td>${safeText(r.department)}</td><td><span class="table-status ${statusClass(r.status)}">${safeText(r.status)}</span></td><td>${safeText(r.gkTypeFilter)}</td><td>${fmtMoney.format(Number(r.completedSavings)||0)}</td><td>${safeText(r.eventDate)}</td></tr>`).join(''); wrap.innerHTML=`<table class="data-table"><thead><tr><th>Reference #</th><th>Improvement title</th><th>Submitter</th><th>GK owner</th><th>Supervisor</th><th>Department</th><th>Status</th><th>GK type</th><th>Completed savings</th><th>Event date</th></tr></thead><tbody>${rows}</tbody></table>`; }
function exportCsv(){ const rows=getTableData(getFilteredData()); if(!rows.length) return; const headers=['Reference #','Improvement Title','Submitter','GK Owner','Supervisor','Department','Status','GK Type','Completed Savings USD','Event Date','Submitted Date','Completed Date','Approval Date']; const escape = (v)=> '"'+String(v ?? '').replaceAll('"','""')+'"'; const csv=[headers.join(',')].concat(rows.map(r=>[r.reference,r.title,r.submitter,r.owner,r.supervisor,r.department,r.status,r.gkTypeFilter,r.completedSavings,r.eventDate,r.submittedDate,r.completedDate,r.approvalDate].map(escape).join(','))).join('\n'); const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='gk_dashboard_filtered_records.csv'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }
function refresh(){ const data=getFilteredData(); updateSummaries(); updateActiveFilters(); updateKpis(data); renderSupervisorCharts(data); renderDetailCharts(data); renderTable(data); }
function syncStateFromInputs(){ state.supervisors=new Set([...document.querySelectorAll('input[data-filter="supervisor"]:checked')].map(el=>el.value)); state.owners=new Set([...document.querySelectorAll('input[data-filter="owner"]:checked')].map(el=>el.value)); state.statuses=new Set([...document.querySelectorAll('input[data-filter="status"]:checked')].map(el=>el.value)); state.gkTypes=new Set([...document.querySelectorAll('input[data-filter="gkType"]:checked')].map(el=>el.value)); state.months=new Set([...document.querySelectorAll('input[data-filter="month"]:checked')].map(el=>el.value)); state.weeks=new Set([...document.querySelectorAll('input[data-filter="week"]:checked')].map(el=>el.value)); state.dateFrom=document.getElementById('dateFrom').value; state.dateTo=document.getElementById('dateTo').value; refresh(); }
function toggleFilterSet(filter,checked){ document.querySelectorAll(`input[data-filter="${filter}"]`).forEach(el=>{ if(el.closest('.check-item')?.style.display==='none') return; el.checked=checked;}); syncStateFromInputs(); }
function resetFilters(){ document.querySelectorAll('input[data-filter]').forEach(el=>el.checked=true); document.querySelectorAll('.filter-search').forEach(input=>input.value=''); document.querySelectorAll('.check-item').forEach(item=>item.style.display='flex'); document.getElementById('dateFrom').value=''; document.getElementById('dateTo').value=''; document.getElementById('tableSearch').value=''; state.tableSearch=''; syncStateFromInputs(); }
function closeOpenFilters(except=null){ document.querySelectorAll('details.multi[open]').forEach(d=>{ if(d!==except) d.open=false; }); }
function initDropdownSearch(){ document.querySelectorAll('.filter-search').forEach(input=>{ input.addEventListener('input',()=>{ const q=input.value.trim().toLowerCase(); const target=document.getElementById(input.dataset.target); if(!target) return; target.querySelectorAll('.check-item').forEach(item=>{ item.style.display=item.textContent.toLowerCase().includes(q)?'flex':'none'; }); }); }); }
function setView(view){ state.activeView=view; document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${view}`)); document.querySelectorAll('[data-view-btn]').forEach(btn=>btn.classList.toggle('active',btn.dataset.viewBtn===view)); if(view==='supervisor'){ document.getElementById('topbarTitle').textContent='Official Dashboard · Supervisor overview'; document.getElementById('heroTitle').textContent='GK Supervisor Performance Overview'; document.getElementById('heroSubtitle').textContent='Track supervisor-level KPI performance, savings, completion rate, and activity trends in one clean dashboard.'; } else { document.getElementById('topbarTitle').textContent='Detail Dashboard · Detailed analytics & records'; document.getElementById('heroTitle').textContent='GK Detail Analytics Dashboard'; document.getElementById('heroSubtitle').textContent='Review weekly activity, owner performance, department trends, and the filtered detail records in one place.'; } window.scrollTo({top:0,behavior:'smooth'}); }
function toggleSidebar(){ if(window.innerWidth<=960){ document.body.classList.toggle('mobile-sidebar-open'); } else { document.body.classList.toggle('sidebar-collapsed'); } }

document.querySelectorAll('input[data-filter]').forEach(el=>el.addEventListener('change',syncStateFromInputs));
document.getElementById('dateFrom').addEventListener('change',syncStateFromInputs);
document.getElementById('dateTo').addEventListener('change',syncStateFromInputs);
document.getElementById('resetFilters').addEventListener('click',resetFilters);
document.getElementById('toggleSidebarBtn').addEventListener('click',toggleSidebar);
document.getElementById('sidebarToggle').addEventListener('click',toggleSidebar);
document.getElementById('mobileNavToggle').addEventListener('click',toggleSidebar);
document.getElementById('exportCsvBtn').addEventListener('click',exportCsv);
document.getElementById('tableSearch').addEventListener('input',e=>{state.tableSearch=e.target.value; renderTable(getFilteredData());});
document.querySelectorAll('.tiny-btn[data-action]').forEach(btn=>btn.addEventListener('click',e=>{ e.preventDefault(); toggleFilterSet(btn.dataset.filter, btn.dataset.action==='all'); }));
document.querySelectorAll('[data-close-filter="true"]').forEach(btn=>btn.addEventListener('click',e=>{ e.preventDefault(); const details=btn.closest('details.multi'); if(details) details.open=false; }));
document.querySelectorAll('details.multi').forEach(details=>details.addEventListener('toggle',()=>{ if(details.open) closeOpenFilters(details); }));
document.addEventListener('click',e=>{ if(!e.target.closest('details.multi')) closeOpenFilters(); });
document.querySelectorAll('[data-view-btn]').forEach(btn=>btn.addEventListener('click',()=>{ setView(btn.dataset.viewBtn); if(window.innerWidth<=960) document.body.classList.remove('mobile-sidebar-open'); }));
initDropdownSearch();
resetFilters();
setView('supervisor');
</script>
</body>
</html>'''

    hero_logo = '<span style="color:rgba(255,255,255,.6);font-size:13px">Logo unavailable</span>'
    if logo_data_uri:
        hero_logo = f'<img src="{logo_data_uri}" alt="Milwaukee logo" />'

    replacements = {
        '__PLOTLY_JS__': plotly_js,
        '__RAW_DATA__': json.dumps(data_records, ensure_ascii=False),
        '__STATUS_COLORS__': json.dumps(STATUS_COLORS),
        '__WEEK_ORDER__': json.dumps(weeks),
        '__SUPERVISORS_JSON__': json.dumps(supervisors, ensure_ascii=False),
        '__OWNERS_JSON__': json.dumps(owners, ensure_ascii=False),
        '__MONTHS_JSON__': json.dumps(months, ensure_ascii=False),
        '__GK_TYPES_JSON__': json.dumps(gk_types, ensure_ascii=False),
        '__SUPERVISOR_CHECKBOXES__': _checkbox_html('supervisor', supervisors),
        '__OWNER_CHECKBOXES__': _checkbox_html('owner', owners),
        '__STATUS_CHECKBOXES__': _checkbox_html('status', statuses),
        '__GK_TYPE_CHECKBOXES__': _checkbox_html('gkType', gk_types),
        '__MONTH_CHECKBOXES__': _checkbox_html('month', months),
        '__WEEK_CHECKBOXES__': _checkbox_html('week', weeks),
        '__HERO_LOGO__': hero_logo,
        '__LATEST_UPDATE_TEXT__': html.escape(latest_update_text or '-'),
        '__SOURCE_NAME__': html.escape(source_name or ''),
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
