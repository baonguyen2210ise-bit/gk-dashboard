from collections import Counter, defaultdict
from pathlib import Path
import re
import unicodedata

import pandas as pd

from gk_dashboard_core import (
    extract_logo_data_uri,
    format_display_date,
    load_data,
    render_dashboard,
    render_home_dashboard,
    serializable_records,
)

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "data" / "Submitter_Tracking_Master_With_Supervisor.xlsx"
OLD_IDL_FILE = BASE_DIR / "data" / "IDL List.xlsx"
NEW_IDL_FILE = BASE_DIR / "data" / "IDL List Sep2026.xlsx"
LOGO_HTML = BASE_DIR / "data" / "Milwaukee-logo.html"
HOME_OUTPUT = BASE_DIR / "index.html"
OFFICIAL_OUTPUT = BASE_DIR / "official.html"
CUTOFF = pd.Timestamp("2026-09-01")


def _norm(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("Đ", "D").replace("đ", "d")
    text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    text = text.lower().strip()
    text = re.sub(r"\s+\d+\s*$", "", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _employee_code(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text) or text


def _valid_idl_rows(df: pd.DataFrame) -> pd.DataFrame:
    if "Employee Code" not in df.columns:
        return df.iloc[0:0].copy()
    work = df.copy()
    work["__code"] = work["Employee Code"].apply(_employee_code)
    work = work[work["__code"].str.fullmatch(r"\d+")].copy()
    return work.drop_duplicates("__code", keep="last")


def build_roster_context(raw_master: pd.DataFrame, supervisors: list[str]) -> dict:
    """
    Build headcount denominators for the Overview dashboard.

    Sep-2026 onward is direct from the approved Sep IDL file.
    Jan-Aug uses the historical master as the bridge back to the old IDL so the
    roster is grouped under the same six KPI supervisors used by the frozen
    historical dashboard, even though the old workbook still contains legacy
    intermediate supervisor names such as Duong Thanh Phat / Dinh Huu Nhat.
    """
    canonical = {_norm(s): s for s in supervisors}

    # ----- New roster: direct Supervisor assignment from the approved Sep file.
    new_counts = Counter()
    if NEW_IDL_FILE.exists():
        new_idl = _valid_idl_rows(pd.read_excel(NEW_IDL_FILE, sheet_name=0))
        for _, row in new_idl.iterrows():
            sup = canonical.get(_norm(row.get("Supervisor", "")))
            if sup:
                new_counts[sup] += 1

    # ----- Historical name -> KPI supervisor bridge from the frozen master.
    history = raw_master.copy()
    submitted = pd.to_datetime(history.get("Submitted Date"), dayfirst=True, errors="coerce")
    history = history[submitted < CUTOFF].copy()

    name_sup_sets: dict[str, set[str]] = defaultdict(set)
    for _, row in history.iterrows():
        name = _norm(row.get("Supervisor Matched Name", ""))
        sup = canonical.get(_norm(row.get("Supervisor", "")))
        if name and sup:
            name_sup_sets[name].add(sup)
    name_to_sup = {name: next(iter(vals)) for name, vals in name_sup_sets.items() if len(vals) == 1}

    old_counts = Counter()
    if OLD_IDL_FILE.exists():
        old_idl = _valid_idl_rows(pd.read_excel(OLD_IDL_FILE, sheet_name=0))
        staged: list[dict] = []
        raw_group_votes: dict[str, Counter] = defaultdict(Counter)

        # First pass: map people whose own/hierarchy name appeared in historical GK mapping.
        for _, row in old_idl.iterrows():
            hit = ""
            for col in ("Full name", "Line Leader", "Shift Leader", "Supervisor"):
                key = _norm(row.get(col, ""))
                if key in name_to_sup:
                    hit = name_to_sup[key]
                    break
            raw_sup = _norm(row.get("Supervisor", ""))
            staged.append({"row": row, "hit": hit, "raw_sup": raw_sup})
            if hit and raw_sup:
                raw_group_votes[raw_sup][hit] += 1

        # Legacy supervisor groups can be bridged by the dominant frozen KPI owner.
        raw_group_bridge: dict[str, str] = {}
        for raw_sup, votes in raw_group_votes.items():
            if not votes:
                continue
            top_sup, top_n = votes.most_common(1)[0]
            if top_n / sum(votes.values()) >= 0.60:
                raw_group_bridge[raw_sup] = top_sup
        raw_group_bridge.update(canonical)  # Direct six-supervisor names self-map.

        for item in staged:
            row, hit, raw_sup = item["row"], item["hit"], item["raw_sup"]
            if not hit:
                full_key = _norm(row.get("Full name", ""))
                hit = canonical.get(full_key) or raw_group_bridge.get(raw_sup, "")
            if hit:
                old_counts[hit] += 1

    # Guarantee every dashboard supervisor has an explicit numeric denominator.
    return {
        "cutoff": CUTOFF.date().isoformat(),
        "old": {s: int(old_counts.get(s, 0)) for s in supervisors},
        "new": {s: int(new_counts.get(s, 0)) for s in supervisors},
    }


def build() -> tuple[Path, Path]:
    raw_master = pd.read_excel(INPUT_FILE, sheet_name=0)
    df = load_data(INPUT_FILE, sheet_name=0)
    if "Supervisor Display" in df.columns:
        df = df[df["Supervisor Display"].astype(str).str.strip().ne("")].copy()

    records = serializable_records(df)
    latest_dates = df["Submitted Date Parsed"].dropna()
    latest_update = format_display_date(latest_dates.max()) if not latest_dates.empty else ""
    logo_data_uri = extract_logo_data_uri(logo_html=LOGO_HTML)
    supervisors = sorted(df["Supervisor Display"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
    roster_context = build_roster_context(raw_master, supervisors)

    home_html = render_home_dashboard(
        records,
        source_name=INPUT_FILE.name,
        logo_data_uri=logo_data_uri,
        latest_update_text=latest_update,
        roster_context=roster_context,
    )
    official_html = render_dashboard(
        records,
        source_name=INPUT_FILE.name,
        logo_data_uri=logo_data_uri,
        latest_update_text=latest_update,
        show_supervisor_filter=True,
    )

    HOME_OUTPUT.write_text(home_html, encoding="utf-8")
    OFFICIAL_OUTPUT.write_text(official_html, encoding="utf-8")
    print(f"Built {HOME_OUTPUT.name} and {OFFICIAL_OUTPUT.name} with {len(records)} records")
    print(f"Historical roster headcount: {roster_context['old']}")
    print(f"Sep+ roster headcount       : {roster_context['new']}")
    return HOME_OUTPUT, OFFICIAL_OUTPUT


if __name__ == "__main__":
    build()
