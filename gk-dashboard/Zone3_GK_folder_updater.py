from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
import unicodedata
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

import pandas as pd

# ============================================================
# CONFIG
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
GK_FOLDER = BASE_DIR / "GK folder"

# Historical organization / Supervisor mapping.
# All GK submitted before 2026-09-01 continue using this file.
OLD_IDL_LIST_FILE = DATA_DIR / "IDL List.xlsx"

# New organization / Supervisor mapping from Sep-2026 onward.
NEW_IDL_LIST_FILE = DATA_DIR / "IDL List Sep2026.xlsx"
IDL_SHEET_NAME = 0
IDL_CUTOFF_DATE = date(2026, 9, 1)

# Single generated master consumed by build_dashboard.py
OUTPUT_FILE = DATA_DIR / "Submitter_Tracking_Master_With_Supervisor.xlsx"

# Optional historical fallback only if OUTPUT_FILE does not exist.
LEGACY_MASTER_FILE = DATA_DIR / "Submitter_Tracking_Master_2.xlsx"

GK_SHEET_NAME = "GK Ideas"
REFERENCE_COL = "Reference #"
SUBMITTED_DATE_COL = "Submitted Date"

GK_OWNER_ID_COL = "GK Owner ID"
GK_OWNER_COL = "GK Owner Name"
SUBMITTER_ID_COL = "Submitter ID"
SUBMITTER_COL = "Submitter"

SUPERVISOR_COL = "Supervisor"
MATCHED_BY_COL = "Supervisor Matched By"
MATCHED_NAME_COL = "Supervisor Matched Name"
MATCHED_IDL_COL = "Supervisor Matched IDL Column"

EMPLOYEE_CODE_CANDIDATES = ["Employee Code", "Employee ID", "MSNV"]
IDL_NAME_COLUMN_CANDIDATES = [
    "Full name",
    "Line Leader",
    "Shift Leader",
    "Supervisor",
]

INPUT_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
KEEP_ONLY_MAPPED_ROWS = False


# ============================================================
# TEXT / NORMALIZE HELPERS
# ============================================================
def strip_accents(text: str) -> str:
    text = str(text).replace("Đ", "D").replace("đ", "d")
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def remove_parentheses(text: str) -> str:
    return re.sub(r"\s*\([^)]*\)", " ", text)


def remove_trailing_number(text: str) -> str:
    # Hoang Hai Lam 2 -> Hoang Hai Lam
    return re.sub(r"\s+\d+\s*$", "", text).strip()


def basic_clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\u00a0", " ").strip()
    if not text:
        return ""

    text = remove_parentheses(text)
    text = strip_accents(text).lower()
    text = re.sub(r"[^a-z0-9\-\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_key(value) -> str:
    text = basic_clean_text(value)
    text = remove_trailing_number(text)
    text = re.sub(r"[^a-z0-9]", "", text)
    return text


def name_variants(value) -> List[str]:
    """Generate conservative fallback-name variants."""
    text = basic_clean_text(value)
    if not text:
        return []

    variants: Set[str] = set()

    def add(v: str) -> None:
        key = normalize_key(v)
        if key:
            variants.add(key)

    add(text)
    add(remove_trailing_number(text))

    if "-" in text:
        add(text.split("-", 1)[0])

    # Handle an English nickname before a Vietnamese full name.
    words = text.replace("-", " ").split()
    for drop_count in (1, 2):
        if len(words) - drop_count >= 3:
            suffix = " ".join(words[drop_count:])
            add(suffix)
            add(remove_trailing_number(suffix))

    return sorted(variants)


def pretty_no_accent_name(value) -> str:
    text = basic_clean_text(value)
    if not text:
        return ""
    text = text.replace("-", " ")
    text = remove_trailing_number(text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return " ".join(part.capitalize() for part in text.split())


def looks_like_name(value) -> bool:
    text = basic_clean_text(value)
    return bool(re.search(r"[a-z]", text)) and len(normalize_key(text)) >= 3


def normalize_reference(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s.upper()


def normalize_employee_code(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    # IDs in current files are numeric. Keep only a clean numeric ID when possible.
    digits = re.sub(r"\D", "", text)
    return digits if digits else text.lower()


def parse_submitted_date(value) -> Optional[date]:
    if pd.isna(value) or value == "":
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None
    return parsed.date()


# ============================================================
# IO HELPERS
# ============================================================
def find_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    col_map = {basic_clean_text(c): c for c in df.columns}
    for cand in candidates:
        key = basic_clean_text(cand)
        if key in col_map:
            return col_map[key]
    return None


def read_gk_sheet(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=GK_SHEET_NAME)
    if REFERENCE_COL not in df.columns:
        raise ValueError(f"Missing column '{REFERENCE_COL}' in {path.name}")
    return df


def raw_file_sort_key(path: Path) -> Tuple[str, str]:
    """
    Prefer timestamps embedded in export names such as:
      ..._20260916_143102.xlsx
    GitHub checkout does not preserve meaningful local file mtimes, so filename
    timestamps are more reliable than Path.stat().st_mtime inside Actions.
    """
    m = re.search(r"(20\d{6})[_-]?(\d{6})", path.stem)
    stamp = f"{m.group(1)}{m.group(2)}" if m else "00000000000000"
    return stamp, path.name.lower()


def find_input_files(folder: Path) -> List[Path]:
    folder.mkdir(parents=True, exist_ok=True)

    ignored_names = {
        OUTPUT_FILE.name,
        LEGACY_MASTER_FILE.name,
        "Submitter_Tracking_Master_Filtered_3.xlsx",
    }

    files: List[Path] = []
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in INPUT_EXTENSIONS:
            continue
        if p.name.startswith("~$"):
            continue
        if p.name in ignored_names:
            continue
        files.append(p)

    # Oldest -> newest. If Reference # is duplicated, the newest raw export wins.
    files.sort(key=raw_file_sort_key)
    return files


def align_columns(frames: List[pd.DataFrame]) -> List[pd.DataFrame]:
    all_cols: List[str] = []
    seen: Set[str] = set()
    for df in frames:
        for col in df.columns:
            if col not in seen:
                seen.add(col)
                all_cols.append(col)

    aligned: List[pd.DataFrame] = []
    for df in frames:
        work = df.copy()
        for col in all_cols:
            if col not in work.columns:
                work[col] = pd.NA
        aligned.append(work[all_cols])
    return aligned


def upsert_by_reference(existing_df: pd.DataFrame, incoming_df: pd.DataFrame) -> pd.DataFrame:
    if existing_df.empty:
        incoming = incoming_df.copy()
        incoming["__ref_key"] = incoming[REFERENCE_COL].apply(normalize_reference)
        with_ref = incoming[incoming["__ref_key"] != ""].drop_duplicates("__ref_key", keep="last")
        no_ref = incoming[incoming["__ref_key"] == ""]
        return pd.concat([with_ref, no_ref], ignore_index=True).drop(columns=["__ref_key"], errors="ignore")

    existing = existing_df.copy()
    incoming = incoming_df.copy()
    existing["__ref_key"] = existing[REFERENCE_COL].apply(normalize_reference)
    incoming["__ref_key"] = incoming[REFERENCE_COL].apply(normalize_reference)

    incoming_with_ref = incoming[incoming["__ref_key"] != ""].drop_duplicates("__ref_key", keep="last")
    incoming_no_ref = incoming[incoming["__ref_key"] == ""]
    existing_with_ref = existing[existing["__ref_key"] != ""]
    existing_no_ref = existing[existing["__ref_key"] == ""]

    existing_keep = existing_with_ref[~existing_with_ref["__ref_key"].isin(incoming_with_ref["__ref_key"])]
    result = pd.concat(
        [existing_keep, incoming_with_ref, existing_no_ref, incoming_no_ref],
        ignore_index=True,
        sort=False,
    )
    return result.drop(columns=["__ref_key"], errors="ignore")


def read_existing_base() -> pd.DataFrame:
    if OUTPUT_FILE.exists():
        return read_gk_sheet(OUTPUT_FILE)
    if LEGACY_MASTER_FILE.exists():
        return read_gk_sheet(LEGACY_MASTER_FILE)
    return pd.DataFrame()


def save_output(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=GK_SHEET_NAME, index=False)


# ============================================================
# HISTORICAL KPI FREEZE
# ============================================================
def capture_historical_assignments(existing_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    Freeze the Supervisor assignment already stored in the master for all GK
    submitted before the Sep-2026 organization change.

    This is intentional: changing matching logic today must not rewrite Jan-Aug
    KPI ownership that was already reported under the old structure.
    """
    if existing_df.empty or REFERENCE_COL not in existing_df.columns:
        return {}

    frozen: Dict[str, Dict[str, str]] = {}
    cols = [SUPERVISOR_COL, MATCHED_BY_COL, MATCHED_NAME_COL, MATCHED_IDL_COL]

    for _, row in existing_df.iterrows():
        submitted = parse_submitted_date(row.get(SUBMITTED_DATE_COL, ""))
        if submitted is None or submitted >= IDL_CUTOFF_DATE:
            continue

        ref = normalize_reference(row.get(REFERENCE_COL, ""))
        if not ref:
            continue

        frozen[ref] = {
            col: ("" if pd.isna(row.get(col, "")) else str(row.get(col, "")).strip())
            for col in cols
        }

    return frozen


def restore_historical_assignments(
    df: pd.DataFrame,
    frozen: Dict[str, Dict[str, str]],
) -> Tuple[pd.DataFrame, int]:
    if not frozen:
        return df, 0

    work = df.copy()
    restored = 0
    for idx, row in work.iterrows():
        submitted = parse_submitted_date(row.get(SUBMITTED_DATE_COL, ""))
        if submitted is None or submitted >= IDL_CUTOFF_DATE:
            continue

        ref = normalize_reference(row.get(REFERENCE_COL, ""))
        saved = frozen.get(ref)
        if saved is None:
            continue

        for col, value in saved.items():
            work.at[idx, col] = value
        restored += 1

    return work, restored


# ============================================================
# IDL MAPPING
# ============================================================
@dataclass(frozen=True)
class MatchInfo:
    supervisor: str
    matched_name: str
    matched_idl_column: str
    employee_code: str = ""


@dataclass
class IDLMapping:
    by_employee_code: Dict[str, MatchInfo]
    by_name: Dict[str, List[MatchInfo]]
    ambiguous_full_name_keys: Set[str]
    label: str


def _dedupe_matches(items: Sequence[MatchInfo]) -> List[MatchInfo]:
    seen: Set[Tuple[str, str, str, str]] = set()
    out: List[MatchInfo] = []
    for item in items:
        key = (item.supervisor, item.matched_name, item.matched_idl_column, item.employee_code)
        if key not in seen:
            seen.add(key)
            out.append(item)
    return out


def build_idl_mapping(idl_path: Path, label: str) -> IDLMapping:
    if not idl_path.exists():
        raise FileNotFoundError(f"IDL file not found: {idl_path}")

    idl_df = pd.read_excel(idl_path, sheet_name=IDL_SHEET_NAME)
    supervisor_col = find_column(idl_df, ["Supervisor"])
    employee_code_col = find_column(idl_df, EMPLOYEE_CODE_CANDIDATES)

    if supervisor_col is None:
        raise ValueError(f"Missing Supervisor column in {idl_path.name}")
    if employee_code_col is None:
        raise ValueError(f"Missing Employee Code/MSNV column in {idl_path.name}")

    name_cols: List[str] = []
    for cand in IDL_NAME_COLUMN_CANDIDATES:
        col = find_column(idl_df, [cand])
        if col is not None and col not in name_cols:
            name_cols.append(col)
    if not name_cols:
        raise ValueError(f"No name columns found in {idl_path.name}")

    full_name_col = find_column(idl_df, ["Full name"])
    by_code: Dict[str, MatchInfo] = {}
    by_name: Dict[str, List[MatchInfo]] = {}
    direct_full_name_codes: Dict[str, Set[str]] = {}

    for _, row in idl_df.iterrows():
        supervisor_raw = row.get(supervisor_col, "")
        if not looks_like_name(supervisor_raw):
            continue
        supervisor_std = pretty_no_accent_name(supervisor_raw)

        code = normalize_employee_code(row.get(employee_code_col, ""))
        full_name_raw = row.get(full_name_col, "") if full_name_col else ""
        full_name_std = pretty_no_accent_name(full_name_raw) if looks_like_name(full_name_raw) else ""

        if code and full_name_std:
            # Employee Code is the primary unique identity.
            by_code[code] = MatchInfo(
                supervisor=supervisor_std,
                matched_name=full_name_std,
                matched_idl_column=employee_code_col,
                employee_code=code,
            )

            # Detect same-name people with different MSNVs. These must never be
            # auto-selected by name fallback (e.g. multiple Le Thanh Dat).
            for key in name_variants(full_name_raw):
                direct_full_name_codes.setdefault(key, set()).add(code)

        for col in name_cols:
            person_raw = row.get(col, "")
            if not looks_like_name(person_raw):
                continue
            person_std = pretty_no_accent_name(person_raw)
            info = MatchInfo(
                supervisor=supervisor_std,
                matched_name=person_std,
                matched_idl_column=col,
                employee_code=code if col == full_name_col else "",
            )
            for key in name_variants(person_raw):
                by_name.setdefault(key, []).append(info)

    by_name = {key: _dedupe_matches(items) for key, items in by_name.items()}
    ambiguous_full_name_keys = {key for key, codes in direct_full_name_codes.items() if len(codes) > 1}

    return IDLMapping(
        by_employee_code=by_code,
        by_name=by_name,
        ambiguous_full_name_keys=ambiguous_full_name_keys,
        label=label,
    )


def lookup_employee_code(value, mapping: IDLMapping) -> Optional[MatchInfo]:
    code = normalize_employee_code(value)
    if not code:
        return None
    return mapping.by_employee_code.get(code)


def lookup_name(value, mapping: IDLMapping) -> Tuple[Optional[MatchInfo], bool]:
    """Return (match, ambiguous). Ambiguous name fallback is intentionally not auto-mapped."""
    for key in name_variants(value):
        if key in mapping.ambiguous_full_name_keys:
            return None, True

        candidates = mapping.by_name.get(key, [])
        if not candidates:
            continue

        # If all candidate rows resolve to one Supervisor, prefer a direct Full name
        # match, then hierarchy columns. Otherwise the name is unsafe to auto-map.
        supervisors = {x.supervisor for x in candidates if x.supervisor}
        if len(supervisors) != 1:
            return None, True

        priority = {"Full name": 0, "Line Leader": 1, "Shift Leader": 2, "Supervisor": 3}
        chosen = sorted(candidates, key=lambda x: priority.get(x.matched_idl_column, 99))[0]
        return chosen, False

    return None, False


def choose_idl_mapping(submitted_date_value, old_mapping: IDLMapping, new_mapping: IDLMapping) -> IDLMapping:
    submitted = parse_submitted_date(submitted_date_value)
    if submitted is not None and submitted >= IDL_CUTOFF_DATE:
        return new_mapping
    # Missing/unparseable date defaults to historical mapping to avoid moving old KPI.
    return old_mapping


def add_supervisor_columns(
    gk_df: pd.DataFrame,
    old_mapping: IDLMapping,
    new_mapping: IDLMapping,
) -> Tuple[pd.DataFrame, int]:
    required = [GK_OWNER_COL, SUBMITTER_COL]
    for col in required:
        if col not in gk_df.columns:
            raise ValueError(f"Missing column '{col}' in GK data")

    work = gk_df.copy()
    work = work.drop(
        columns=[SUPERVISOR_COL, MATCHED_BY_COL, MATCHED_NAME_COL, MATCHED_IDL_COL],
        errors="ignore",
    )

    supervisors: List[str] = []
    matched_by: List[str] = []
    matched_names: List[str] = []
    matched_idl_cols: List[str] = []
    ambiguous_name_rows = 0

    for _, row in work.iterrows():
        mapping = choose_idl_mapping(row.get(SUBMITTED_DATE_COL, ""), old_mapping, new_mapping)

        hit: Optional[MatchInfo] = None
        by = ""
        ambiguous = False

        # 1) Employee Code / MSNV is the primary identity.
        if GK_OWNER_ID_COL in work.columns:
            hit = lookup_employee_code(row.get(GK_OWNER_ID_COL, ""), mapping)
            if hit:
                by = GK_OWNER_ID_COL

        if hit is None and SUBMITTER_ID_COL in work.columns:
            hit = lookup_employee_code(row.get(SUBMITTER_ID_COL, ""), mapping)
            if hit:
                by = SUBMITTER_ID_COL

        # 2) Fallback to name only when ID does not map.
        if hit is None:
            hit, ambiguous = lookup_name(row.get(GK_OWNER_COL, ""), mapping)
            if hit:
                by = GK_OWNER_COL

        if hit is None and not ambiguous:
            hit, ambiguous = lookup_name(row.get(SUBMITTER_COL, ""), mapping)
            if hit:
                by = SUBMITTER_COL

        if hit:
            supervisors.append(hit.supervisor)
            matched_by.append(by)
            matched_names.append(hit.matched_name)
            matched_idl_cols.append(hit.matched_idl_column)
        else:
            supervisors.append("")
            matched_by.append("Ambiguous Name" if ambiguous else "")
            matched_names.append("")
            matched_idl_cols.append(mapping.label)
            if ambiguous:
                ambiguous_name_rows += 1

    work[SUPERVISOR_COL] = supervisors
    work[MATCHED_BY_COL] = matched_by
    work[MATCHED_NAME_COL] = matched_names
    work[MATCHED_IDL_COL] = matched_idl_cols

    if KEEP_ONLY_MAPPED_ROWS:
        work = work[work[SUPERVISOR_COL].astype(str).str.strip() != ""].copy()

    return work, ambiguous_name_rows


# ============================================================
# MAIN FLOW
# ============================================================
def main() -> None:
    old_mapping = build_idl_mapping(OLD_IDL_LIST_FILE, "IDL List.xlsx")
    new_mapping = build_idl_mapping(NEW_IDL_LIST_FILE, "IDL List Sep2026.xlsx")

    input_files = find_input_files(GK_FOLDER)
    existing_base = read_existing_base()
    frozen_historical = capture_historical_assignments(existing_base)

    incoming_frames: List[pd.DataFrame] = []
    for fp in input_files:
        incoming_frames.append(read_gk_sheet(fp))

    if incoming_frames:
        frames_to_align = [existing_base] if not existing_base.empty else []
        frames_to_align.extend(incoming_frames)
        aligned = align_columns(frames_to_align)

        if existing_base.empty:
            existing_aligned = pd.DataFrame(columns=aligned[0].columns)
            incoming_aligned = aligned
        else:
            existing_aligned = aligned[0]
            incoming_aligned = aligned[1:]

        incoming_all = pd.concat(incoming_aligned, ignore_index=True, sort=False)
        updated_df = upsert_by_reference(existing_aligned, incoming_all)
    else:
        # Useful when only an IDL file or mapping code changes: remap the existing master.
        updated_df = existing_base.copy()

    if updated_df.empty:
        raise ValueError(
            "No GK data found. Upload raw export files into 'GK folder' or provide an existing master."
        )

    final_df, ambiguous_name_rows = add_supervisor_columns(updated_df, old_mapping, new_mapping)
    final_df, restored_historical_rows = restore_historical_assignments(final_df, frozen_historical)
    save_output(final_df, OUTPUT_FILE)

    mapped_count = int((final_df[SUPERVISOR_COL].astype(str).str.strip() != "").sum())
    blank_count = len(final_df) - mapped_count
    new_period_count = 0
    if SUBMITTED_DATE_COL in final_df.columns:
        new_period_count = sum(
            1
            for v in final_df[SUBMITTED_DATE_COL]
            if (parse_submitted_date(v) or date.min) >= IDL_CUTOFF_DATE
        )

    print("Done.")
    print(f"GK folder              : {GK_FOLDER}")
    print(f"Old IDL                : {OLD_IDL_LIST_FILE.name}")
    print(f"New IDL                : {NEW_IDL_LIST_FILE.name}")
    print(f"IDL cutoff             : {IDL_CUTOFF_DATE.isoformat()}")
    print(f"Input raw files        : {len(input_files)}")
    for i, fp in enumerate(input_files, start=1):
        print(f"  {i}. {fp.relative_to(BASE_DIR)}")
    print(f"Output file            : {OUTPUT_FILE}")
    print(f"Output rows            : {len(final_df)}")
    print(f"Rows from Sep mapping  : {new_period_count}")
    print(f"Frozen Jan-Aug rows    : {restored_historical_rows}")
    print(f"Mapped rows            : {mapped_count}")
    print(f"Blank Supervisor       : {blank_count}")
    print(f"Ambiguous-name fallbacks: {ambiguous_name_rows}")


if __name__ == "__main__":
    main()
