from pathlib import Path
import re
import unicodedata
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

# ============================================================
# CONFIG
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
GK_FOLDER = BASE_DIR / "GK folder"

# File IDL list: đặt cùng folder với file .py này
IDL_LIST_FILE = DATA_DIR / "IDL List.xlsx"
IDL_SHEET_NAME = 0

# Chỉ tạo 1 file output duy nhất
OUTPUT_FILE = DATA_DIR / "Submitter_Tracking_Master_With_Supervisor.xlsx"

# Nếu OUTPUT_FILE chưa tồn tại, code có thể lấy file master cũ làm base để không mất history.
# Code CHỈ đọc file này khi cần, không ghi/không tạo lại file này.
LEGACY_MASTER_FILE = DATA_DIR / "Submitter_Tracking_Master_2.xlsx"

GK_SHEET_NAME = "GK Ideas"
REFERENCE_COL = "Reference #"

GK_OWNER_COL = "GK Owner Name"
SUBMITTER_COL = "Submitter"

SUPERVISOR_COL = "Supervisor"
MATCHED_BY_COL = "Supervisor Matched By"
MATCHED_NAME_COL = "Supervisor Matched Name"
MATCHED_IDL_COL = "Supervisor Matched IDL Column"

INPUT_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}

# False = giữ toàn bộ raw data, dòng nào không map được thì Supervisor để trống.
# True  = chỉ giữ những dòng map được Supervisor.
KEEP_ONLY_MAPPED_ROWS = False

# Các cột trong IDL List được dùng để dò tên người.
# Nếu GK Owner Name không match, mới dò tiếp Submitter.
IDL_NAME_COLUMN_CANDIDATES = [
    "Full name",
]


# ============================================================
# TEXT / NORMALIZE HELPERS
# ============================================================
def strip_accents(text: str) -> str:
    """
    Convert Vietnamese text to no-accent text.

    Important:
    Unicode normalization removes accents for most Vietnamese letters,
    but it does NOT convert Đ/đ to D/d automatically.
    Without this replacement:
        Đặng -> ang
        DANG -> dang
    so names like "Võ Văn Đặng Anh" will not match "VO VAN DANG ANH".
    """
    text = str(text).replace("Đ", "D").replace("đ", "d")
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def remove_parentheses(text: str) -> str:
    return re.sub(r"\s*\([^)]*\)", " ", text)


def remove_trailing_number(text: str) -> str:
    # Hoang Hai Lam 2 -> Hoang Hai Lam
    return re.sub(r"\s+\d+\s*$", "", text).strip()


def basic_clean_text(value) -> str:
    """Clean text but still keep spaces between words."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\u00a0", " ").strip()
    if not text:
        return ""

    text = remove_parentheses(text)
    text = strip_accents(text)
    text = text.lower()

    # Một số tên trong file có dạng "Nguyen Thi Phu - Mi".
    # Khi tạo key chính vẫn giữ đủ text; variant bên dưới sẽ tạo thêm key bỏ phần sau dấu -.
    text = re.sub(r"[^a-z0-9\-\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_key(value) -> str:
    """
    Key để match tên:
    - bỏ dấu tiếng Việt
    - lowercase
    - bỏ số cuối tên, ví dụ Hoang Hai Lam 2 -> Hoang Hai Lam
    - bỏ toàn bộ ký tự không phải a-z/0-9
    - bỏ khoảng trắng để match được vokhanhhoa / truongkimchung
    """
    text = basic_clean_text(value)
    text = remove_trailing_number(text)
    text = re.sub(r"[^a-z0-9]", "", text)
    return text


def name_variants(value) -> List[str]:
    """
    Tạo nhiều key để tăng khả năng match tên:
    - tên viết liền / có dấu / không dấu
    - có số cuối, ví dụ Hoang Hai Lam 2
    - có dấu '-', ví dụ Nguyen Thi Phu - Mi
    - có English name đứng đầu, ví dụ:
        Dia Le Thanh Vi -> Le Thanh Vi
        Benz Huynh Lam Huyen Mai -> Huynh Lam Huyen Mai
    """
    text = basic_clean_text(value)
    if not text:
        return []

    variants = set()

    def add(v: str) -> None:
        key = normalize_key(v)
        if key:
            variants.add(key)

    # 1. Tên đầy đủ như hiện tại
    add(text)
    add(remove_trailing_number(text))

    # 2. Nếu có dấu '-' thì lấy thêm phần trước dấu '-'
    # VD: Nguyen Thi Phu - Mi -> Nguyen Thi Phu
    if "-" in text:
        add(text.split("-", 1)[0])

    # 3. Xử lý English name / nickname đứng đầu
    # VD: Dia Le Thanh Vi -> Le Thanh Vi
    # VD: Benz Huynh Lam Huyen Mai -> Huynh Lam Huyen Mai
    words = text.replace("-", " ").split()

    # Chỉ tạo suffix nếu phần còn lại vẫn giống một tên Việt Nam hợp lý
    # Thường >= 3 words để tránh match nhầm.
    for drop_count in [1, 2]:
        if len(words) - drop_count >= 3:
            suffix_name = " ".join(words[drop_count:])
            add(suffix_name)
            add(remove_trailing_number(suffix_name))

    return sorted(variants)

def pretty_no_accent_name(value) -> str:
    """Tên chuẩn để ghi ra file output: không dấu, Title Case, bỏ số cuối."""
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
    # bỏ các dòng header phụ kiểu 1, 2, 3, 9...
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


# ============================================================
# IO HELPERS
# ============================================================
def find_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    """Find column by case-insensitive normalized name."""
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
        raise ValueError(f"Không tìm thấy cột '{REFERENCE_COL}' trong file: {path.name}")
    return df


def find_input_files(folder: Path) -> List[Path]:
    if not folder.exists():
        folder.mkdir(parents=True, exist_ok=True)
        return []

    ignored_names = {
        OUTPUT_FILE.name,
        LEGACY_MASTER_FILE.name,
        "Submitter_Tracking_Master_Filtered_3.xlsx",
    }

    files = []
    for p in folder.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in INPUT_EXTENSIONS:
            continue
        if p.name.startswith("~$"):
            continue
        if p.name in ignored_names:
            continue
        files.append(p)

    # Oldest -> newest, newest wins when duplicate Reference # exists
    files.sort(key=lambda x: (x.stat().st_mtime, x.name.lower()))
    return files


def align_columns(frames: List[pd.DataFrame]) -> List[pd.DataFrame]:
    all_cols = []
    seen = set()

    for df in frames:
        for col in df.columns:
            if col not in seen:
                seen.add(col)
                all_cols.append(col)

    aligned = []
    for df in frames:
        work = df.copy()
        for col in all_cols:
            if col not in work.columns:
                work[col] = pd.NA
        aligned.append(work[all_cols])

    return aligned


def upsert_by_reference(existing_df: pd.DataFrame, incoming_df: pd.DataFrame) -> pd.DataFrame:
    if existing_df.empty:
        return incoming_df.copy().drop(columns=["__ref_key"], errors="ignore")

    existing = existing_df.copy()
    incoming = incoming_df.copy()

    existing["__ref_key"] = existing[REFERENCE_COL].apply(normalize_reference)
    incoming["__ref_key"] = incoming[REFERENCE_COL].apply(normalize_reference)

    incoming_with_ref = incoming[incoming["__ref_key"] != ""].copy()
    incoming_no_ref = incoming[incoming["__ref_key"] == ""].copy()

    # Trong file incoming, nếu trùng Reference # thì dòng cuối cùng thắng
    incoming_with_ref = incoming_with_ref.drop_duplicates(subset="__ref_key", keep="last")

    existing_with_ref = existing[existing["__ref_key"] != ""].copy()
    existing_no_ref = existing[existing["__ref_key"] == ""].copy()

    existing_keep = existing_with_ref[~existing_with_ref["__ref_key"].isin(incoming_with_ref["__ref_key"])]

    result = pd.concat(
        [existing_keep, incoming_with_ref, existing_no_ref, incoming_no_ref],
        ignore_index=True,
        sort=False,
    )

    return result.drop(columns=["__ref_key"], errors="ignore")


def read_existing_base() -> pd.DataFrame:
    """
    Chỉ dùng 1 file output mới.
    Nếu output đã có -> lấy output làm base để upsert.
    Nếu output chưa có nhưng master cũ có -> lấy master cũ làm base để migrate.
    """
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
# IDL MAPPING LOGIC
# ============================================================
def build_idl_mapping(idl_path: Path) -> Dict[str, Dict[str, str]]:
    if not idl_path.exists():
        raise FileNotFoundError(
            f"Không tìm thấy file IDL List: {idl_path}\n"
            "Hãy đặt file 'IDL List.xlsx' cùng folder với file .py."
        )

    idl_df = pd.read_excel(idl_path, sheet_name=IDL_SHEET_NAME)

    supervisor_col = find_column(idl_df, ["Supervisor"])
    if supervisor_col is None:
        raise ValueError("Không tìm thấy cột 'Supervisor' trong IDL List.xlsx")

    name_cols = []
    for cand in IDL_NAME_COLUMN_CANDIDATES:
        col = find_column(idl_df, [cand])
        if col is not None and col not in name_cols:
            name_cols.append(col)

    if not name_cols:
        raise ValueError("Không tìm thấy cột tên nào trong IDL List.xlsx")

    mapping: Dict[str, Dict[str, str]] = {}

    for _, row in idl_df.iterrows():
        supervisor_raw = row.get(supervisor_col, "")
        if not looks_like_name(supervisor_raw):
            continue

        supervisor_std = pretty_no_accent_name(supervisor_raw)

        for col in name_cols:
            person_raw = row.get(col, "")
            if not looks_like_name(person_raw):
                continue

            person_std = pretty_no_accent_name(person_raw)

            for key in name_variants(person_raw):
                if not key:
                    continue

                # Nếu cùng 1 người xuất hiện nhiều dòng, giữ mapping đầu tiên.
                # Trường hợp Hoang Hai Lam / Hoang Hai Lam 2 đã được gom qua pretty_no_accent_name.
                mapping.setdefault(
                    key,
                    {
                        "supervisor": supervisor_std,
                        "matched_name": person_std,
                        "matched_idl_column": col,
                    },
                )

    return mapping


def lookup_person(value, mapping: Dict[str, Dict[str, str]]) -> Optional[Dict[str, str]]:
    for key in name_variants(value):
        if key in mapping:
            return mapping[key]
    return None


def add_supervisor_columns(gk_df: pd.DataFrame, mapping: Dict[str, Dict[str, str]]) -> pd.DataFrame:
    required_cols = [GK_OWNER_COL, SUBMITTER_COL]
    for col in required_cols:
        if col not in gk_df.columns:
            raise ValueError(f"Không tìm thấy cột '{col}' trong GK data")

    work = gk_df.copy()

    # Xóa các cột mapping cũ nếu có, sau đó add lại để không giữ dữ liệu mapping cũ.
    work = work.drop(
        columns=[SUPERVISOR_COL, MATCHED_BY_COL, MATCHED_NAME_COL, MATCHED_IDL_COL],
        errors="ignore",
    )

    supervisors = []
    matched_by = []
    matched_names = []
    matched_idl_cols = []

    for _, row in work.iterrows():
        owner_value = row.get(GK_OWNER_COL, "")
        submitter_value = row.get(SUBMITTER_COL, "")

        # Ưu tiên match bằng GK Owner Name trước.
        hit = lookup_person(owner_value, mapping)
        by = GK_OWNER_COL if hit else ""

        # Nếu GK Owner không match thì mới match Submitter.
        if hit is None:
            hit = lookup_person(submitter_value, mapping)
            by = SUBMITTER_COL if hit else ""

        if hit:
            supervisors.append(hit["supervisor"])
            matched_by.append(by)
            matched_names.append(hit["matched_name"])
            matched_idl_cols.append(hit["matched_idl_column"])
        else:
            supervisors.append("")
            matched_by.append("")
            matched_names.append("")
            matched_idl_cols.append("")

    work[SUPERVISOR_COL] = supervisors
    work[MATCHED_BY_COL] = matched_by
    work[MATCHED_NAME_COL] = matched_names
    work[MATCHED_IDL_COL] = matched_idl_cols

    if KEEP_ONLY_MAPPED_ROWS:
        work = work[work[SUPERVISOR_COL].astype(str).str.strip() != ""].copy()

    return work


# ============================================================
# MAIN FLOW
# ============================================================
def main() -> None:
    idl_mapping = build_idl_mapping(IDL_LIST_FILE)

    input_files = find_input_files(GK_FOLDER)
    existing_base = read_existing_base()

    incoming_frames = []
    for fp in input_files:
        df = read_gk_sheet(fp)
        incoming_frames.append(df)

    if incoming_frames:
        frames_to_align = [existing_base] if not existing_base.empty else []
        frames_to_align.extend(incoming_frames)
        aligned_frames = align_columns(frames_to_align)

        if existing_base.empty:
            existing_aligned = pd.DataFrame(columns=aligned_frames[0].columns)
            incoming_aligned = aligned_frames
        else:
            existing_aligned = aligned_frames[0]
            incoming_aligned = aligned_frames[1:]

        incoming_all = pd.concat(incoming_aligned, ignore_index=True, sort=False)
        updated_df = upsert_by_reference(existing_aligned, incoming_all)
    else:
        updated_df = existing_base.copy()

    if updated_df.empty:
        raise ValueError(
            "Không có dữ liệu để cập nhật. "
            "Hãy đặt file export GK vào thư mục 'GK folder' hoặc kiểm tra file output/master hiện có."
        )

    final_df = add_supervisor_columns(updated_df, idl_mapping)
    save_output(final_df, OUTPUT_FILE)

    mapped_count = int((final_df[SUPERVISOR_COL].astype(str).str.strip() != "").sum())
    blank_count = len(final_df) - mapped_count

    print("Done.")
    print(f"GK folder       : {GK_FOLDER}")
    print(f"IDL List        : {IDL_LIST_FILE}")
    print(f"Input files     : {len(input_files)}")
    for i, fp in enumerate(input_files, start=1):
        print(f"  {i}. {fp.name}")
    print(f"Output file     : {OUTPUT_FILE}")
    print(f"Output rows     : {len(final_df)}")
    print(f"Mapped rows     : {mapped_count}")
    print(f"Blank Supervisor: {blank_count}")


if __name__ == "__main__":
    main()
