from pathlib import Path

from gk_dashboard_core import (
    extract_logo_data_uri,
    format_display_date,
    load_data,
    render_dashboard,
    serializable_records,
)

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "data" / "Submitter_Tracking_Master_With_Supervisor.xlsx"
LOGO_HTML = BASE_DIR / "data" / "Milwaukee-logo.html"
OUTPUT_FILE = BASE_DIR / "index.html"


def build() -> Path:
    df = load_data(INPUT_FILE, sheet_name=0)
    if "Supervisor Display" in df.columns:
        df = df[df["Supervisor Display"].astype(str).str.strip().ne("")].copy()

    records = serializable_records(df)
    latest_dates = df["Event Date"].dropna()
    latest_update = format_display_date(latest_dates.max()) if not latest_dates.empty else ""

    html_text = render_dashboard(
        records,
        source_name=INPUT_FILE.name,
        logo_data_uri=extract_logo_data_uri(logo_html=LOGO_HTML),
        latest_update_text=latest_update,
    )
    OUTPUT_FILE.write_text(html_text, encoding="utf-8")
    print(f"Built {OUTPUT_FILE.name} with {len(records)} records")
    return OUTPUT_FILE


if __name__ == "__main__":
    build()
