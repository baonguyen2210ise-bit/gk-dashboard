from pathlib import Path

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
LOGO_HTML = BASE_DIR / "data" / "Milwaukee-logo.html"
HOME_OUTPUT = BASE_DIR / "index.html"
OFFICIAL_OUTPUT = BASE_DIR / "official.html"


def build() -> tuple[Path, Path]:
    df = load_data(INPUT_FILE, sheet_name=0)
    if "Supervisor Display" in df.columns:
        df = df[df["Supervisor Display"].astype(str).str.strip().ne("")].copy()

    records = serializable_records(df)
    latest_dates = df["Event Date"].dropna()
    latest_update = format_display_date(latest_dates.max()) if not latest_dates.empty else ""
    logo_data_uri = extract_logo_data_uri(logo_html=LOGO_HTML)

    home_html = render_home_dashboard(
        records,
        source_name=INPUT_FILE.name,
        logo_data_uri=logo_data_uri,
        latest_update_text=latest_update,
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
    return HOME_OUTPUT, OFFICIAL_OUTPUT


if __name__ == "__main__":
    build()
