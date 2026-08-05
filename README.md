# GK Static Dashboard

This repository is a **read-only, self-contained static dashboard**. It has:

- **one collapsible navigation panel**
- **Official Dashboard** for the supervisor overview
- **Detail Dashboard** for deeper analytics and the filtered record table
- **no backend, no Power Automate, no save/load API, and no data-entry feature**

## Timeline rule

- `Completed` cases use `Approved Date` or `Approval Date` when present; the current export uses `Completed Date` as the approval/completion date.
- `In Progress` and `Rejected` cases use `Submitted Date`.
- If a completed case has none of those dates, the dashboard falls back to `Submitted Date`.

This event date drives the Month, Week, Date Range, and timeline charts.

## Rebuild the static dashboard

```bash
pip install -r requirements.txt
python build_dashboard.py
```

The command regenerates `index.html` from:

```text
data/Submitter_Tracking_Master_With_Supervisor.xlsx
```

## Publish with GitHub Pages

1. Push this folder to a GitHub repository.
2. Open repository **Settings → Pages**.
3. Set **Source** to **GitHub Actions**.
4. Commit new Excel data to trigger the workflow.

No Python server is required after the file has been generated.

## Update source data

Replace or regenerate the master Excel file, then commit it to GitHub or run locally:

```bash
python build_dashboard.py
```

`Zone3_GK_folder_updater.py` is retained for rebuilding the master file from exports and the IDL mapping list.

## Automatic GitHub Pages deployment

This repository includes `.github/workflows/deploy-dashboard.yml`.
Replacing `data/Submitter_Tracking_Master_With_Supervisor.xlsx` and committing the change triggers GitHub Actions to rebuild `index.html` and deploy it to GitHub Pages.

> Important: files committed to a public repository and the generated GitHub Pages website are publicly accessible. Remove or anonymize sensitive employee data before publishing.
