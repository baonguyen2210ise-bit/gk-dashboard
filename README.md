# GK Static Dashboard

This repository is a **read-only, self-contained static dashboard**. It is read-only and has no backend, external workflow calls, save/load API, or data-entry feature.

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
3. Publish from the branch containing `index.html`, using the repository root folder.

GitHub Pages serves `index.html` directly. No Python server is required after the file has been generated.

## Update source data

Replace or regenerate the master Excel file, then run:

```bash
python build_dashboard.py
```

`Zone3_GK_folder_updater.py` is retained for rebuilding the master file from exports and the IDL mapping list.

## Data visibility warning

The generated `index.html` embeds the summary fields required by the charts. Anyone who can access the GitHub Pages site can inspect those embedded fields. The source Excel files are ignored by Git by default and should remain local. Use a private/internal hosting approach if even the summarized dashboard data is confidential.

## Automatic GitHub Pages deployment

This repository includes `.github/workflows/deploy-dashboard.yml`.
Replacing `data/Submitter_Tracking_Master_With_Supervisor.xlsx` and committing the change triggers GitHub Actions to rebuild `index.html` and deploy it to GitHub Pages.

In **Settings → Pages**, set **Source** to **GitHub Actions**.

> Important: files committed to a public repository and the generated GitHub Pages website are publicly accessible. Remove or anonymize sensitive employee data before publishing.
