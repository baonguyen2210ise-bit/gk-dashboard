# GK Static Dashboard — restored original layout

This package keeps the original dashboard structure and format:

- `index.html`: Supervisor overview / home page.
- `official.html`: Official detailed dashboard.
- The home page keeps only the **Official Dashboard** button.
- **Create Custom** has been removed.
- **Follow-up GK**, tracking, comments, deadlines, updated-by fields, and related Power Automate behavior have been removed.
- **Pending Timeline** is retained as a read-only view.
- Completed cases use `Approved Date`, `Approval Date`, or `Completed Date` for timeline reporting; other cases use `Submitted Date`.

## Automatic rebuild

Replace this file without changing its name:

```text
data/Submitter_Tracking_Master_With_Supervisor.xlsx
```

Then commit the change to GitHub. The workflow in `.github/workflows/deploy-dashboard.yml` automatically rebuilds and publishes both static pages.

To rebuild locally:

```bash
pip install -r requirements.txt
python build_dashboard.py
```

## GitHub Pages

Use **Settings → Pages → Source: GitHub Actions**.

> The Excel file and generated static pages may expose internal information when the repository or GitHub Pages site is public.
