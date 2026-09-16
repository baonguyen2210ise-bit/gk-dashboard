# GK Dashboard — automatic raw-data pipeline

## Data flow

```text
GK folder/*.xlsx
        ↓
Zone3_GK_folder_updater.py
        ↓
data/Submitter_Tracking_Master_With_Supervisor.xlsx
        ↓
build_dashboard.py
        ↓
index.html + official.html
        ↓
GitHub Pages
```

## Supervisor mapping periods

- `Submitted Date < 2026-09-01` → `data/IDL List.xlsx`
- `Submitted Date >= 2026-09-01` → `data/IDL List Sep2026.xlsx`

Historical Jan–Aug GK therefore stays assigned to the historical Supervisor structure even after the Sep organization change.

## Identity matching

The updater maps in this order:

1. `GK Owner ID` → `Employee Code`
2. `Submitter ID` → `Employee Code`
3. `GK Owner Name` fallback
4. `Submitter` fallback

Employee Code / MSNV is the primary identity. If a fallback name belongs to multiple employee codes, the updater intentionally leaves the Supervisor blank instead of guessing.

## Updating raw data on GitHub

Put any department GK export Excel files in:

```text
GK folder/
```

The export must contain sheet `GK Ideas` and column `Reference #`.

You may keep multiple dated exports in the folder. If the same `Reference #` exists more than once, a newer filename timestamp such as `20260916_143102` wins.

Every push that changes `GK folder/**` automatically:

1. merges/upserts all raw exports,
2. applies the correct IDL by Submitted Date,
3. regenerates the Master,
4. saves the new Master back to `main`,
5. rebuilds `index.html` and `official.html`,
6. deploys GitHub Pages.

## Required files

```text
data/IDL List.xlsx
data/IDL List Sep2026.xlsx
```

Do not replace the historical IDL when the organization changes. Add a new dated IDL and update the cutoff logic instead.

## GitHub Pages

Use **Settings → Pages → Source: GitHub Actions**.

The repository currently contains internal operational data. If the repository is Public, files committed to it are publicly accessible.
