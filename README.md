# Timesheet Transformer (Web)

A browser-based tool that transforms worklog CSV data into:

- an Excel table (`.xlsx`), and
- a DOCX table (`.docx`) based on a template.

## Requirements

- Node.js 20+
- npm

## Setup

```bash
npm install
```

## Run Locally

Start development server:

```bash
npm run dev
```

Open the app at:

- `http://localhost:4567`

## How to Use the Webpage

### 1) Worklog Section

1. Upload **Worklog (csv)** (required).
2. Upload **Work Areas (csv)** (optional).
3. Optionally enable:

- **Weekly Aggregation**
- **Include legend** (enabled only when Work Areas file is present)

4. Click **Generate Excel**.
5. After success, click **Download Excel**.

### 2) Timesheet Section

1. Upload **Template (docx)** (required).
2. Click **Generate DOCX**.
3. After success, click **Download DOCX**.

Selected file names are shown next to each file input.

### Notes on UI Behavior

- Any file or checkbox change resets previous generated results.
- The log output is scrollable and records parsing/filtering warnings and export status.
- If Work Areas CSV is provided, warnings are logged for worklog rows with no matching work area key.
- Output filenames are based on the selected worklog CSV filename:
  - `<worklog-name>.xlsx`
  - `<worklog-name>.docx`

## Input Format

### Worklog CSV (required)

Expected headers:

- `User`
- `Worklog`
- `Key`
- `Logged`
- `Date`

### Work Areas CSV (optional)

Expected headers:

- `Key`
- `Name`
- `Alias`

## Output Behavior

- Without Work Areas CSV:
  - The `Bereich` column is not shown.
  - Legend is not included.
- With Work Areas CSV:
  - `Bereich` values are resolved by matching worklog keys.
  - Legend contains only work area keys that are actually referenced in the worklog.
- Summary row keeps the cells that previously contained `Arbeitstage` empty.

### Excel Output Details

- Summary uses an Excel formula in the hours cell (`SUM(...)`) instead of a fixed value.
- Weekly mode:
  - merges first-column cells for consecutive rows in the same week,
  - renders week number and date range with a line break,
  - centers weekly cells horizontally/vertically.
- Worklog and legend tables have black borders, with a visual gap between them.
- Column widths follow DOCX-like proportions.

## Build

```bash
npm run build
```

Preview production build locally:

```bash
npm run preview
```

## Quality Checks

```bash
npm run check
```

`check` includes:

- formatting check
- lint
- typecheck
- tests
- build

## GitHub Pages Deployment

This repository already contains a deployment workflow:

- `.github/workflows/deploy-pages.yml`

One-time setup in your GitHub repository:

1. Open **Settings → Pages**
2. Under **Build and deployment**, set **Source** to **GitHub Actions**

Deploy by:

- pushing to `main`, or
- running **Actions → Deploy to GitHub Pages → Run workflow**

## Project Structure

- `src/`: web app source
- `src/transformer/`: CSV parsing, aggregation, DOCX generation modules
- `tests/`: Vitest test suite
- `sources/`: sample CSV files for manual upload testing
- `dist/`: build output
