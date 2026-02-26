# Timesheet Transformer (Web)

A browser-based tool that transforms worklog CSV data into a DOCX table using a DOCX template.

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

1. **Template DOCX**: Upload your DOCX template file.
2. **Worklog CSV**: Upload your worklog export CSV.
3. **Work Areas CSV (optional)**: Upload work area mapping CSV.
4. (Optional) Enable **Weekly Aggregation**.
5. (Optional) Enable **Include legend** (available only when a Work Areas CSV is selected).
6. Click **Generate DOCX**.
7. After success, the button switches to **Download result**.

### Notes on UI Behavior

- Any file/checkbox change resets the previous result state.
- If Work Areas CSV is provided, warnings are logged for worklog rows with no matching work area key.
- The output filename is based on the selected worklog CSV filename (same base name, `.docx` extension).

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
