# CourtSide Analytics

CourtSide Analytics is a Streamlit dashboard and CLI for analyzing SwingVision exports. It summarizes serve performance, return performance, and rally-length outcomes from a single match or a folder of matches.

## Highlights
- Streamlit dashboard for single-file review or multi-match folder analysis.
- Player filtering with both single-player and compare modes.
- Date-based filtering when filenames include a `YYYY-MM-DD` match date.
- CSV and Excel export for filtered summaries.
- CLI workflow for scripting or batch use.
- Flexible input handling for SwingVision workbooks, raw shots data, and common column aliases.

## Quick Start
Install dependencies and launch the app:

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

Typical dashboard flow:

1. Choose `Single file` or `Folder` in the sidebar.
2. Upload a SwingVision export in `.xlsx`, `.xls`, `.xlsm`, or `.csv` format.
3. Select one player or compare multiple players.
4. Optionally filter matches by date and include or exclude individual files.
5. Download the filtered summary as CSV or Excel.

## What The App Handles
- If an Excel workbook contains `Stats` and `Settings`, the app uses those sheets for fast summary generation.
- If `Shots` data is available, the app augments the summary with return and rally-length metrics.
- If aggregate sheets are missing, the app falls back to raw shot-by-shot analysis.
- If column names differ slightly, the loader attempts to map common aliases automatically.

## CLI Usage
Print a summary to the terminal:

```bash
python cli.py --input "path/to/SwingVision-export.xlsx"
```

Write a CSV summary:

```bash
python cli.py --input "path/to/SwingVision-export.xlsx" --output "serve_summary.csv" --format csv
```

Write an Excel summary:

```bash
python cli.py --input "path/to/SwingVision-export.xlsx" --output "serve_summary.xlsx" --format xlsx
```

Specify a sheet by name or index for Excel input:

```bash
python cli.py --input "path/to/SwingVision-export.xlsx" --sheet "Shots"
```

## Input Expectations
Supported input types:

- Excel: `.xlsx`, `.xls`, `.xlsm`
- CSV: `.csv`

Expected core fields after normalization:

- `Player`
- `Shot`
- `Type`
- `Result`
- `Point`

For folder workflows, filenames that contain a date such as `2026-02-07` can be filtered in the dashboard by date window.

## Build A Windows EXE
This repo already includes [CourtSideAnalytics.spec](CourtSideAnalytics.spec) and [launcher.py](launcher.py) for packaging the Streamlit app.

```bash
pip install pyinstaller
pyinstaller CourtSideAnalytics.spec
```

The built executable will be created at `dist\CourtSideAnalytics.exe`.

## Development
Run the test suite:

```bash
python -m unittest
```

Key files:

- `app.py`: Streamlit dashboard
- `cli.py`: command-line entry point
- `core/analysis.py`: parsing, normalization, and summary logic
- `launcher.py`: packaged app bootstrapper
- `tests/test_analysis.py`: analysis coverage

## Privacy
Real match exports are intentionally not stored in this public repository.

- Keep local workbook files in `Data Files/` or another local folder.
- Spreadsheet exports, local temp spreadsheets, environment files, and Streamlit secrets are gitignored.
