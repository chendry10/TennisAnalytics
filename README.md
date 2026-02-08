# CourtSide Analytics

A polished Streamlit dashboard and CLI for SwingVision tennis serve analysis.

## Quick Start (Web App)
1. Install dependencies: `pip install -r requirements.txt`
2. Run the app: `python.exe -m streamlit run app.py`
3. Drag & drop a SwingVision Excel/CSV export.

## Build Windows EXE (No Python Needed)
1. Install PyInstaller: `pip install pyinstaller`
2. Build the executable:
	`pyinstaller --onefile --name CourtSideAnalytics --collect-all streamlit --add-data "app.py;." --add-data "core;core" launcher.py`
3. Share: `dist\CourtSideAnalytics.exe`

## CLI Usage
Analyze a file and print the summary:
`python cli.py --input "path/to/SwingVision-export.xlsx"`

Write a CSV summary:
`python cli.py --input "path/to/SwingVision-export.xlsx" --output "serve_summary.csv" --format csv`

Write an Excel summary:
`python cli.py --input "path/to/SwingVision-export.xlsx" --output "serve_summary.xlsx" --format xlsx`

Optional: specify a sheet by name or index (Excel only):
`python cli.py --input "path/to/SwingVision-export.xlsx" --sheet "Sheet1"`

## Notes
- Required columns: `Player`, `Shot`, `Type`, `Result`, `Point`.
- Supports Excel (.xlsx/.xls/.xlsm) and CSV (.csv).
