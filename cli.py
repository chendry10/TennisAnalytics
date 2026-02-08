import argparse
from pathlib import Path

import pandas as pd

from core.analysis import load_df, summarize_all, export_summary_bytes
from core.errors import DataLoadError, DataValidationError


def write_bytes(data: bytes, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(data)


def main() -> int:
    parser = argparse.ArgumentParser(description="Summarize serve stats from a SwingVision export")
    parser.add_argument("--input", "-i", required=True, help="Path to SwingVision Excel/CSV export")
    parser.add_argument("--sheet", "-s", help="Excel sheet name or index", default=None)
    parser.add_argument("--output", "-o", help="Write summary to CSV/XLSX", default=None)
    parser.add_argument(
        "--format",
        "-f",
        choices=["csv", "xlsx"],
        default="csv",
        help="Output format when --output is provided",
    )
    args = parser.parse_args()

    try:
        sheet = args.sheet
        if sheet is None and Path(args.input).suffix.lower() in {".xlsx", ".xls", ".xlsm"}:
            sheet = 1
        df = load_df(args.input, sheet=sheet, file_name=str(args.input))
        summary = summarize_all(df)
        pd.options.display.float_format = "{:.2f}".format
        print("\nServe Statistics per Player:\n")
        print(summary)

        if args.output:
            data, _ = export_summary_bytes(summary, args.format)
            write_bytes(data, Path(args.output))
            print(f"\nWrote summary to: {args.output}")
        return 0
    except (DataLoadError, DataValidationError) as exc:
        print(f"Error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
