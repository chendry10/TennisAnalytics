from __future__ import annotations

from typing import Dict, Iterable, Optional
import io
from pathlib import Path
import pandas as pd

from .errors import DataLoadError, DataValidationError

REQUIRED_COLUMNS = ["Player", "Shot", "Type", "Result", "Point"]
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
CSV_EXTENSIONS = {".csv"}
GENERIC_OPPONENT_LABELS = {
    "opponent",
    "opp",
    "player2",
    "player 2",
    "player two",
    "other player",
    "unknown",
    "unknown player",
    "n/a",
    "na",
}


def canonicalize_player_label(name: object) -> str:
    """Normalize generic opponent placeholders to a single canonical label."""
    raw = "" if name is None else str(name).strip()
    normalized = " ".join(raw.lower().replace("_", " ").split())
    compact = "".join(ch for ch in normalized if ch.isalnum())

    if not raw:
        return "Opponent"
    if normalized in GENERIC_OPPONENT_LABELS or compact in {"player2", "opponent", "unknown", "na"}:
        return "Opponent"
    return raw


def normalize_summary_players(summary: pd.DataFrame) -> pd.DataFrame:
    """Collapse equivalent placeholder player labels in a summary dataframe."""
    if summary.empty:
        return summary

    grouped = summary.copy()
    grouped["__player__"] = [canonicalize_player_label(player) for player in grouped.index]
    grouped = grouped.groupby("__player__", as_index=True).sum(numeric_only=True).fillna(0)

    def pct(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
        denominator = denominator.replace(0, pd.NA)
        return (numerator / denominator * 100).fillna(0)

    if {"First Serve In", "First Serve Attempts"}.issubset(grouped.columns):
        grouped["Overall First Serve %"] = pct(grouped["First Serve In"], grouped["First Serve Attempts"])
    if {"Second Serve In", "Second Serve Attempts"}.issubset(grouped.columns):
        grouped["Overall Second Serve %"] = pct(grouped["Second Serve In"], grouped["Second Serve Attempts"])
    if {"First Serve Wins", "First Serve Attempts"}.issubset(grouped.columns):
        grouped["First Serve Win %"] = pct(grouped["First Serve Wins"], grouped["First Serve Attempts"])
    if {"Second Serve Wins", "Second Serve Attempts"}.issubset(grouped.columns):
        grouped["Second Serve Win %"] = pct(grouped["Second Serve Wins"], grouped["Second Serve Attempts"])

    grouped.index.name = "Player"
    return grouped


def _excel_engine(file_name: Optional[str]) -> Optional[str]:
    """Return the Excel engine name to use for a given filename (if any)."""
    if _suffix_from_name(file_name) == ".xls":
        return "xlrd"
    return None


def _suffix_from_name(file_name: Optional[str]) -> str:
    """Extract a lowercase file extension from a filename string."""
    if not file_name:
        return ""
    return Path(file_name).suffix.lower()


def _is_excel(file_name: Optional[str]) -> bool:
    """Return True if the filename appears to be an Excel file."""
    return _suffix_from_name(file_name) in EXCEL_EXTENSIONS


def _is_csv(file_name: Optional[str]) -> bool:
    """Return True if the filename appears to be a CSV file."""
    return _suffix_from_name(file_name) in CSV_EXTENSIONS


def _to_bytes(source) -> Optional[bytes]:
    """Normalize file-like or bytes input into raw bytes."""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, "getvalue"):
        return source.getvalue()
    if hasattr(source, "read"):
        return source.read()
    return None


def get_excel_sheet_names(source, file_name: Optional[str] = None) -> list[str]:
    """Return the list of sheet names for an Excel source (path or bytes)."""
    try:
        if isinstance(source, (str, Path)):
            engine = _excel_engine(str(source))
            with pd.ExcelFile(source, engine=engine) as xls:
                return list(xls.sheet_names)
        data = _to_bytes(source)
        if data is None:
            return []
        engine = _excel_engine(file_name)
        with pd.ExcelFile(io.BytesIO(data), engine=engine) as xls:
            return list(xls.sheet_names)
    except Exception as exc:
        raise DataLoadError("Oops! This doesn't look like a valid Excel file.") from exc


def get_raw_columns(source, sheet: Optional[str | int] = None, file_name: Optional[str] = None) -> list[str]:
    """Return the column names from the source without loading full data."""
    try:
        if isinstance(source, (str, Path)):
            if _is_csv(str(source)):
                return list(pd.read_csv(source, nrows=0).columns)
            engine = _excel_engine(str(source))
            return list(pd.read_excel(source, sheet_name=sheet, nrows=0, engine=engine).columns)
        data = _to_bytes(source)
        if data is None:
            return []
        if _is_csv(file_name):
            return list(pd.read_csv(io.BytesIO(data), nrows=0).columns)
        engine = _excel_engine(file_name)
        return list(pd.read_excel(io.BytesIO(data), sheet_name=sheet, nrows=0, engine=engine).columns)
    except Exception as exc:
        raise DataLoadError("Oops! This doesn't look like a SwingVision export.") from exc


def guess_column_map(columns: Iterable[str]) -> Dict[str, str]:
    """Guess a mapping from required column names to source column names."""
    def normalize(value: str) -> str:
        return "".join(ch for ch in str(value).strip().lower() if ch.isalnum())

    normalized = {normalize(col): col for col in columns}

    aliases = {
        "Player": ["player", "playername", "name"],
        "Shot": ["shot", "shotname", "shottype", "stroke"],
        "Type": ["type", "servetype", "stroketype", "shotcategory"],
        "Result": ["result", "outcome", "inout", "winloss"],
        "Point": ["point", "pointid", "pointnumber", "pointno", "pointnum"],
    }

    mapping: Dict[str, str] = {}
    for required in REQUIRED_COLUMNS:
        for alias in aliases.get(required, []):
            if alias in normalized:
                mapping[required] = normalized[alias]
                break
        if required in mapping:
            continue
        for key, original in normalized.items():
            for alias in aliases.get(required, []):
                if alias in key:
                    mapping[required] = original
                    break
            if required in mapping:
                break
    return mapping


def _validate_and_rename(df: pd.DataFrame, column_map: Optional[Dict[str, str]] = None) -> pd.DataFrame:
    """Validate required columns and normalize the dataframe to expected names."""
    if column_map:
        missing = [req for req in REQUIRED_COLUMNS if req not in column_map]
        if missing:
            raise DataValidationError(
                "Missing required column mappings: " + ", ".join(missing)
            )
        unknown = [col for col in column_map.values() if col not in df.columns]
        if unknown:
            raise DataValidationError(
                "Selected columns not found in file: " + ", ".join(unknown)
            )
        df = df.rename(columns={v: k for k, v in column_map.items()})
    else:
        auto_map = guess_column_map(df.columns)
        missing = [req for req in REQUIRED_COLUMNS if req not in auto_map]
        if missing:
            raise DataValidationError(
                "Oops! This doesn't look like a SwingVision export. "
                "Missing columns: " + ", ".join(missing)
                + ". Available columns: " + ", ".join(map(str, df.columns))
            )
        df = df.rename(columns={v: k for k, v in auto_map.items()})

    return df[REQUIRED_COLUMNS]


def _read_dataframe(source, sheet: Optional[str | int] = None, file_name: Optional[str] = None) -> pd.DataFrame:
    """Load a dataframe from a path or in-memory bytes payload."""
    try:
        if isinstance(source, (str, Path)):
            if _is_csv(str(source)):
                return pd.read_csv(source)
            engine = _excel_engine(str(source))
            return pd.read_excel(source, sheet_name=sheet, engine=engine)

        data = _to_bytes(source)
        if data is None:
            raise DataLoadError("No data was provided.")

        if _is_csv(file_name):
            return pd.read_csv(io.BytesIO(data))

        try:
            engine = _excel_engine(file_name)
            return pd.read_excel(io.BytesIO(data), sheet_name=sheet, engine=engine)
        except Exception:
            return pd.read_csv(io.BytesIO(data))
    except DataLoadError:
        raise
    except Exception as exc:
        raise DataLoadError("Oops! This doesn't look like a valid file.") from exc


def _read_excel_sheet(source, sheet_name: str, file_name: Optional[str] = None) -> pd.DataFrame:
    """Read a named Excel sheet from a path or in-memory bytes payload."""
    try:
        if isinstance(source, (str, Path)):
            engine = _excel_engine(str(source))
            return pd.read_excel(source, sheet_name=sheet_name, engine=engine)

        data = _to_bytes(source)
        if data is None:
            raise DataLoadError("No data was provided.")

        engine = _excel_engine(file_name)
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, engine=engine)
    except DataLoadError:
        raise
    except Exception as exc:
        raise DataLoadError("Oops! This doesn't look like a valid Excel file.") from exc


def load_df(source, sheet: Optional[str | int] = None, column_map: Optional[Dict[str, str]] = None,
            file_name: Optional[str] = None) -> pd.DataFrame:
    """Load data and return a cleaned dataframe with required columns."""
    df = _read_dataframe(source, sheet=sheet, file_name=file_name)
    df = _validate_and_rename(df, column_map=column_map)
    return df.dropna(subset=["Point"])


def get_point_servers(df: pd.DataFrame) -> pd.DataFrame:
    """Return one row per point with the server and serve type."""
    serves = df[df["Type"].isin(["first_serve", "second_serve"])]
    last_serves = serves.groupby("Point", sort=False).last().reset_index()
    return last_serves.rename(columns={"Player": "Server", "Type": "Serve_Type"})[
        ["Point", "Server", "Serve_Type"]
    ]


def get_point_winners(df: pd.DataFrame) -> pd.DataFrame:
    """Infer the winner for each point based on the last shot result."""
    last_shots = df.groupby("Point", sort=False).last().reset_index()
    players_map = df.groupby("Point")["Player"].unique().to_dict()

    def infer_winner(row):
        res = row["Result"]
        last_player = row["Player"]
        if res == "In":
            return last_player
        if res in ("Out", "Net"):
            other = [p for p in players_map.get(row["Point"], []) if p != last_player]
            return other[0] if other else None
        return None

    last_shots["Winner"] = last_shots.apply(infer_winner, axis=1)
    return last_shots[["Point", "Winner"]].dropna()


def build_serve_win_data(df: pd.DataFrame) -> pd.DataFrame:
    """Combine servers and winners to compute serve-win outcomes."""
    servers = get_point_servers(df)
    winners = get_point_winners(df)
    return pd.merge(servers, winners, on="Point")


def calculate_serve_win_percentages(serve_win: pd.DataFrame) -> pd.DataFrame:
    """Calculate first/second serve win rates and attempt counts per player."""
    if serve_win.empty:
        return pd.DataFrame(
            columns=[
                "First Serve Win %",
                "Second Serve Win %",
            ]
        ).rename_axis("Player")

    serve_win["Server_Won"] = serve_win["Server"] == serve_win["Winner"]
    grouped = (
        serve_win.groupby(["Server", "Serve_Type"])
        .agg(attempts=("Point", "size"), wins=("Server_Won", "sum"))
        .reset_index()
    )

    attempts_pivot = grouped.pivot(index="Server", columns="Serve_Type", values="attempts").fillna(0)
    wins_pivot = grouped.pivot(index="Server", columns="Serve_Type", values="wins").fillna(0)

    def pct(wins, attempts):
        return (wins / attempts * 100).replace([float("inf")], 0).fillna(0)

    first_attempts = attempts_pivot.get("first_serve", pd.Series(0, index=attempts_pivot.index))
    second_attempts = attempts_pivot.get("second_serve", pd.Series(0, index=attempts_pivot.index))
    first_wins = wins_pivot.get("first_serve", pd.Series(0, index=wins_pivot.index))
    second_wins = wins_pivot.get("second_serve", pd.Series(0, index=wins_pivot.index))

    result = pd.DataFrame(
        {
            "First Serve Win %": pct(first_wins, first_attempts),
            "Second Serve Win %": pct(second_wins, second_attempts),
        },
        index=attempts_pivot.index,
    ).rename_axis("Player")

    return result


def calculate_serve_win_counts(serve_win: pd.DataFrame) -> pd.DataFrame:
    """Count first/second serve wins per player based on point outcomes."""
    if serve_win.empty:
        return pd.DataFrame(columns=["First Serve Wins", "Second Serve Wins"]).rename_axis("Player")

    serve_win = serve_win.copy()
    serve_win["Server_Won"] = serve_win["Server"] == serve_win["Winner"]
    grouped = (
        serve_win.groupby(["Server", "Serve_Type"])["Server_Won"]
        .sum()
        .reset_index(name="wins")
    )

    wins = grouped.pivot(index="Server", columns="Serve_Type", values="wins").fillna(0)
    first_wins = wins.get("first_serve", pd.Series(0, index=wins.index)).astype(int)
    second_wins = wins.get("second_serve", pd.Series(0, index=wins.index)).astype(int)

    return pd.DataFrame(
        {
            "First Serve Wins": first_wins,
            "Second Serve Wins": second_wins,
        }
    ).rename_axis("Player")


def calculate_serve_attempts(df: pd.DataFrame) -> pd.DataFrame:
    """Count first/second serve attempts per player from all serve rows."""
    serves = df[df["Type"].isin(["first_serve", "second_serve"])]
    if serves.empty:
        return pd.DataFrame(columns=["First Serve Attempts", "Second Serve Attempts"]).rename_axis(
            "Player"
        )

    grouped = (
        serves.groupby(["Player", "Type"])["Point"]
        .size()
        .reset_index(name="attempts")
    )

    attempts = grouped.pivot(index="Player", columns="Type", values="attempts").fillna(0)
    first_attempts = attempts.get("first_serve", pd.Series(0, index=attempts.index)).astype(int)
    second_attempts = attempts.get("second_serve", pd.Series(0, index=attempts.index)).astype(int)

    return pd.DataFrame(
        {
            "First Serve Attempts": first_attempts,
            "Second Serve Attempts": second_attempts,
        }
    )


def calculate_serve_in_counts(df: pd.DataFrame) -> pd.DataFrame:
    """Count first/second serves that landed in per player."""
    serves_in = df[
        df["Type"].isin(["first_serve", "second_serve"]) & (df["Result"] == "In")
    ]
    if serves_in.empty:
        return pd.DataFrame(columns=["First Serve In", "Second Serve In"]).rename_axis("Player")

    grouped = (
        serves_in.groupby(["Player", "Type"])["Point"]
        .size()
        .reset_index(name="in_count")
    )

    counts = grouped.pivot(index="Player", columns="Type", values="in_count").fillna(0)
    first_in = counts.get("first_serve", pd.Series(0, index=counts.index)).astype(int)
    second_in = counts.get("second_serve", pd.Series(0, index=counts.index)).astype(int)

    return pd.DataFrame(
        {
            "First Serve In": first_in,
            "Second Serve In": second_in,
        }
    )


def calculate_overall_serve_percentages(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate overall in-play serve percentages per player."""
    serves = df[
        df["Type"].isin(["first_serve", "second_serve"])
        & df["Result"].isin(["In", "Out", "Net"])
    ]
    if serves.empty:
        return pd.DataFrame(columns=["Overall First Serve %", "Overall Second Serve %"]).rename_axis(
            "Player"
        )

    grouped = (
        serves.groupby(["Player", "Type"])["Result"]
        .agg(attempts="count", wins=lambda x: (x == "In").sum())
        .reset_index()
    )

    attempts = grouped.pivot(index="Player", columns="Type", values="attempts").fillna(0)
    wins = grouped.pivot(index="Player", columns="Type", values="wins").fillna(0)

    pct_first = (wins.get("first_serve", 0) / attempts.get("first_serve", 1) * 100).replace(
        [float("inf")], 0
    ).fillna(0)
    pct_second = (wins.get("second_serve", 0) / attempts.get("second_serve", 1) * 100).replace(
        [float("inf")], 0
    ).fillna(0)

    return pd.DataFrame(
        {
            "Overall First Serve %": pct_first,
            "Overall Second Serve %": pct_second,
        }
    )


def summarize_from_stats(source, file_name: Optional[str] = None) -> pd.DataFrame:
    """Build a serve summary from the aggregated Stats and Settings sheets."""
    stats = _read_excel_sheet(source, "Stats", file_name=file_name)
    settings = _read_excel_sheet(source, "Settings", file_name=file_name)

    if "Stat Name" not in stats.columns:
        raise DataValidationError("Stats sheet is missing the 'Stat Name' column.")

    host_name = settings.get("Host Team", pd.Series(dtype=object)).dropna()
    guest_name = settings.get("Guest Team", pd.Series(dtype=object)).dropna()
    host_label = host_name.iloc[0] if not host_name.empty else "Host"
    guest_label = guest_name.iloc[0] if not guest_name.empty else "Guest"

    host_cols = [col for col in stats.columns if str(col).startswith("Host Set")]
    guest_cols = [col for col in stats.columns if str(col).startswith("Guest Set")]

    def sum_stat(stat_name: str, cols: list[str]) -> float:
        row = stats.loc[stats["Stat Name"] == stat_name, cols]
        if row.empty:
            return 0.0
        return float(row.fillna(0).sum(axis=1).iloc[0])

    def build_player_summary(cols: list[str]) -> dict[str, float]:
        first_serves = sum_stat("1st Serves", cols)
        first_in = sum_stat("1st Serves In", cols)
        first_won = sum_stat("1st Serves Won", cols)
        second_serves = sum_stat("2nd Serves", cols)
        second_in = sum_stat("2nd Serves In", cols)
        second_won = sum_stat("2nd Serves Won", cols)

        def pct(numerator: float, denominator: float) -> float:
            return 0.0 if denominator == 0 else (numerator / denominator * 100)

        return {
            "Overall First Serve %": pct(first_in, first_serves),
            "Overall Second Serve %": pct(second_in, second_serves),
            "First Serve In": first_in,
            "Second Serve In": second_in,
            "First Serve Attempts": first_serves,
            "Second Serve Attempts": second_serves,
            "First Serve Wins": first_won,
            "Second Serve Wins": second_won,
            "First Serve Win %": pct(first_won, first_serves),
            "Second Serve Win %": pct(second_won, second_serves),
        }

    summary = pd.DataFrame(
        [build_player_summary(host_cols), build_player_summary(guest_cols)],
        index=[host_label, guest_label],
    )
    summary.index.name = "Player"
    return summary


def aggregate_season_summaries(summaries: list[pd.DataFrame]) -> pd.DataFrame:
    """Aggregate per-match summaries into a season summary for all players."""
    if not summaries:
        return pd.DataFrame().rename_axis("Player")

    def pct(numerator: float, denominator: float) -> float:
        return 0.0 if denominator == 0 else (numerator / denominator * 100)

    rows = []
    all_players = sorted({player for df in summaries for player in df.index})
    for player in all_players:
        player_rows = [df.loc[player] for df in summaries if player in df.index]
        combined = pd.DataFrame(player_rows)

        def sum_col(name: str) -> float:
            if name not in combined.columns:
                return 0.0
            return float(combined[name].fillna(0).sum())

        first_attempts = sum_col("First Serve Attempts")
        second_attempts = sum_col("Second Serve Attempts")
        first_in = sum_col("First Serve In")
        second_in = sum_col("Second Serve In")
        first_wins = sum_col("First Serve Wins")
        second_wins = sum_col("Second Serve Wins")

        rows.append(
            {
                "Player": player,
                "Overall First Serve %": pct(first_in, first_attempts),
                "Overall Second Serve %": pct(second_in, second_attempts),
                "First Serve In": first_in,
                "Second Serve In": second_in,
                "First Serve Attempts": first_attempts,
                "Second Serve Attempts": second_attempts,
                "First Serve Wins": first_wins,
                "Second Serve Wins": second_wins,
                "First Serve Win %": pct(first_wins, first_attempts),
                "Second Serve Win %": pct(second_wins, second_attempts),
            }
        )

    summary = pd.DataFrame(rows).set_index("Player")
    summary.index.name = "Player"
    return summary


def summarize_all(df: pd.DataFrame) -> pd.DataFrame:
    """Compute the full serve summary table for each player."""
    serve_win = build_serve_win_data(df)
    serve_win_stats = calculate_serve_win_percentages(serve_win)
    serve_win_counts = calculate_serve_win_counts(serve_win)
    serve_attempts = calculate_serve_attempts(df)
    serve_in_counts = calculate_serve_in_counts(df)
    overall = calculate_overall_serve_percentages(df)
    summary = (
        overall
        .join(serve_win_stats, how="outer")
        .join(serve_win_counts, how="outer")
        .join(serve_attempts, how="outer")
        .join(serve_in_counts, how="outer")
        .fillna(0)
    )
    summary.index.name = "Player"
    return summary


def export_summary_bytes(df: pd.DataFrame, file_type: str = "csv") -> tuple[bytes, str]:
    """Serialize the summary dataframe to CSV/XLSX bytes and filename."""
    file_type = file_type.lower()
    if file_type not in {"csv", "xlsx"}:
        file_type = "csv"

    buffer = io.BytesIO()
    if file_type == "xlsx":
        df.to_excel(buffer, index=True)
        filename = "serve_summary.xlsx"
    else:
        df.to_csv(buffer, index=True)
        filename = "serve_summary.csv"
    buffer.seek(0)
    return buffer.getvalue(), filename
