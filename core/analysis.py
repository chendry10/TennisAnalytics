from __future__ import annotations

import logging
from typing import Dict, Iterable, Optional
import io
from pathlib import Path
import pandas as pd

pd.set_option("future.no_silent_downcasting", True)

from .errors import DataLoadError, DataValidationError
from .metrics import POINT_LENGTH_BUCKETS, SUMMARY_COUNT_KEYS, SUMMARY_RATE_COMPONENTS

logger = logging.getLogger(__name__)

POINT_FACT_COLUMNS = [
    "Point",
    "Server",
    "Serve_Type",
    "Serve_Result",
    "First_Serve_Attempted",
    "First_Serve_In",
    "Second_Serve_Attempted",
    "Second_Serve_In",
    "Double_Fault",
    "Returner",
    "Return_Type",
    "Return_Result",
    "Return_In",
    "Winner",
    "Server_Won",
    "Returner_Won",
    "Serve_Plus_One_Attempted",
    "Serve_Plus_One_Result",
    "Serve_Plus_One_In",
    "Serve_Plus_One_Won",
    "Return_Plus_One_Attempted",
    "Return_Plus_One_Result",
    "Return_Plus_One_In",
    "Return_Plus_One_Won",
    "ShotCount",
    "LengthBucket",
]

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
    if {"First Return In", "First Return Attempts"}.issubset(grouped.columns):
        grouped["First Return In %"] = pct(grouped["First Return In"], grouped["First Return Attempts"])
    if {"Second Return In", "Second Return Attempts"}.issubset(grouped.columns):
        grouped["Second Return In %"] = pct(grouped["Second Return In"], grouped["Second Return Attempts"])
    if {"First Return Wins", "First Return Attempts"}.issubset(grouped.columns):
        grouped["First Return Win %"] = pct(grouped["First Return Wins"], grouped["First Return Attempts"])
    if {"Second Return Wins", "Second Return Attempts"}.issubset(grouped.columns):
        grouped["Second Return Win %"] = pct(grouped["Second Return Wins"], grouped["Second Return Attempts"])
    if {"Double Faults", "Second Serve Attempts"}.issubset(grouped.columns):
        grouped["Double Fault Rate"] = pct(grouped["Double Faults"], grouped["Second Serve Attempts"])
    if {"Serve +1 In", "Serve +1 Attempts"}.issubset(grouped.columns):
        grouped["Serve +1 In %"] = pct(grouped["Serve +1 In"], grouped["Serve +1 Attempts"])
    if {"Serve +1 Wins", "Serve +1 Attempts"}.issubset(grouped.columns):
        grouped["Serve +1 Win %"] = pct(grouped["Serve +1 Wins"], grouped["Serve +1 Attempts"])
    if {"Return +1 In", "Return +1 Attempts"}.issubset(grouped.columns):
        grouped["Return +1 In %"] = pct(grouped["Return +1 In"], grouped["Return +1 Attempts"])
    if {"Return +1 Wins", "Return +1 Attempts"}.issubset(grouped.columns):
        grouped["Return +1 Win %"] = pct(grouped["Return +1 Wins"], grouped["Return +1 Attempts"])

    # Recalculate point-length win %
    for bucket in ["0-4_shots", "5-10_shots", "11plus_shots"]:
        wins_col = f"{bucket}_Wins"
        total_col = f"{bucket}_Total"
        pct_col = f"{bucket}_Win%"
        if {wins_col, total_col}.issubset(grouped.columns):
            grouped[pct_col] = pct(grouped[wins_col], grouped[total_col])

    grouped.index.name = "Player"
    return grouped


try:
    import python_calamine as _calamine  # noqa: F401
    _HAS_CALAMINE = True
except ImportError:
    _HAS_CALAMINE = False


def excel_engine(file_name: Optional[str]) -> Optional[str]:
    """Return the Excel engine name to use for a given filename (if any)."""
    if _suffix_from_name(file_name) == ".xls":
        return "xlrd"
    if _HAS_CALAMINE:
        return "calamine"
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
            engine = excel_engine(str(source))
            with pd.ExcelFile(source, engine=engine) as xls:
                return list(xls.sheet_names)
        data = _to_bytes(source)
        if data is None:
            return []
        engine = excel_engine(file_name)
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
            engine = excel_engine(str(source))
            return list(pd.read_excel(source, sheet_name=sheet, nrows=0, engine=engine).columns)
        data = _to_bytes(source)
        if data is None:
            return []
        if _is_csv(file_name):
            return list(pd.read_csv(io.BytesIO(data), nrows=0).columns)
        engine = excel_engine(file_name)
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


def validate_and_rename(df: pd.DataFrame, column_map: Optional[Dict[str, str]] = None) -> pd.DataFrame:
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

    df = df[REQUIRED_COLUMNS].copy()
    df["Result"] = df["Result"].astype(str).str.strip()
    df["Type"] = df["Type"].astype(str).str.strip()
    df["Player"] = df["Player"].astype(str).str.strip()
    return df


def _read_dataframe(source, sheet: Optional[str | int] = None, file_name: Optional[str] = None) -> pd.DataFrame:
    """Load a dataframe from a path or in-memory bytes payload."""
    try:
        if isinstance(source, (str, Path)):
            if _is_csv(str(source)):
                return pd.read_csv(source)
            engine = excel_engine(str(source))
            return pd.read_excel(source, sheet_name=sheet, engine=engine)

        data = _to_bytes(source)
        if data is None:
            raise DataLoadError("No data was provided.")

        if _is_csv(file_name):
            return pd.read_csv(io.BytesIO(data))

        try:
            engine = excel_engine(file_name)
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
            engine = excel_engine(str(source))
            return pd.read_excel(source, sheet_name=sheet_name, engine=engine)

        data = _to_bytes(source)
        if data is None:
            raise DataLoadError("No data was provided.")

        engine = excel_engine(file_name)
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, engine=engine)
    except DataLoadError:
        raise
    except Exception as exc:
        raise DataLoadError("Oops! This doesn't look like a valid Excel file.") from exc


def load_df(source, sheet: Optional[str | int] = None, column_map: Optional[Dict[str, str]] = None,
            file_name: Optional[str] = None) -> pd.DataFrame:
    """Load data and return a cleaned dataframe with required columns."""
    df = _read_dataframe(source, sheet=sheet, file_name=file_name)
    df = validate_and_rename(df, column_map=column_map)
    return df.dropna(subset=["Point"])


def _empty_point_facts() -> pd.DataFrame:
    return pd.DataFrame(columns=POINT_FACT_COLUMNS)


def _safe_pct(numerator, denominator):
    denominator = denominator.replace(0, pd.NA) if hasattr(denominator, "replace") else denominator
    return (numerator / denominator * 100).replace([float("inf")], 0).fillna(0)


def _ensure_point_facts(data: pd.DataFrame) -> pd.DataFrame:
    if set(POINT_FACT_COLUMNS).issubset(data.columns):
        return data.copy()
    return build_point_facts(data)


def _infer_point_winner(point_rows: pd.DataFrame, match_players: list[str]) -> str | None:
    if point_rows.empty:
        return None

    last_row = point_rows.iloc[-1]
    last_player = last_row["Player"]
    result = str(last_row["Result"]).strip()
    point_players = [player for player in dict.fromkeys(point_rows["Player"].tolist()) if player]

    if result == "In":
        return last_player
    if result in {"Out", "Net"}:
        opponents = [player for player in point_players if player != last_player]
        if not opponents:
            opponents = [player for player in match_players if player != last_player]
        return opponents[0] if opponents else None
    return None


def _find_transition_row(
    point_rows: pd.DataFrame,
    player: str | None,
    start_after: int | None,
    explicit_type: str,
):
    if player is None or start_after is None:
        return None

    explicit = point_rows[
        (point_rows.index > start_after)
        & (point_rows["Player"] == player)
        & (point_rows["Type"] == explicit_type)
    ]
    if not explicit.empty:
        return explicit.iloc[0]

    non_start_types = {"first_serve", "second_serve", "first_return", "second_return"}
    derived = point_rows[
        (point_rows.index > start_after)
        & (point_rows["Player"] == player)
        & ~point_rows["Type"].isin(non_start_types)
    ]
    if derived.empty:
        return None
    return derived.iloc[0]


def build_point_facts(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize raw shot rows into one fact row per point."""
    if df.empty:
        return _empty_point_facts()

    ordered = df.copy()
    ordered["__sequence__"] = range(len(ordered))
    ordered["__shot_order__"] = pd.to_numeric(ordered["Shot"], errors="coerce")
    match_players = [player for player in dict.fromkeys(ordered["Player"].tolist()) if player]

    rows: list[dict[str, object]] = []
    for point, point_rows in ordered.groupby("Point", sort=False):
        point_rows = point_rows.sort_values(
            by=["__shot_order__", "__sequence__"],
            kind="stable",
            na_position="last",
        ).reset_index(drop=True)

        serve_rows = point_rows[point_rows["Type"].isin(["first_serve", "second_serve"])]
        return_rows = point_rows[point_rows["Type"].isin(["first_return", "second_return"])]

        first_serve_row = serve_rows[serve_rows["Type"] == "first_serve"].iloc[0] if not serve_rows[serve_rows["Type"] == "first_serve"].empty else None
        second_serve_row = serve_rows[serve_rows["Type"] == "second_serve"].iloc[0] if not serve_rows[serve_rows["Type"] == "second_serve"].empty else None
        return_row = return_rows.iloc[-1] if not return_rows.empty else None

        server = None
        if first_serve_row is not None:
            server = first_serve_row["Player"]
        elif second_serve_row is not None:
            server = second_serve_row["Player"]

        returner = return_row["Player"] if return_row is not None else None
        if returner is None and server is not None:
            opponents = [player for player in match_players if player != server]
            returner = opponents[0] if len(opponents) == 1 else None

        serve_type = None
        serve_result = None
        if second_serve_row is not None:
            serve_type = "second_serve"
            serve_result = second_serve_row["Result"]
        elif first_serve_row is not None:
            serve_type = "first_serve"
            serve_result = first_serve_row["Result"]

        winner = _infer_point_winner(point_rows, match_players)

        return_index = int(return_row.name) if return_row is not None else None
        serve_plus_one_row = _find_transition_row(point_rows, server, return_index, "serve_plus_one")
        serve_plus_one_index = int(serve_plus_one_row.name) if serve_plus_one_row is not None else None
        return_plus_one_row = _find_transition_row(point_rows, returner, serve_plus_one_index, "return_plus_one")

        first_serve_attempted = first_serve_row is not None
        first_serve_in = first_serve_attempted and str(first_serve_row["Result"]).strip() == "In"
        second_serve_attempted = second_serve_row is not None
        second_serve_in = second_serve_attempted and str(second_serve_row["Result"]).strip() == "In"

        rows.append(
            {
                "Point": point,
                "Server": server,
                "Serve_Type": serve_type,
                "Serve_Result": serve_result,
                "First_Serve_Attempted": bool(first_serve_attempted),
                "First_Serve_In": bool(first_serve_in),
                "Second_Serve_Attempted": bool(second_serve_attempted),
                "Second_Serve_In": bool(second_serve_in),
                "Double_Fault": bool(second_serve_attempted and not second_serve_in),
                "Returner": returner,
                "Return_Type": return_row["Type"] if return_row is not None else None,
                "Return_Result": return_row["Result"] if return_row is not None else None,
                "Return_In": bool(return_row is not None and str(return_row["Result"]).strip() == "In"),
                "Winner": winner,
                "Server_Won": bool(server is not None and winner == server),
                "Returner_Won": bool(returner is not None and winner == returner),
                "Serve_Plus_One_Attempted": bool(serve_plus_one_row is not None),
                "Serve_Plus_One_Result": serve_plus_one_row["Result"] if serve_plus_one_row is not None else None,
                "Serve_Plus_One_In": bool(serve_plus_one_row is not None and str(serve_plus_one_row["Result"]).strip() == "In"),
                "Serve_Plus_One_Won": bool(serve_plus_one_row is not None and winner == server),
                "Return_Plus_One_Attempted": bool(return_plus_one_row is not None),
                "Return_Plus_One_Result": return_plus_one_row["Result"] if return_plus_one_row is not None else None,
                "Return_Plus_One_In": bool(return_plus_one_row is not None and str(return_plus_one_row["Result"]).strip() == "In"),
                "Return_Plus_One_Won": bool(return_plus_one_row is not None and winner == returner),
                "ShotCount": int(len(point_rows)),
                "LengthBucket": _length_bucket(int(len(point_rows))),
            }
        )

    point_facts = pd.DataFrame(rows, columns=POINT_FACT_COLUMNS)
    for column in [
        "First_Serve_Attempted",
        "First_Serve_In",
        "Second_Serve_Attempted",
        "Second_Serve_In",
        "Double_Fault",
        "Return_In",
        "Server_Won",
        "Returner_Won",
        "Serve_Plus_One_Attempted",
        "Serve_Plus_One_In",
        "Serve_Plus_One_Won",
        "Return_Plus_One_Attempted",
        "Return_Plus_One_In",
        "Return_Plus_One_Won",
    ]:
        point_facts[column] = point_facts[column].astype(bool)
    return point_facts


def get_point_servers(df: pd.DataFrame) -> pd.DataFrame:
    """Return one row per point with the server and serve type."""
    point_facts = _ensure_point_facts(df)
    return point_facts[["Point", "Server", "Serve_Type"]].dropna(subset=["Server"])


def get_point_winners(df: pd.DataFrame) -> pd.DataFrame:
    """Infer the winner for each point based on the last shot result."""
    point_facts = _ensure_point_facts(df)
    return point_facts[["Point", "Winner"]].dropna(subset=["Winner"])


def build_serve_win_data(df: pd.DataFrame) -> pd.DataFrame:
    """Combine servers and winners to compute serve-win outcomes."""
    if {"Point", "Server", "Serve_Type", "Winner"}.issubset(df.columns):
        return df[["Point", "Server", "Serve_Type", "Winner"]].dropna(subset=["Server", "Serve_Type"]).copy()
    point_facts = _ensure_point_facts(df)
    return point_facts[["Point", "Server", "Serve_Type", "Winner"]].dropna(subset=["Server", "Serve_Type"])


def calculate_serve_win_percentages(serve_win: pd.DataFrame) -> pd.DataFrame:
    """Calculate first/second serve win rates and attempt counts per player."""
    point_facts = _ensure_point_facts(serve_win)
    serve_win = build_serve_win_data(point_facts)
    serve_attempts = calculate_serve_attempts(point_facts)
    if serve_win.empty and serve_attempts.empty:
        return pd.DataFrame(
            columns=[
                "First Serve Win %",
                "Second Serve Win %",
            ]
        ).rename_axis("Player")

    if serve_win.empty:
        wins_pivot = pd.DataFrame(index=serve_attempts.index)
    else:
        serve_win = serve_win.copy()
        serve_win["Server_Won"] = serve_win["Server"] == serve_win["Winner"]
        grouped = (
            serve_win.groupby(["Server", "Serve_Type"])
            .agg(wins=("Server_Won", "sum"))
            .reset_index()
        )
        wins_pivot = grouped.pivot(index="Server", columns="Serve_Type", values="wins").fillna(0)

    first_attempts = serve_attempts.get("First Serve Attempts", pd.Series(0, index=serve_attempts.index))
    second_attempts = serve_attempts.get("Second Serve Attempts", pd.Series(0, index=serve_attempts.index))
    first_wins = wins_pivot.get("first_serve", pd.Series(0, index=wins_pivot.index))
    second_wins = wins_pivot.get("second_serve", pd.Series(0, index=wins_pivot.index))
    player_index = first_attempts.index.union(second_attempts.index).union(first_wins.index).union(second_wins.index)
    first_attempts = first_attempts.reindex(player_index, fill_value=0)
    second_attempts = second_attempts.reindex(player_index, fill_value=0)
    first_wins = first_wins.reindex(player_index, fill_value=0)
    second_wins = second_wins.reindex(player_index, fill_value=0)

    result = pd.DataFrame(
        {
            "First Serve Win %": _safe_pct(first_wins, first_attempts),
            "Second Serve Win %": _safe_pct(second_wins, second_attempts),
        },
        index=player_index,
    ).rename_axis("Player")

    return result


def calculate_serve_win_counts(serve_win: pd.DataFrame) -> pd.DataFrame:
    """Count first/second serve wins per player based on point outcomes."""
    serve_win = build_serve_win_data(serve_win)
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
    point_facts = _ensure_point_facts(df)
    point_facts = point_facts.dropna(subset=["Server"])
    if point_facts.empty:
        return pd.DataFrame(columns=["First Serve Attempts", "Second Serve Attempts"]).rename_axis(
            "Player"
        )

    grouped = point_facts.groupby("Server", sort=False).agg(
        **{
            "First Serve Attempts": ("First_Serve_Attempted", "sum"),
            "Second Serve Attempts": ("Second_Serve_Attempted", "sum"),
        }
    )

    return grouped.astype(int).rename_axis("Player")


def calculate_serve_in_counts(df: pd.DataFrame) -> pd.DataFrame:
    """Count first/second serves that landed in per player."""
    point_facts = _ensure_point_facts(df)
    point_facts = point_facts.dropna(subset=["Server"])
    if point_facts.empty:
        return pd.DataFrame(columns=["First Serve In", "Second Serve In"]).rename_axis("Player")

    grouped = point_facts.groupby("Server", sort=False).agg(
        **{
            "First Serve In": ("First_Serve_In", "sum"),
            "Second Serve In": ("Second_Serve_In", "sum"),
        }
    )

    return grouped.astype(int).rename_axis("Player")


def calculate_overall_serve_percentages(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate overall in-play serve percentages per player."""
    attempts = calculate_serve_attempts(df)
    in_counts = calculate_serve_in_counts(df)
    if attempts.empty and in_counts.empty:
        return pd.DataFrame(columns=["Overall First Serve %", "Overall Second Serve %"]).rename_axis(
            "Player"
        )

    combined = attempts.join(in_counts, how="outer").fillna(0)
    first_in = combined.get("First Serve In", pd.Series(0, index=combined.index))
    first_attempts = combined.get("First Serve Attempts", pd.Series(0, index=combined.index))
    second_in = combined.get("Second Serve In", pd.Series(0, index=combined.index))
    second_attempts = combined.get("Second Serve Attempts", pd.Series(0, index=combined.index))

    return pd.DataFrame(
        {
            "Overall First Serve %": _safe_pct(first_in, first_attempts),
            "Overall Second Serve %": _safe_pct(second_in, second_attempts),
        }
    )


def calculate_double_fault_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Count double faults and compute their rate per player."""
    point_facts = _ensure_point_facts(df)
    point_facts = point_facts.dropna(subset=["Server"])
    if point_facts.empty:
        return pd.DataFrame(columns=["Double Faults", "Double Fault Rate"]).rename_axis("Player")

    grouped = point_facts.groupby("Server", sort=False).agg(
        **{
            "Double Faults": ("Double_Fault", "sum"),
            "Second Serve Attempts": ("Second_Serve_Attempted", "sum"),
        }
    )
    grouped["Double Fault Rate"] = _safe_pct(
        grouped["Double Faults"], grouped["Second Serve Attempts"]
    )
    return grouped[["Double Faults", "Double Fault Rate"]].rename_axis("Player")


def calculate_plus_one_stats(df: pd.DataFrame, phase: str) -> pd.DataFrame:
    """Return serve +1 or return +1 counts and rates per player."""
    point_facts = _ensure_point_facts(df)

    if phase == "serve":
        player_col = "Server"
        attempted_col = "Serve_Plus_One_Attempted"
        in_col = "Serve_Plus_One_In"
        won_col = "Serve_Plus_One_Won"
        prefix = "Serve +1"
    elif phase == "return":
        player_col = "Returner"
        attempted_col = "Return_Plus_One_Attempted"
        in_col = "Return_Plus_One_In"
        won_col = "Return_Plus_One_Won"
        prefix = "Return +1"
    else:
        raise ValueError(f"Unsupported plus-one phase: {phase}")

    point_facts = point_facts.dropna(subset=[player_col])
    if point_facts.empty:
        return pd.DataFrame(
            columns=[
                f"{prefix} Attempts",
                f"{prefix} In",
                f"{prefix} In %",
                f"{prefix} Wins",
                f"{prefix} Win %",
            ]
        ).rename_axis("Player")

    grouped = point_facts.groupby(player_col, sort=False).agg(
        **{
            f"{prefix} Attempts": (attempted_col, "sum"),
            f"{prefix} In": (in_col, "sum"),
            f"{prefix} Wins": (won_col, "sum"),
        }
    )
    grouped[f"{prefix} In %"] = _safe_pct(grouped[f"{prefix} In"], grouped[f"{prefix} Attempts"])
    grouped[f"{prefix} Win %"] = _safe_pct(grouped[f"{prefix} Wins"], grouped[f"{prefix} Attempts"])
    return grouped.rename_axis("Player")


def summarize_from_stats(
    source=None,
    file_name: Optional[str] = None,
    *,
    stats_df: Optional[pd.DataFrame] = None,
    settings_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """Build a serve summary from the aggregated Stats and Settings sheets.

    Pass *stats_df* and *settings_df* directly to skip re-reading the file.
    """
    stats = stats_df if stats_df is not None else _read_excel_sheet(source, "Stats", file_name=file_name)
    try:
        settings = settings_df if settings_df is not None else _read_excel_sheet(source, "Settings", file_name=file_name)
    except (DataLoadError, Exception):
        settings = pd.DataFrame()

    if "Stat Name" not in stats.columns:
        logger.warning("Stats sheet is missing 'Stat Name' column in %s", file_name)
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

    rows = []
    all_players = sorted({player for df in summaries for player in df.index})
    for player in all_players:
        player_rows = [df.loc[player] for df in summaries if player in df.index]
        combined = pd.DataFrame(player_rows)

        def sum_col(name: str) -> float:
            if name not in combined.columns:
                return 0.0
            return float(combined[name].fillna(0).sum())

        row = {"Player": player}
        for metric in SUMMARY_COUNT_KEYS:
            row[metric] = sum_col(metric)
        for metric, (numerator, denominator) in SUMMARY_RATE_COMPONENTS.items():
            row[metric] = 0.0 if row.get(denominator, 0.0) == 0 else (row.get(numerator, 0.0) / row.get(denominator, 0.0) * 100)

        rows.append(row)

    summary = pd.DataFrame(rows).set_index("Player")
    summary.index.name = "Player"
    return summary


def summarize_all(df: pd.DataFrame) -> pd.DataFrame:
    """Compute the full serve and return summary table for each player."""
    point_facts = build_point_facts(df)
    serve_win_stats = calculate_serve_win_percentages(point_facts)
    serve_win_counts = calculate_serve_win_counts(point_facts)
    serve_attempts = calculate_serve_attempts(point_facts)
    serve_in_counts = calculate_serve_in_counts(point_facts)
    overall = calculate_overall_serve_percentages(point_facts)
    double_fault_stats = calculate_double_fault_stats(point_facts)

    return_in_counts = calculate_return_in_counts(point_facts)
    return_attempts = calculate_return_attempts(point_facts)
    return_pcts = calculate_return_percentages(point_facts)
    return_win_counts = calculate_return_win_counts(point_facts)
    return_win_pcts = calculate_return_win_percentages(point_facts)
    serve_plus_one = calculate_plus_one_stats(point_facts, "serve")
    return_plus_one = calculate_plus_one_stats(point_facts, "return")
    point_length = calculate_point_length_outcomes(point_facts)

    summary = (
        overall
        .join(serve_win_stats, how="outer")
        .join(serve_win_counts, how="outer")
        .join(serve_attempts, how="outer")
        .join(serve_in_counts, how="outer")
        .join(double_fault_stats, how="outer")
        .join(return_in_counts, how="outer")
        .join(return_attempts, how="outer")
        .join(return_pcts, how="outer")
        .join(return_win_counts, how="outer")
        .join(return_win_pcts, how="outer")
        .join(serve_plus_one, how="outer")
        .join(return_plus_one, how="outer")
        .join(point_length, how="outer")
        .fillna(0)
    )
    summary.index.name = "Player"
    return summary


def get_point_shot_counts(df: pd.DataFrame) -> pd.DataFrame:
    """Return one row per point with the total number of shots in that point."""
    counts = df.groupby("Point", sort=False).size().reset_index(name="ShotCount")
    return counts


def _length_bucket(shot_count: int) -> str:
    """Classify a shot count into a rally-length bucket."""
    if shot_count <= 4:
        return "0-4 shots"
    elif shot_count <= 10:
        return "5-10 shots"
    else:
        return "11+ shots"


def calculate_point_length_outcomes(df: pd.DataFrame) -> pd.DataFrame:
    """Per-player W/L counts and win % for each rally-length bucket.

    Returns a DataFrame indexed by Player with columns for each bucket's
    wins, losses, total, and win %.
    """
    point_facts = _ensure_point_facts(df)
    if point_facts.empty:
        return pd.DataFrame(columns=[]).rename_axis("Player")

    point_info = point_facts[["Point", "LengthBucket", "Winner"]].copy()
    all_players = sorted(
        {
            player
            for column in ["Server", "Returner", "Winner"]
            for player in point_facts[column].dropna().tolist()
        }
    )
    rows = []
    for player in all_players:
        row: dict = {}
        for bucket in POINT_LENGTH_BUCKETS:
            bucket_points = point_info[point_info["LengthBucket"] == bucket]
            total = len(bucket_points)
            wins = int((bucket_points["Winner"] == player).sum())
            losses = total - wins
            win_pct = (wins / total * 100) if total > 0 else 0.0
            tag = bucket.replace(" ", "_").replace("+", "plus")
            row[f"{tag}_Wins"] = wins
            row[f"{tag}_Losses"] = losses
            row[f"{tag}_Total"] = total
            row[f"{tag}_Win%"] = win_pct
        rows.append(row)

    result = pd.DataFrame(rows, index=all_players)
    result.index.name = "Player"
    return result


def calculate_return_in_counts(df: pd.DataFrame) -> pd.DataFrame:
    """Count first/second returns that landed in per player."""
    point_facts = _ensure_point_facts(df)
    point_facts = point_facts.dropna(subset=["Returner", "Return_Type"])
    if point_facts.empty:
        return pd.DataFrame(columns=["First Return In", "Second Return In"]).rename_axis("Player")

    grouped = (
        point_facts.groupby(["Returner", "Return_Type"])["Return_In"]
        .sum()
        .reset_index(name="in_count")
    )
    counts = grouped.pivot(index="Returner", columns="Return_Type", values="in_count").fillna(0)
    first_in = counts.get("first_return", pd.Series(0, index=counts.index)).astype(int)
    second_in = counts.get("second_return", pd.Series(0, index=counts.index)).astype(int)

    return pd.DataFrame({"First Return In": first_in, "Second Return In": second_in}).rename_axis("Player")


def calculate_return_attempts(df: pd.DataFrame) -> pd.DataFrame:
    """Count first/second return attempts per player."""
    point_facts = _ensure_point_facts(df)
    point_facts = point_facts.dropna(subset=["Returner", "Return_Type"])
    if point_facts.empty:
        return pd.DataFrame(columns=["First Return Attempts", "Second Return Attempts"]).rename_axis("Player")

    grouped = (
        point_facts.groupby(["Returner", "Return_Type"])["Point"]
        .size()
        .reset_index(name="attempts")
    )
    attempts = grouped.pivot(index="Returner", columns="Return_Type", values="attempts").fillna(0)
    first_att = attempts.get("first_return", pd.Series(0, index=attempts.index)).astype(int)
    second_att = attempts.get("second_return", pd.Series(0, index=attempts.index)).astype(int)

    return pd.DataFrame({"First Return Attempts": first_att, "Second Return Attempts": second_att}).rename_axis("Player")


def calculate_return_percentages(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate overall return-in percentages per player."""
    attempts = calculate_return_attempts(df)
    in_counts = calculate_return_in_counts(df)
    if attempts.empty and in_counts.empty:
        return pd.DataFrame(columns=["First Return In %", "Second Return In %"]).rename_axis("Player")

    combined = attempts.join(in_counts, how="outer").fillna(0)
    first_in = combined.get("First Return In", pd.Series(0, index=combined.index))
    first_attempts = combined.get("First Return Attempts", pd.Series(0, index=combined.index))
    second_in = combined.get("Second Return In", pd.Series(0, index=combined.index))
    second_attempts = combined.get("Second Return Attempts", pd.Series(0, index=combined.index))

    return pd.DataFrame({
        "First Return In %": _safe_pct(first_in, first_attempts),
        "Second Return In %": _safe_pct(second_in, second_attempts),
    }).rename_axis("Player")


def get_point_returners(df: pd.DataFrame) -> pd.DataFrame:
    """Return one row per point with the returner and return type."""
    point_facts = _ensure_point_facts(df)
    return point_facts[["Point", "Returner", "Return_Type"]].dropna(subset=["Returner"])


def calculate_return_win_counts(df: pd.DataFrame) -> pd.DataFrame:
    """Count first/second return wins per player based on point outcomes."""
    point_facts = _ensure_point_facts(df)
    merged = point_facts[["Returner", "Return_Type", "Returner_Won"]].dropna(subset=["Returner", "Return_Type"])
    if merged.empty:
        return pd.DataFrame(columns=["First Return Wins", "Second Return Wins"]).rename_axis("Player")

    grouped = (
        merged.groupby(["Returner", "Return_Type"])["Returner_Won"]
        .sum()
        .reset_index(name="wins")
    )
    wins = grouped.pivot(index="Returner", columns="Return_Type", values="wins").fillna(0)
    first_wins = wins.get("first_return", pd.Series(0, index=wins.index)).astype(int)
    second_wins = wins.get("second_return", pd.Series(0, index=wins.index)).astype(int)

    result = pd.DataFrame({"First Return Wins": first_wins, "Second Return Wins": second_wins})
    result.index.name = "Player"
    return result


def calculate_return_win_percentages(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate first/second return win rates per player."""
    point_facts = _ensure_point_facts(df)
    merged = point_facts[["Point", "Returner", "Return_Type", "Returner_Won"]].dropna(subset=["Returner", "Return_Type"])
    if merged.empty:
        return pd.DataFrame(columns=["First Return Win %", "Second Return Win %"]).rename_axis("Player")

    grouped = (
        merged.groupby(["Returner", "Return_Type"])
        .agg(attempts=("Point", "size"), wins=("Returner_Won", "sum"))
        .reset_index()
    )

    attempts_pivot = grouped.pivot(index="Returner", columns="Return_Type", values="attempts").fillna(0)
    wins_pivot = grouped.pivot(index="Returner", columns="Return_Type", values="wins").fillna(0)

    first_att = attempts_pivot.get("first_return", pd.Series(0, index=attempts_pivot.index))
    second_att = attempts_pivot.get("second_return", pd.Series(0, index=attempts_pivot.index))
    first_w = wins_pivot.get("first_return", pd.Series(0, index=wins_pivot.index))
    second_w = wins_pivot.get("second_return", pd.Series(0, index=wins_pivot.index))

    result = pd.DataFrame(
        {"First Return Win %": _safe_pct(first_w, first_att), "Second Return Win %": _safe_pct(second_w, second_att)},
        index=attempts_pivot.index,
    )
    result.index.name = "Player"
    return result


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
