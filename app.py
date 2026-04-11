import hashlib
import io
import logging
import re
from datetime import date, timedelta
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

from core.analysis import (
    excel_engine,
    validate_and_rename,
    aggregate_season_summaries,
    export_summary_bytes,
    get_excel_sheet_names,
    load_df,
    normalize_summary_players,
    summarize_from_stats,
    summarize_all,
)
from core.disk_cache import load_cache_entry, save_cache_entry
from core.errors import DataLoadError, DataValidationError
from core.metrics import (
    METRIC_DEFINITIONS,
    POINT_LENGTH_BUCKETS,
    RETURN_TABLE_KEYS,
    SERVE_TABLE_KEYS,
    TIMELINE_METRIC_KEYS,
    TRANSITION_TABLE_KEYS,
)

st.set_page_config(
    page_title="CourtSide Analytics",
    page_icon="🎾",
    layout="wide",
)

CSS = """
<style>
:root {
    --court-green: #0f3d2e;
    --court-bright: #18a66c;
    --accent: #d6ff3d;
    --card: #15201d;
    --ink: #ffffff;
    --muted: #d6e0dc;
}

html, body, [class*="css"] {
    font-family: "Inter", "Segoe UI", system-ui, sans-serif;
    background-color: #0b1210;
    color: var(--ink);
    font-size: 17px;
    line-height: 1.45;
}

[data-testid="stAppViewContainer"] {
    background-color: #0b1210;
    color: var(--ink);
}

[data-testid="stHeader"] {
    background-color: #0b1210;
}

[data-testid="stToolbarActions"],
[data-testid="stToolbarActions"] *,
[data-testid="stToolbarActions"] button,
[data-testid="stToolbarActions"] a,
[data-testid="stToolbarActions"] span,
[data-testid="stToolbarActions"] svg,
[data-testid="stToolbarActions"] path,
[data-testid="stToolbarActions"] img,
[data-testid="stToolbarActions"] [data-testid="stIconMaterial"],
#MainMenu button,
#MainMenu button *,
#MainMenu button svg,
#MainMenu button path {
    color: #f7fbf9 !important;
    fill: #f7fbf9 !important;
    stroke: #f7fbf9 !important;
    opacity: 1 !important;
}

[data-testid="stToolbarActions"] img,
[data-testid="stToolbarActions"] svg {
    filter: brightness(0) invert(1) !important;
}

[data-testid="stToolbarActions"] button,
[data-testid="stToolbarActions"] a,
#MainMenu button {
    background: transparent !important;
    background-color: transparent !important;
    border: 0 !important;
    box-shadow: none !important;
    outline: none !important;
}

[data-testid="stToolbarActions"] button:hover,
[data-testid="stToolbarActions"] a:hover,
[data-testid="stToolbarActions"] button:hover span,
[data-testid="stToolbarActions"] button:hover svg,
[data-testid="stToolbarActions"] button:hover path,
[data-testid="stToolbarActions"] a:hover svg,
[data-testid="stToolbarActions"] a:hover path,
#MainMenu button:hover,
#MainMenu button:focus,
#MainMenu button:active,
#MainMenu button:hover svg,
#MainMenu button:hover path,
#MainMenu button:focus svg,
#MainMenu button:focus path,
#MainMenu button:active svg,
#MainMenu button:active path {
    color: var(--accent) !important;
    fill: var(--accent) !important;
    stroke: var(--accent) !important;
}

#MainMenu,
#MainMenu button,
#MainMenu button:hover,
#MainMenu button:focus,
#MainMenu button:active {
    background: transparent !important;
    background-color: transparent !important;
    box-shadow: none !important;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

.hero {
    background: linear-gradient(120deg, #0f3d2e 0%, #0b1210 70%);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 18px;
    padding: 28px 32px;
    margin-bottom: 1.5rem;
    box-shadow: 0 12px 30px rgba(0,0,0,0.35);
}

.hero h1 {
    color: var(--accent);
    margin-bottom: 0.4rem;
    font-size: 2.35rem;
}

.hero p {
    color: var(--muted);
    margin: 0;
    font-size: 1.12rem;
    font-weight: 500;
}

.metric-card {
    background: var(--card);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 14px;
    padding: 14px 18px;
}

.stMetric label,
div[data-testid="stMetricLabel"] {
    color: #f2f7f5 !important;
    font-size: 1.02rem !important;
}

.stButton>button {
    background-color: var(--court-bright) !important;
    color: #08100d !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
}

.stDownloadButton>button {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    border-radius: 12px !important;
}

.stMetricValue,
div[data-testid="stMetricValue"] {
    color: #ffffff !important;
    font-size: 2rem !important;
    font-weight: 700 !important;
}

.stFileUploader button,
.stFileUploader [data-testid="stFileUploader"] button,
.stFileUploader .stButton>button {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
}

.stFileUploader [data-testid="stFileUploaderFileName"],
.stFileUploader [data-testid="stFileUploaderFileName"] span,
.stFileUploader [data-testid="stFileUploaderFileName"] small,
.stFileUploader [data-testid="stFileUploaderFileName"] svg,
.stFileUploader [data-testid="stFileUploaderFileName"] path,
.stFileUploader [data-testid="stFileUploaderFileName"] div,
.stFileUploader [data-testid="stFileUploaderDropzone"] svg,
.stFileUploader [data-testid="stFileUploaderDropzone"] path,
.stFileUploader [data-testid="stFileUploaderDropzone"] {
    color: var(--ink) !important;
    fill: var(--ink) !important;
}

.stFileUploader [data-testid="stFileUploaderFileName"] * {
    color: var(--ink) !important;
    fill: var(--ink) !important;
}

.stFileUploader small,
.stFileUploader p,
.stFileUploader span {
    color: var(--ink) !important;
    font-size: 1rem !important;
}

section[data-testid="stSidebar"] {
    background-color: #0f1715;
    border-right: 1px solid rgba(255,255,255,0.06);
}

section[data-testid="stSidebar"],
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] .stCaption,
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stRadio label,
section[data-testid="stSidebar"] .stFileUploader label {
    color: var(--ink) !important;
    font-size: 1.03rem !important;
    font-weight: 600 !important;
}

.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div,
.stTextInput input,
.stTextArea textarea,
.stFileUploader section,
.stDateInput input,
.stNumberInput input {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
    border-color: rgba(255,255,255,0.12) !important;
    font-size: 1.02rem !important;
}

div[role="listbox"],
ul[role="listbox"] {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
}

div[role="option"] {
    color: var(--ink) !important;
    font-size: 1.02rem !important;
}

div[role="menu"],
ul[role="menu"],
div[data-baseweb="menu"],
div[data-baseweb="menu"] ul,
div[data-baseweb="popover"],
div[data-baseweb="popover"] > div,
div[data-baseweb="popover"] [role="menuitem"],
div[data-baseweb="popover"] [role="menuitem"] span {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
}

div[data-baseweb="popover"] [role="menuitem"]:hover,
div[data-baseweb="popover"] [role="menuitem"]:focus {
    background-color: #141f1c !important;
}

[data-testid="stMainMenuPopover"],
[data-testid="stMainMenuPopover"] *,
[data-testid="stMainMenuPopover"] [role="menuitem"],
[data-testid="stMainMenuPopover"] [role="menuitem"] * {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
}

[data-testid="stMainMenuPopover"] [role="menuitem"]:hover,
[data-testid="stMainMenuPopover"] [role="menuitem"]:focus {
    background-color: #141f1c !important;
}

h1, h2, h3, h4,
.stMarkdown p,
.stMainBlockContainer label,
.stMainBlockContainer small,
.stMainBlockContainer span,
div[data-testid="stCaptionContainer"],
.stMainBlockContainer p {
    color: var(--ink) !important;
}

h2 {
    font-size: 1.9rem !important;
}

h3 {
    font-size: 1.45rem !important;
}

div[data-testid="stCaptionContainer"],
.stCaption,
small {
    color: #dce8e4 !important;
    font-size: 0.98rem !important;
}

section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label {
    color: #f5fbf8 !important;
}

.stDataFrame,
.stDataFrame * {
    color: #f5fbf8 !important;
    font-size: 1rem !important;
}

div[data-testid="stDataFrame"] th,
div[data-testid="stDataFrame"] td {
    color: #f5fbf8 !important;
}

.stAlert,
div[data-baseweb="notification"],
[data-testid="stNotification"] {
    color: #ffffff !important;
    font-size: 1rem !important;
}

.stRadio label,
.stCheckbox label,
.stSelectbox label,
.stMultiSelect label {
    font-size: 1.02rem !important;
    font-weight: 600 !important;
}

.stRadio [role="radiogroup"] label,
.stCheckbox [data-testid="stMarkdownContainer"] p {
    color: #eef6f3 !important;
}

/* Reset text color for Streamlit deploy dialog and modals */
[data-testid="stModal"],
[data-testid="stModal"] *,
div[data-baseweb="modal"],
div[data-baseweb="modal"] * {
    color: initial !important;
    fill: initial !important;
}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)

st.markdown(
    """
    <div class="hero">
        <h1>CourtSide Analytics</h1>
        <p>Upload one file or a full folder to instantly view serve performance, win rates, and consistency.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv")
EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm")
ANALYSIS_CACHE_VERSION = "2026-04-10-point-facts-v2"


def metric_label(metric_key: str) -> str:
    definition = METRIC_DEFINITIONS.get(metric_key)
    return definition.label if definition else metric_key


def metric_help(metric_key: str) -> str | None:
    definition = METRIC_DEFINITIONS.get(metric_key)
    return definition.description if definition else None


def metric_format_pattern(metric_key: str) -> str:
    definition = METRIC_DEFINITIONS.get(metric_key)
    if definition and definition.kind == "percent":
        return "{:.1f}%"
    return "{:.0f}"


def metric_axis_title(metric_key: str) -> str:
    definition = METRIC_DEFINITIONS.get(metric_key)
    if definition and definition.kind == "percent":
        return "Percentage"
    return "Count"


def build_analysis_cache_key(file_bytes: bytes, sheet_for_file: str | int | None) -> str:
    file_hash = hashlib.sha256(file_bytes).hexdigest()
    sheet_key = "none" if sheet_for_file is None else str(sheet_for_file)
    raw_key = f"{ANALYSIS_CACHE_VERSION}:{file_hash}:{sheet_key}"
    return hashlib.sha256(raw_key.encode("utf-8")).hexdigest()


def preferred_sheet_index(sheet_names: list[str]) -> int:
    for index, sheet_name in enumerate(sheet_names):
        if str(sheet_name).strip().lower() == "shots":
            return index
    for index, sheet_name in enumerate(sheet_names):
        if str(sheet_name).strip().lower() == "stats":
            return index
    return 0


@st.cache_data(show_spinner=False)
def cached_excel_sheet_names(file_bytes: bytes, file_name: str) -> list[str]:
    return get_excel_sheet_names(file_bytes, file_name)


@st.cache_data(show_spinner=False)
def cached_file_summary(
    file_bytes: bytes,
    file_name: str,
    sheet_for_file: str | int | None,
) -> pd.DataFrame:
    cache_key = build_analysis_cache_key(file_bytes, sheet_for_file)
    cached_summary = load_cache_entry(cache_key)
    if isinstance(cached_summary, pd.DataFrame):
        return cached_summary

    base_name = file_name.replace("\\", "/").split("/")[-1]
    if base_name.startswith("~$"):
        return pd.DataFrame().rename_axis("Player")

    lower_name = file_name.lower()
    is_excel = lower_name.endswith(EXCEL_EXTENSIONS)

    summary = None
    if is_excel:
        engine = excel_engine(file_name)
        logger.info("Loading %s (engine=%s)", file_name, engine)
        with pd.ExcelFile(io.BytesIO(file_bytes), engine=engine) as xls:
            sheet_names = list(xls.sheet_names)

            shots_sheet = next(
                (name for name in sheet_names if str(name).strip().lower() == "shots"),
                sheet_for_file if sheet_for_file is not None else 0,
            )
            try:
                raw_df = validate_and_rename(xls.parse(shots_sheet))
                raw_df = raw_df.dropna(subset=["Point"])
            except Exception:
                raw_df = None

            if raw_df is not None and not raw_df.empty:
                logger.info("Using point-fact analysis for %s", file_name)
                summary = summarize_all(raw_df)
            elif "Stats" in sheet_names:
                logger.info("Using Stats sheet fallback for %s", file_name)
                stats_df = xls.parse("Stats")
                try:
                    settings_df = xls.parse("Settings")
                except Exception:
                    settings_df = None
                summary = summarize_from_stats(stats_df=stats_df, settings_df=settings_df)

    if summary is None or summary.empty:
        logger.warning("Falling back to full load_df for %s", file_name)
        df = load_df(
            file_bytes,
            sheet=sheet_for_file,
            column_map=None,
            file_name=file_name,
        )
        summary = summarize_all(df)

    summary = normalize_summary_players(summary)
    save_cache_entry(cache_key, summary)
    return summary


def parse_match_date_from_filename(file_name: str) -> date | None:
    normalized_name = file_name.replace("\\", "/")
    match = re.search(r"(\d{4}-\d{2}-\d{2})", normalized_name)
    if not match:
        return None
    try:
        return date.fromisoformat(match.group(1))
    except ValueError:
        return None

files_to_process = []
sheet_names_by_file = {}
default_sheet_by_file = {}

with st.sidebar:
    st.header("Upload Data")
    upload_mode = st.radio(
        "Choose input type",
        ["Single file", "Folder"],
        index=1,
    )

    if upload_mode == "Single file":
        uploaded = st.file_uploader(
            "Drag & Drop SwingVision File Here",
            type=["xlsx", "xls", "xlsm", "csv"],
            accept_multiple_files=False,
        )
        files_to_process = [uploaded] if uploaded else []
    else:
        uploaded = st.file_uploader(
            "Drag & Drop Folder with SwingVision Files Here",
            accept_multiple_files="directory",
        )
        files_to_process = uploaded if uploaded else []

    sheet_name = None
    column_map = None

    if files_to_process:
        is_excel_any = False
        seen_hashes = set()
        invalid_excel_files = []
        skipped_temp_files = 0
        skipped_metadata_files = 0
        validated_files = []

        for file in files_to_process:
            file_name = file.name
            normalized_name = file_name.replace("\\", "/")
            base_name = normalized_name.split("/")[-1]
            if base_name.startswith("~$"):
                skipped_temp_files += 1
                continue
            if base_name.lower() in {"desktop.ini", "thumbs.db", ".ds_store"} or base_name.startswith("."):
                skipped_metadata_files += 1
                continue

            file_bytes = file.getvalue()
            file_hash = hashlib.sha256(file_bytes).hexdigest()
            if file_hash in seen_hashes:
                st.error("Duplicate file detected. Please remove duplicates and try again.")
                st.stop()
            seen_hashes.add(file_hash)

            lower_name = file_name.lower()
            if not lower_name.endswith(SUPPORTED_EXTENSIONS):
                st.error(
                    f"Unsupported file type: {file_name}. "
                    "Only .xlsx/.xls/.xlsm/.csv files are allowed."
                )
                st.stop()

            is_excel = lower_name.endswith(EXCEL_EXTENSIONS)
            is_excel_any = is_excel_any or is_excel

            if is_excel:
                try:
                    sheet_names_by_file[file_name] = cached_excel_sheet_names(file_bytes, file_name)
                except DataLoadError:
                    invalid_excel_files.append(file_name)
                    continue

                if not sheet_names_by_file[file_name]:
                    st.error(f"No readable sheets found in {file_name}.")
                    st.stop()
                default_sheet_by_file[file_name] = preferred_sheet_index(sheet_names_by_file[file_name])

            validated_files.append(file)

        if skipped_temp_files:
            logger.info("Skipped %d temporary Excel lock file(s)", skipped_temp_files)
        if skipped_metadata_files:
            logger.info("Skipped %d metadata/system file(s)", skipped_metadata_files)

        if invalid_excel_files:
            preview = "\n".join(f"- {name}" for name in invalid_excel_files[:10])
            more = ""
            if len(invalid_excel_files) > 10:
                more = f"\n- ...and {len(invalid_excel_files) - 10} more"
            st.error(
                "Some uploaded Excel files are not valid and were blocked. "
                "Please remove or re-export these files:\n"
                f"{preview}{more}"
            )
            st.stop()

        files_to_process = validated_files
        if not files_to_process:
            st.error("No valid files found to analyze. Please upload at least one valid SwingVision file.")
            st.stop()

        if is_excel_any:
            example_file = files_to_process[0]
            example_names = sheet_names_by_file.get(example_file.name, [])
            if example_names:
                sheet_name = default_sheet_by_file.get(example_file.name, 0)
                st.caption(f"Using sheet: {example_names[sheet_name]}")

    st.markdown("---")
    if files_to_process:
        output_type = "xlsx"
        st.caption(f"Download format: {output_type.upper()}")
    else:
        output_type = "csv"

def render_grouped_bar_chart(
    summary: pd.DataFrame,
    metric_keys: list[str],
    title: str,
    color_sequence: list[str],
) -> None:
    keys = [key for key in metric_keys if key in summary.columns]
    if not keys:
        return

    long_df = summary.reset_index().melt(
        id_vars="Player",
        value_vars=keys,
        var_name="Metric",
        value_name="Value",
    )
    long_df["Metric"] = long_df["Metric"].map(metric_label)
    chart = px.bar(
        long_df,
        x="Player",
        y="Value",
        color="Metric",
        barmode="group",
        text="Value",
        color_discrete_sequence=color_sequence,
    )
    is_percent = all(
        METRIC_DEFINITIONS.get(key) and METRIC_DEFINITIONS[key].kind == "percent"
        for key in keys
    )
    chart.update_traces(
        texttemplate="%{text:.1f}%" if is_percent else "%{text:.0f}",
        textposition="outside",
    )
    chart.update_layout(
        template="plotly_dark",
        height=340,
        margin=dict(l=20, r=20, t=40, b=40),
        yaxis=dict(title="Percentage" if is_percent else "Count", range=[0, 100] if is_percent else None),
        xaxis=dict(title=None),
        legend_title_text="",
        title=title,
    )
    st.plotly_chart(chart, width="stretch")


def render_player_group_chart(
    summary: pd.DataFrame,
    metric_keys: list[str],
    title: str,
    category_order: list[str] | None = None,
) -> None:
    keys = [key for key in metric_keys if key in summary.columns]
    if not keys:
        return

    long_df = summary.reset_index().melt(
        id_vars="Player",
        value_vars=keys,
        var_name="Metric",
        value_name="Value",
    )
    long_df["Metric"] = long_df["Metric"].map(metric_label)
    chart = px.bar(
        long_df,
        x="Metric",
        y="Value",
        color="Player",
        barmode="group",
        text="Value",
        color_discrete_sequence=["#18a66c", "#d6ff3d", "#00c2ff", "#ff6f61"],
    )
    chart.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    xaxis = dict(title=None)
    if category_order:
        xaxis["categoryorder"] = "array"
        xaxis["categoryarray"] = category_order
    chart.update_layout(
        template="plotly_dark",
        height=340,
        margin=dict(l=20, r=20, t=40, b=40),
        yaxis=dict(range=[0, 100], title="Percentage"),
        xaxis=xaxis,
        legend_title_text="",
        title=title,
    )
    st.plotly_chart(chart, width="stretch")


def render_charts(summary: pd.DataFrame) -> None:
    render_grouped_bar_chart(
        summary,
        [
            "First Serve Attempts",
            "First Serve In",
            "Second Serve Attempts",
            "Second Serve In",
        ],
        title="Serve Volume",
        color_sequence=["#00c2ff", "#7ae9ff", "#18a66c", "#9be564"],
    )
    render_grouped_bar_chart(
        summary,
        ["First Serve Wins", "Second Serve Wins", "Double Faults"],
        title="Serve Outcomes",
        color_sequence=["#d6ff3d", "#5b8cff", "#ff6f61"],
    )
    render_grouped_bar_chart(
        summary,
        [
            "Overall First Serve %",
            "First Serve Win %",
            "Overall Second Serve %",
            "Second Serve Win %",
            "Double Fault Rate",
        ],
        title="Serve Rates",
        color_sequence=["#00c2ff", "#d6ff3d", "#18a66c", "#9be564", "#ff6f61"],
    )
    render_grouped_bar_chart(
        summary,
        [
            "First Return Attempts",
            "First Return In",
            "Second Return Attempts",
            "Second Return In",
        ],
        title="Return Volume",
        color_sequence=["#ff6f61", "#ff9f80", "#ffb347", "#ffd166"],
    )
    render_grouped_bar_chart(
        summary,
        ["First Return Wins", "Second Return Wins"],
        title="Return Wins",
        color_sequence=["#d1495b", "#f28482"],
    )
    render_grouped_bar_chart(
        summary,
        ["First Return In %", "First Return Win %", "Second Return In %", "Second Return Win %"],
        title="Return Rates",
        color_sequence=["#ff6f61", "#ffb347", "#ffd166", "#d1495b"],
    )

    if any(key in summary.columns for key in TRANSITION_TABLE_KEYS):
        render_grouped_bar_chart(
            summary,
            [
                "Serve +1 Attempts",
                "Serve +1 In",
                "Serve +1 Wins",
                "Return +1 Attempts",
                "Return +1 In",
                "Return +1 Wins",
            ],
            title="Transition Volume",
            color_sequence=["#18a66c", "#4ed49a", "#9be564", "#00c2ff", "#5fd8ff", "#5b8cff"],
        )
        render_grouped_bar_chart(
            summary,
            ["Serve +1 In %", "Serve +1 Win %", "Return +1 In %", "Return +1 Win %"],
            title="Transition Rates",
            color_sequence=["#18a66c", "#9be564", "#00c2ff", "#5b8cff"],
        )

    rally_keys = [f"{bucket.replace(' ', '_').replace('+', 'plus')}_Win%" for bucket in POINT_LENGTH_BUCKETS]
    render_player_group_chart(
        summary,
        rally_keys,
        title="Win % by Rally Length",
        category_order=[metric_label(key) for key in rally_keys],
    )


def format_timeline_axis_label(match_date: date | None, order: int) -> str:
    if match_date is not None:
        return match_date.isoformat()
    return f"Match {order + 1}"


def build_timeline_match_label(file_name: str, match_date: date | None, order: int) -> str:
    stem = Path(file_name).stem
    cleaned = re.sub(r"^SwingVision-match-", "", stem, flags=re.IGNORECASE)
    cleaned = cleaned.replace("_", " ")
    cleaned = re.sub(r"\bat (\d{2})\.(\d{2})(?:\.\d{2})?", r" \1:\2", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" -")
    if not cleaned:
        cleaned = f"Match {order + 1}"
    if len(cleaned) > 28:
        cleaned = cleaned[:25].rstrip() + "..."
    if match_date is None:
        return cleaned
    if cleaned.startswith(match_date.isoformat()):
        return cleaned
    return f"{match_date.isoformat()} | {cleaned}"


def render_timeline_view(
    selected_files: list[str],
    summaries_by_file: dict[str, pd.DataFrame],
    selected_players: list[str],
    parsed_dates_by_file: dict[str, date | None],
) -> None:
    st.subheader("Timeline")
    default_metrics = [
        key
        for key in ["First Serve Win %", "First Return Win %"]
        if key in TIMELINE_METRIC_KEYS
    ]
    timeline_metrics = st.multiselect(
        "Timeline metrics",
        options=TIMELINE_METRIC_KEYS,
        default=default_metrics or TIMELINE_METRIC_KEYS[:1],
        format_func=metric_label,
        help="Overlay one or more metrics across the selected matches.",
    )
    if not timeline_metrics:
        st.info("Select at least one metric to draw the timeline.")
        return

    annotate_matches = st.checkbox(
        "Show match annotations",
        value=len(selected_files) <= 8,
        help="Add vertical match callouts so the overlay stays tied to specific files.",
    )
    st.caption("Color identifies the player. Line style identifies the selected metric.")

    timeline_rows = []
    match_rows = []
    for order, file_name in enumerate(selected_files):
        summary = summaries_by_file[file_name]
        match_date = parsed_dates_by_file.get(file_name)
        axis_label = format_timeline_axis_label(match_date, order)
        match_label = build_timeline_match_label(file_name, match_date, order)
        match_rows.append(
            {
                "Match Order": order + 1,
                "Axis Label": axis_label,
                "Match Label": match_label,
                "File": file_name,
            }
        )
        for player in selected_players:
            if player not in summary.index:
                continue
            for metric_key in timeline_metrics:
                if metric_key not in summary.columns:
                    continue
                value = summary.loc[player, metric_key]
                if pd.isna(value):
                    continue
                timeline_rows.append(
                    {
                        "Match Order": order + 1,
                        "Axis Label": axis_label,
                        "Match Label": match_label,
                        "File": file_name,
                        "Match Date": match_date.isoformat() if match_date is not None else "Undated",
                        "Player": player,
                        "Metric": metric_label(metric_key),
                        "Metric Key": metric_key,
                        "Series": f"{player} | {metric_label(metric_key)}",
                        "Value": float(value),
                    }
                )

    if not timeline_rows:
        st.info("No timeline data is available for the current selection.")
        return

    timeline_df = pd.DataFrame(timeline_rows).sort_values(["Match Order", "Player", "Metric"])
    match_df = pd.DataFrame(match_rows).drop_duplicates(subset=["Match Order"])

    metric_kinds = {
        METRIC_DEFINITIONS[metric_key].kind
        for metric_key in timeline_df["Metric Key"].unique()
        if metric_key in METRIC_DEFINITIONS
    }
    percent_only = metric_kinds == {"percent"}

    line_chart = px.line(
        timeline_df,
        x="Match Order",
        y="Value",
        color="Player",
        line_dash="Metric",
        symbol="Metric",
        markers=True,
        hover_name="Series",
        custom_data=["Axis Label", "Match Label", "Match Date", "File"],
        color_discrete_sequence=["#18a66c", "#d6ff3d", "#00c2ff", "#ff6f61"],
        line_dash_sequence=["solid", "dot", "dash", "longdash", "dashdot"],
        symbol_sequence=["circle", "diamond", "square", "x", "triangle-up", "pentagon"],
    )
    line_chart.update_layout(
        template="plotly_dark",
        height=380,
        margin=dict(l=20, r=20, t=72 if annotate_matches else 20, b=70),
        yaxis=dict(title="Percent" if percent_only else "Metric Value", range=[0, 100] if percent_only else None),
        xaxis=dict(
            title=None,
            tickmode="array",
            tickvals=match_df["Match Order"].tolist(),
            ticktext=match_df["Axis Label"].tolist(),
            tickangle=-35,
        ),
        legend_title_text="",
    )
    if percent_only:
        line_chart.update_traces(
            hovertemplate="<b>%{hovertext}</b><br>%{customdata[0]}<br>%{customdata[1]}<br>Date: %{customdata[2]}<br>File: %{customdata[3]}<br>%{y:.1f}%<extra></extra>"
        )
    else:
        line_chart.update_traces(
            hovertemplate="<b>%{hovertext}</b><br>%{customdata[0]}<br>%{customdata[1]}<br>Date: %{customdata[2]}<br>File: %{customdata[3]}<br>%{y:.1f}<extra></extra>"
        )

    if annotate_matches:
        annotation_rows = match_df.tail(10) if len(match_df) > 10 else match_df
        if len(match_df) > 10:
            st.caption("Match annotations are limited to the latest 10 selected files to keep the chart readable.")
        for row in annotation_rows.to_dict("records"):
            line_chart.add_vline(
                x=row["Match Order"],
                line_dash="dot",
                line_width=1,
                line_color="rgba(214, 224, 220, 0.18)",
            )
            line_chart.add_annotation(
                x=row["Match Order"],
                y=1.05,
                yref="paper",
                text=row["Match Label"],
                showarrow=False,
                textangle=-28,
                font=dict(size=10, color="#d6e0dc"),
                bgcolor="rgba(11, 18, 16, 0.82)",
                bordercolor="rgba(214, 224, 220, 0.16)",
                borderpad=4,
            )

    metric_descriptions = [
        f"{metric_label(metric_key)}: {metric_help(metric_key)}"
        for metric_key in timeline_metrics
        if metric_help(metric_key)
    ]
    if metric_descriptions:
        st.caption(" | ".join(metric_descriptions))
    st.plotly_chart(line_chart, width="stretch")


if files_to_process:
    try:
        with st.spinner("Analyzing uploaded data..."):
            summaries = []
            summaries_by_file = {}
            skipped_analysis = []
            for file in files_to_process:
                file_name = file.name
                file_bytes = file.getvalue()
                sheet_for_file = default_sheet_by_file.get(file_name, sheet_name)
                try:
                    summary = cached_file_summary(
                        file_bytes=file_bytes,
                        file_name=file_name,
                        sheet_for_file=sheet_for_file,
                    )
                except Exception as exc:
                    logger.error("Failed to analyze %s: %s", file_name, exc)
                    skipped_analysis.append(file_name)
                    continue
                summaries.append(summary)
                summaries_by_file[file_name] = summary

            if skipped_analysis:
                st.warning(
                    f"Could not analyze {len(skipped_analysis)} file(s) due to data errors: "
                    + ", ".join(skipped_analysis)
                )

            if not summaries:
                st.error("No files could be analyzed. Please check your data files.")
                st.stop()

            full_summary = summaries[0] if len(summaries) == 1 else aggregate_season_summaries(summaries)
            full_summary = normalize_summary_players(full_summary)

        available_players = sorted(map(str, full_summary.index.tolist()))
        if len(available_players) == 1:
            selected_players = [available_players[0]]
        else:
            view_mode = st.radio(
                "View mode",
                ["Focused player", "Compare players"],
                index=0,
                horizontal=True,
            )

            if view_mode == "Focused player":
                selected_player = st.selectbox(
                    "Select player",
                    options=available_players,
                    index=0,
                )
                selected_players = [selected_player]
            else:
                default_compare = available_players[: min(4, len(available_players))]
                selected_players = st.multiselect(
                    "Select players to compare",
                    options=available_players,
                    default=default_compare,
                    help="Choose only the players you want to compare.",
                )

                if not selected_players:
                    st.warning("Please select at least one player to view stats.")
                    st.stop()

        eligible_files = sorted(
            [
                file_name
                for file_name, file_summary in summaries_by_file.items()
                if any(player in file_summary.index for player in selected_players)
            ]
        )

        if not eligible_files:
            st.warning("No files found for the selected player(s).")
            st.stop()

        parsed_dates_by_file = {file_name: parse_match_date_from_filename(file_name) for file_name in eligible_files}
        dated_files = [file_name for file_name in eligible_files if parsed_dates_by_file[file_name] is not None]
        undated_files = [file_name for file_name in eligible_files if parsed_dates_by_file[file_name] is None]

        files_after_date_filter = list(eligible_files)
        if dated_files:
            date_mode = st.radio(
                "Date filter",
                ["All dates", "Last 7 days", "Last 30 days", "Custom range"],
                index=0,
                horizontal=True,
            )

            include_undated = st.checkbox(
                "Include undated files",
                value=True,
                help="Keep files whose names do not include a YYYY-MM-DD date.",
            )

            date_values = [parsed_dates_by_file[file_name] for file_name in dated_files]
            min_date = min(date_values)
            max_date = max(date_values)
            start_date = min_date
            end_date = max_date

            if date_mode == "Last 7 days":
                end_date = date.today()
                start_date = end_date - timedelta(days=6)
            elif date_mode == "Last 30 days":
                end_date = date.today()
                start_date = end_date - timedelta(days=29)
            elif date_mode == "Custom range":
                selected_range = st.date_input(
                    "Match date range",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date,
                )
                if isinstance(selected_range, tuple) and len(selected_range) == 2:
                    start_date, end_date = selected_range

            filtered_dated_files = [
                file_name
                for file_name in dated_files
                if start_date <= parsed_dates_by_file[file_name] <= end_date
            ]

            files_after_date_filter = sorted(
                filtered_dated_files + (undated_files if include_undated else [])
            )

            st.caption(
                f"Date filter keeps {len(files_after_date_filter)} of {len(eligible_files)} file(s)."
            )
        else:
            st.caption("No dates detected in filenames, so all eligible files are included.")

        if not files_after_date_filter:
            st.warning("No files match the selected date window.")
            st.stop()

        selection_key = "selected_files_for_players"
        prev_options_key = "selected_files_options"
        if (
            selection_key not in st.session_state
            or prev_options_key not in st.session_state
            or st.session_state[prev_options_key] != files_after_date_filter
        ):
            st.session_state[selection_key] = files_after_date_filter
            st.session_state[prev_options_key] = files_after_date_filter

        st.caption("Include or exclude files for the selected player(s).")
        action_col1, action_col2 = st.columns(2)
        if action_col1.button("Select all files"):
            st.session_state[selection_key] = files_after_date_filter
        if action_col2.button("Clear all files"):
            st.session_state[selection_key] = []

        st.session_state[selection_key] = [
            file_name
            for file_name in st.session_state.get(selection_key, files_after_date_filter)
            if file_name in files_after_date_filter
        ]

        base_labels = {
            file_name: (
                parsed_dates_by_file[file_name].isoformat()
                if parsed_dates_by_file[file_name] is not None
                else "Undated"
            )
            for file_name in files_after_date_filter
        }
        label_counts = {}
        for label in base_labels.values():
            label_counts[label] = label_counts.get(label, 0) + 1

        seen_labels = {}
        display_labels = {}
        for file_name in files_after_date_filter:
            label = base_labels[file_name]
            if label_counts[label] == 1:
                display_labels[file_name] = label
            else:
                seen_labels[label] = seen_labels.get(label, 0) + 1
                display_labels[file_name] = f"{label} ({seen_labels[label]})"

        selected_files = st.multiselect(
            "Included files",
            options=files_after_date_filter,
            key=selection_key,
            format_func=lambda file_name: display_labels.get(file_name, file_name),
            help="All files are selected by default. Deselect any match you want to leave out.",
        )

        if not selected_files:
            st.warning("Please select at least one file to view stats.")
            st.stop()

        selected_summaries = [summaries_by_file[file_name] for file_name in selected_files]
        combined_summary = (
            selected_summaries[0]
            if len(selected_summaries) == 1
            else aggregate_season_summaries(selected_summaries)
        )
        combined_summary = normalize_summary_players(combined_summary)

        players_after_file_filter = [player for player in selected_players if player in combined_summary.index]
        if not players_after_file_filter:
            st.warning("Selected file set does not contain the selected player(s).")
            st.stop()

        filtered_summary = combined_summary.loc[players_after_file_filter]

        if len(selected_files) > 1:
            render_timeline_view(
                selected_files=selected_files,
                summaries_by_file=summaries_by_file,
                selected_players=players_after_file_filter,
                parsed_dates_by_file=parsed_dates_by_file,
            )
            st.markdown("---")

        render_charts(filtered_summary)

        download_data, filename = export_summary_bytes(filtered_summary, output_type)
        st.download_button(
            label="Download Summary",
            data=download_data,
            file_name=filename,
            mime="text/csv" if output_type == "csv" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except (DataLoadError, DataValidationError) as exc:
        st.error(str(exc))
        st.info("Tip: Confirm the file is a SwingVision export and the columns map correctly.")
else:
    st.info("Upload one SwingVision file or a folder of files to get started.")
