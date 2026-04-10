import hashlib
import io
import logging
import re
from datetime import date, timedelta
import pandas as pd
import plotly.express as px
import streamlit as st

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

from core.analysis import (
    REQUIRED_COLUMNS,
    POINT_LENGTH_BUCKETS,
    excel_engine,
    validate_and_rename,
    aggregate_season_summaries,
    calculate_point_length_outcomes,
    calculate_return_attempts,
    calculate_return_in_counts,
    calculate_return_percentages,
    calculate_return_win_counts,
    calculate_return_win_percentages,
    export_summary_bytes,
    get_excel_sheet_names,
    get_raw_columns,
    guess_column_map,
    load_df,
    normalize_summary_players,
    summarize_from_stats,
    summarize_all,
)
from core.errors import DataLoadError, DataValidationError

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


@st.cache_data(show_spinner=False)
def cached_excel_sheet_names(file_bytes: bytes, file_name: str) -> list[str]:
    return get_excel_sheet_names(file_bytes, file_name)


@st.cache_data(show_spinner=False)
def cached_file_summary(
    file_bytes: bytes,
    file_name: str,
    sheet_for_file: str | int | None,
) -> pd.DataFrame:
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

            if "Stats" in sheet_names:
                logger.info("Using Stats sheet for %s", file_name)
                stats_df = xls.parse("Stats")
                settings_df = xls.parse("Settings")
                summary = summarize_from_stats(
                    stats_df=stats_df, settings_df=settings_df,
                )

                if summary is not None and not summary.empty:
                    attempt_cols = ["First Serve Attempts", "Second Serve Attempts"]
                    if all(col in summary.columns for col in attempt_cols):
                        total_attempts = summary[attempt_cols].to_numpy().sum()
                        if total_attempts == 0:
                            summary = None

            # Read Shots sheet once for return + point-length stats (or full analysis)
            shots_sheet = sheet_for_file if sheet_for_file is not None else (
                "Shots" if "Shots" in sheet_names else 0
            )
            try:
                raw_df = validate_and_rename(xls.parse(shots_sheet))
                raw_df = raw_df.dropna(subset=["Point"])
            except Exception:
                raw_df = None

            if summary is not None and not summary.empty:
                # Augment Stats-based summary with return + point-length data from raw shots
                if raw_df is not None and not raw_df.empty:
                    ret_in = calculate_return_in_counts(raw_df)
                    ret_att = calculate_return_attempts(raw_df)
                    ret_pcts = calculate_return_percentages(raw_df)
                    ret_wins = calculate_return_win_counts(raw_df)
                    ret_win_pcts = calculate_return_win_percentages(raw_df)
                    pl_outcomes = calculate_point_length_outcomes(raw_df)

                    for extra in [ret_in, ret_att, ret_pcts, ret_wins, ret_win_pcts, pl_outcomes]:
                        if not extra.empty:
                            summary = summary.join(extra, how="left").fillna(0)
            else:
                # No Stats sheet or empty stats — compute everything from raw shots
                logger.info("No Stats sheet for %s, using raw shots", file_name)
                if raw_df is not None and not raw_df.empty:
                    summary = summarize_all(raw_df)

    if summary is None or summary.empty:
        logger.warning("Falling back to full load_df for %s", file_name)
        df = load_df(
            file_bytes,
            sheet=sheet_for_file,
            column_map=None,
            file_name=file_name,
        )
        summary = summarize_all(df)

    return normalize_summary_players(summary)


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
                default_sheet_by_file[file_name] = 1 if len(sheet_names_by_file[file_name]) > 1 else 0

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



def render_metrics(summary: pd.DataFrame) -> None:
    st.subheader("Key Metrics")
    if summary.empty:
        st.info("No player data available for the current selection.")
        return

    if len(summary.index) == 1:
        player = summary.index[0]
        row = summary.loc[player]
        st.caption(f"Focused view: {player}")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("1st Serve Win %", f"{float(row.get('First Serve Win %', 0)):.1f}%")
        col2.metric("2nd Serve Win %", f"{float(row.get('Second Serve Win %', 0)):.1f}%")
        col3.metric("1st Serve In %", f"{float(row.get('Overall First Serve %', 0)):.1f}%")
        col4.metric("2nd Serve In %", f"{float(row.get('Overall Second Serve %', 0)):.1f}%")

        col5, col6, col7, col8 = st.columns(4)
        col5.metric("1st Return Win %", f"{float(row.get('First Return Win %', 0)):.1f}%")
        col6.metric("2nd Return Win %", f"{float(row.get('Second Return Win %', 0)):.1f}%")
        col7.metric("1st Return In %", f"{float(row.get('First Return In %', 0)):.1f}%")
        col8.metric("2nd Return In %", f"{float(row.get('Second Return In %', 0)):.1f}%")
        return

    key_cols = [
        "First Serve Win %",
        "Second Serve Win %",
        "Overall First Serve %",
        "Overall Second Serve %",
        "First Return Win %",
        "Second Return Win %",
        "First Return In %",
        "Second Return In %",
    ]
    compact = summary.reindex(columns=[col for col in key_cols if col in summary.columns]).sort_values(
        by="First Serve Win %", ascending=False
    )
    styled = compact.style.format(
        {
            "First Serve Win %": "{:.1f}%",
            "Second Serve Win %": "{:.1f}%",
            "Overall First Serve %": "{:.1f}%",
            "Overall Second Serve %": "{:.1f}%",
            "First Return Win %": "{:.1f}%",
            "Second Return Win %": "{:.1f}%",
            "First Return In %": "{:.1f}%",
            "Second Return In %": "{:.1f}%",
        }
    )
    st.dataframe(styled, width="stretch")


def render_table(summary: pd.DataFrame) -> None:
    st.subheader("Full Serve Summary")
    ordered = [
        "First Serve In",
        "First Serve Attempts",
        "Overall First Serve %",
        "Second Serve In",
        "Second Serve Attempts",
        "Overall Second Serve %",
        "First Serve Win %",
        "Second Serve Win %",
    ]
    display = summary.reindex(columns=[col for col in ordered if col in summary.columns])
    formatters = {
        "Overall First Serve %": "{:.2f}",
        "Overall Second Serve %": "{:.2f}",
        "First Serve Win %": "{:.2f}",
        "Second Serve Win %": "{:.2f}",
        "First Serve Attempts": "{:.0f}",
        "Second Serve Attempts": "{:.0f}",
        "First Serve In": "{:.0f}",
        "Second Serve In": "{:.0f}",
    }
    styled = (
        display.style.format(formatters)
        .set_table_styles([
            {"selector": "th", "props": [("color", "#f4f7f6"), ("background", "#0f1715")]},
            {"selector": "td", "props": [("color", "#f4f7f6"), ("background", "#101816")]},
        ])
    )
    st.dataframe(styled, width="stretch")

    st.subheader("Full Return Summary")
    return_ordered = [
        "First Return In",
        "First Return Attempts",
        "First Return In %",
        "Second Return In",
        "Second Return Attempts",
        "Second Return In %",
        "First Return Win %",
        "Second Return Win %",
    ]
    return_display = summary.reindex(columns=[col for col in return_ordered if col in summary.columns])
    if return_display.empty or return_display.columns.empty:
        st.caption("No return data available.")
    else:
        ret_formatters = {
            "First Return In %": "{:.2f}",
            "Second Return In %": "{:.2f}",
            "First Return Win %": "{:.2f}",
            "Second Return Win %": "{:.2f}",
            "First Return Attempts": "{:.0f}",
            "Second Return Attempts": "{:.0f}",
            "First Return In": "{:.0f}",
            "Second Return In": "{:.0f}",
        }
        ret_styled = (
            return_display.style.format(ret_formatters)
            .set_table_styles([
                {"selector": "th", "props": [("color", "#f4f7f6"), ("background", "#0f1715")]},
                {"selector": "td", "props": [("color", "#f4f7f6"), ("background", "#101816")]},
            ])
        )
        st.dataframe(ret_styled, width="stretch")


def render_charts(summary: pd.DataFrame) -> None:
    st.subheader("Overall Serve In % vs Win %")
    compare_long = summary.reset_index().melt(
        id_vars="Player",
        value_vars=[
            "Overall First Serve %",
            "First Serve Win %",
            "Overall Second Serve %",
            "Second Serve Win %",
        ],
        var_name="Serve Metric",
        value_name="Percentage",
    )
    compare_chart = px.bar(
        compare_long,
        x="Player",
        y="Percentage",
        color="Serve Metric",
        barmode="group",
        text="Percentage",
        color_discrete_sequence=["#00c2ff", "#d6ff3d", "#18a66c", "#9be564"],
    )
    compare_chart.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    compare_chart.update_layout(
        template="plotly_dark",
        height=340,
        margin=dict(l=20, r=20, t=20, b=40),
        yaxis=dict(range=[0, 100], title="Percentage"),
        xaxis=dict(title=None),
        legend_title_text="",
    )
    st.plotly_chart(compare_chart, width="stretch")

    # ---------- Return charts ----------
    return_win_vars = [c for c in ["First Return Win %", "Second Return Win %"] if c in summary.columns]
    if return_win_vars:
        st.subheader("Return Win % by Player")
        ret_win_long = summary.reset_index().melt(
            id_vars="Player",
            value_vars=return_win_vars,
            var_name="Return Type",
            value_name="Win %",
        )
        ret_win_chart = px.bar(
            ret_win_long,
            x="Player",
            y="Win %",
            color="Return Type",
            barmode="group",
            text="Win %",
            color_discrete_sequence=["#ff6f61", "#ffb347"],
        )
        ret_win_chart.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        ret_win_chart.update_layout(
            template="plotly_dark",
            height=320,
            margin=dict(l=20, r=20, t=20, b=40),
            yaxis=dict(range=[0, 100], title="Win %"),
            xaxis=dict(title=None),
            legend_title_text="",
        )
        st.plotly_chart(ret_win_chart, width="stretch")

    # ---------- Point-length outcome chart ----------
    bucket_win_cols = [f"{b.replace(' ', '_').replace('+', 'plus')}_Win%" for b in POINT_LENGTH_BUCKETS]
    available_bucket_cols = [c for c in bucket_win_cols if c in summary.columns]
    if available_bucket_cols:
        st.subheader("Win % by Rally Length")
        bucket_labels = {
            "0-4_shots_Win%": "0-4 shots",
            "5-10_shots_Win%": "5-10 shots",
            "11plus_shots_Win%": "11+ shots",
        }
        pl_long = summary.reset_index().melt(
            id_vars="Player",
            value_vars=available_bucket_cols,
            var_name="Rally Length",
            value_name="Win %",
        )
        pl_long["Rally Length"] = pl_long["Rally Length"].map(bucket_labels)
        pl_chart = px.bar(
            pl_long,
            x="Rally Length",
            y="Win %",
            color="Player",
            barmode="group",
            text="Win %",
            color_discrete_sequence=["#18a66c", "#d6ff3d", "#00c2ff", "#ff6f61"],
        )
        pl_chart.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        pl_chart.update_layout(
            template="plotly_dark",
            height=340,
            margin=dict(l=20, r=20, t=20, b=40),
            yaxis=dict(range=[0, 100], title="Win %"),
            xaxis=dict(title=None, categoryorder="array", categoryarray=["0-4 shots", "5-10 shots", "11+ shots"]),
            legend_title_text="",
        )
        st.plotly_chart(pl_chart, width="stretch")


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

        render_metrics(filtered_summary)
        render_table(filtered_summary)
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
