import hashlib
import pandas as pd
import plotly.express as px
import streamlit as st

from core.analysis import (
    REQUIRED_COLUMNS,
    aggregate_season_summaries,
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
    --card: #141a19;
    --ink: #f4f7f6;
    --muted: #a7b3af;
}

html, body, [class*="css"] {
    font-family: "Inter", "Segoe UI", system-ui, sans-serif;
    background-color: #0b1210;
    color: var(--ink);
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
    font-size: 2.2rem;
}

.hero p {
    color: var(--muted);
    margin: 0;
    font-size: 1rem;
}

.metric-card {
    background: var(--card);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 14px;
    padding: 14px 18px;
}

.stMetric label,
div[data-testid="stMetricLabel"] {
    color: #d7e3df !important;
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
    color: #f4f7f6 !important;
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
}

div[role="listbox"],
ul[role="listbox"] {
    background-color: #0f1715 !important;
    color: var(--ink) !important;
}

div[role="option"] {
    color: var(--ink) !important;
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

files_to_process = []
sheet_names_by_file = {}
default_sheet_by_file = {}

with st.sidebar:
    st.header("Upload Data")
    upload_mode = st.radio(
        "Choose input type",
        ["Single file", "Folder"],
        index=0,
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
            type=["xlsx", "xls", "xlsm", "csv"],
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
                    sheet_names_by_file[file_name] = get_excel_sheet_names(file_bytes, file_name)
                except DataLoadError:
                    invalid_excel_files.append(file_name)
                    continue

                if not sheet_names_by_file[file_name]:
                    st.error(f"No readable sheets found in {file_name}.")
                    st.stop()
                default_sheet_by_file[file_name] = 1 if len(sheet_names_by_file[file_name]) > 1 else 0

            validated_files.append(file)

        if skipped_temp_files:
            st.caption(f"Skipped {skipped_temp_files} temporary Excel lock file(s) (~$...).")
        if skipped_metadata_files:
            st.caption(f"Skipped {skipped_metadata_files} metadata/system file(s).")

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
        return

    key_cols = [
        "First Serve Win %",
        "Second Serve Win %",
        "Overall First Serve %",
        "Overall Second Serve %",
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


def render_charts(summary: pd.DataFrame) -> None:
    st.subheader("Serve Win % by Player")
    win_long = summary.reset_index().melt(
        id_vars="Player",
        value_vars=["First Serve Win %", "Second Serve Win %"],
        var_name="Serve Type",
        value_name="Win %",
    )
    win_chart = px.bar(
        win_long,
        x="Player",
        y="Win %",
        color="Serve Type",
        barmode="group",
        text="Win %",
        color_discrete_sequence=["#18a66c", "#d6ff3d"],
    )
    win_chart.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    win_chart.update_layout(
        template="plotly_dark",
        height=320,
        margin=dict(l=20, r=20, t=20, b=40),
        yaxis=dict(range=[0, 100], title="Win %"),
        xaxis=dict(title=None),
        legend_title_text="",
    )
    st.plotly_chart(win_chart, width="stretch")

    st.subheader("Overall Serve In %")
    overall_long = summary.reset_index().melt(
        id_vars="Player",
        value_vars=["Overall First Serve %", "Overall Second Serve %"],
        var_name="Serve Type",
        value_name="In %",
    )
    overall_chart = px.bar(
        overall_long,
        x="Player",
        y="In %",
        color="Serve Type",
        barmode="group",
        text="In %",
        color_discrete_sequence=["#00c2ff", "#18a66c"],
    )
    overall_chart.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    overall_chart.update_layout(
        template="plotly_dark",
        height=320,
        margin=dict(l=20, r=20, t=20, b=40),
        yaxis=dict(range=[0, 100], title="In %"),
        xaxis=dict(title=None),
        legend_title_text="",
    )
    st.plotly_chart(overall_chart, width="stretch")

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


if files_to_process:
    try:
        with st.spinner("Analyzing uploaded data..."):
            summaries = []
            for file in files_to_process:
                file_name = file.name
                file_bytes = file.getvalue()
                lower_name = file_name.lower()
                is_excel = lower_name.endswith(EXCEL_EXTENSIONS)
                sheet_names = sheet_names_by_file.get(file_name, []) if is_excel else []

                summary = None
                if is_excel and "Stats" in sheet_names:
                    summary = summarize_from_stats(file_bytes, file_name=file_name)

                    if summary is not None and not summary.empty:
                        attempt_cols = ["First Serve Attempts", "Second Serve Attempts"]
                        if all(col in summary.columns for col in attempt_cols):
                            total_attempts = summary[attempt_cols].to_numpy().sum()
                            if total_attempts == 0:
                                summary = None

                if summary is None or summary.empty:
                    column_map_to_use = column_map if column_map else None
                    sheet_for_file = default_sheet_by_file.get(file_name, sheet_name)
                    df = load_df(
                        file_bytes,
                        sheet=sheet_for_file,
                        column_map=column_map_to_use,
                        file_name=file_name,
                    )
                    summary = summarize_all(df)

                summary = normalize_summary_players(summary)
                summaries.append(summary)

            summary = summaries[0] if len(summaries) == 1 else aggregate_season_summaries(summaries)
            summary = normalize_summary_players(summary)

        available_players = sorted(map(str, summary.index.tolist()))
        if len(available_players) == 1:
            filtered_summary = summary.loc[[available_players[0]]]
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
                filtered_summary = summary.loc[[selected_player]]
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
                filtered_summary = summary.loc[selected_players]

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
