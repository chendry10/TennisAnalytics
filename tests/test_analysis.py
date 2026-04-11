import io
import unittest

import pandas as pd

from core.analysis import (
    aggregate_season_summaries,
    build_point_facts,
    canonicalize_player_label,
    calculate_double_fault_stats,
    calculate_plus_one_stats,
    calculate_point_length_outcomes,
    calculate_return_in_counts,
    calculate_return_attempts,
    calculate_return_percentages,
    calculate_return_win_counts,
    calculate_return_win_percentages,
    get_excel_sheet_names,
    load_df,
    normalize_summary_players,
    summarize_all,
    summarize_from_stats,
    validate_and_rename,
    POINT_LENGTH_BUCKETS,
)
from core.errors import DataLoadError, DataValidationError

class TestAnalysis(unittest.TestCase):
    def setUp(self) -> None:
        self.sample = pd.DataFrame(
            [
                {"Player": "A", "Shot": "Serve", "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": "Return", "Type": "return", "Result": "Out", "Point": "1"},
                {"Player": "A", "Shot": "Serve", "Type": "second_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": "Return", "Type": "return", "Result": "Net", "Point": "2"},
                {"Player": "B", "Shot": "Serve", "Type": "first_serve", "Result": "Out", "Point": "3"},
                {"Player": "A", "Shot": "Return", "Type": "return", "Result": "In", "Point": "3"},
                {"Player": "B", "Shot": "Serve", "Type": "second_serve", "Result": "In", "Point": "4"},
                {"Player": "A", "Shot": "Return", "Type": "return", "Result": "Out", "Point": "4"},
            ]
        )

    def test_summarize_all(self) -> None:
        summary = summarize_all(self.sample)

        self.assertIn("A", summary.index)
        self.assertIn("B", summary.index)

        self.assertAlmostEqual(summary.loc["A", "Overall First Serve %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "Overall Second Serve %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "First Serve Win %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "Second Serve Win %"], 100.0)
        self.assertEqual(summary.loc["A", "First Serve Attempts"], 1)
        self.assertEqual(summary.loc["A", "Second Serve Attempts"], 1)
        self.assertEqual(summary.loc["A", "First Serve In"], 1)
        self.assertEqual(summary.loc["A", "Second Serve In"], 1)

        self.assertAlmostEqual(summary.loc["B", "Overall First Serve %"], 0.0)
        self.assertAlmostEqual(summary.loc["B", "Overall Second Serve %"], 100.0)
        self.assertAlmostEqual(summary.loc["B", "First Serve Win %"], 0.0)
        self.assertAlmostEqual(summary.loc["B", "Second Serve Win %"], 100.0)
        self.assertEqual(summary.loc["B", "First Serve Attempts"], 1)
        self.assertEqual(summary.loc["B", "Second Serve Attempts"], 1)
        self.assertEqual(summary.loc["B", "First Serve In"], 0)
        self.assertEqual(summary.loc["B", "Second Serve In"], 1)

    def test_load_df_with_column_aliases(self) -> None:
        aliased = self.sample.rename(
            columns={
                "Player": "Player Name",
                "Shot": "Shot Type",
                "Type": "Serve Type",
                "Result": "Outcome",
                "Point": "Point #",
            }
        )
        buffer = io.BytesIO()
        aliased.to_csv(buffer, index=False)
        buffer.seek(0)

        df = load_df(buffer.getvalue(), file_name="sample.csv")
        self.assertListEqual(list(df.columns), ["Player", "Shot", "Type", "Result", "Point"])
        self.assertEqual(len(df), len(self.sample))

    def test_perfect_and_zero_percentages(self) -> None:
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": "Serve", "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": "Return", "Type": "return", "Result": "Out", "Point": "1"},
                {"Player": "A", "Shot": "Serve", "Type": "second_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": "Return", "Type": "return", "Result": "Out", "Point": "2"},
                {"Player": "B", "Shot": "Serve", "Type": "first_serve", "Result": "Out", "Point": "3"},
                {"Player": "A", "Shot": "Return", "Type": "return", "Result": "In", "Point": "3"},
                {"Player": "B", "Shot": "Serve", "Type": "second_serve", "Result": "Net", "Point": "4"},
                {"Player": "A", "Shot": "Return", "Type": "return", "Result": "In", "Point": "4"},
            ]
        )

        summary = summarize_all(df)

        self.assertAlmostEqual(summary.loc["A", "Overall First Serve %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "Overall Second Serve %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "First Serve Win %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "Second Serve Win %"], 100.0)
        self.assertEqual(summary.loc["A", "First Serve Attempts"], 1)
        self.assertEqual(summary.loc["A", "Second Serve Attempts"], 1)

        self.assertAlmostEqual(summary.loc["B", "Overall First Serve %"], 0.0)
        self.assertAlmostEqual(summary.loc["B", "Overall Second Serve %"], 0.0)
        self.assertAlmostEqual(summary.loc["B", "First Serve Win %"], 0.0)
        self.assertAlmostEqual(summary.loc["B", "Second Serve Win %"], 0.0)
        self.assertEqual(summary.loc["B", "First Serve Attempts"], 1)
        self.assertEqual(summary.loc["B", "Second Serve Attempts"], 1)
        self.assertEqual(summary.loc["B", "First Serve In"], 0)
        self.assertEqual(summary.loc["B", "Second Serve In"], 0)

    def test_first_serve_fault_counts_attempt(self) -> None:
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": "Serve", "Type": "first_serve", "Result": "Out", "Point": "1"},
                {"Player": "A", "Shot": "Serve", "Type": "second_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": "Return", "Type": "return", "Result": "Out", "Point": "1"},
            ]
        )

        summary = summarize_all(df)

        self.assertAlmostEqual(summary.loc["A", "Overall First Serve %"], 0.0)
        self.assertAlmostEqual(summary.loc["A", "Overall Second Serve %"], 100.0)
        self.assertAlmostEqual(summary.loc["A", "First Serve Win %"], 0.0)
        self.assertAlmostEqual(summary.loc["A", "Second Serve Win %"], 100.0)
        self.assertEqual(summary.loc["A", "First Serve Attempts"], 1)
        self.assertEqual(summary.loc["A", "Second Serve Attempts"], 1)
        self.assertEqual(summary.loc["A", "First Serve In"], 0)
        self.assertEqual(summary.loc["A", "Second Serve In"], 1)

    def test_summarize_from_stats_sheet(self) -> None:
        stats = pd.DataFrame(
            {
                "Stat Name": [
                    "1st Serves",
                    "1st Serves In",
                    "1st Serves Won",
                    "2nd Serves",
                    "2nd Serves In",
                    "2nd Serves Won",
                ],
                "Host Set 1": [10, 6, 4, 5, 4, 2],
                "Guest Set 1": [8, 4, 3, 6, 3, 1],
            }
        )
        settings = pd.DataFrame(
            {
                "Host Team": ["Host Player"],
                "Guest Team": ["Guest Player"],
            }
        )

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            settings.to_excel(writer, sheet_name="Settings", index=False)
            stats.to_excel(writer, sheet_name="Stats", index=False)
        buffer.seek(0)

        summary = summarize_from_stats(buffer.getvalue(), file_name="sample.xlsx")

        self.assertIn("Host Player", summary.index)
        self.assertIn("Guest Player", summary.index)
        self.assertAlmostEqual(summary.loc["Host Player", "Overall First Serve %"], 60.0)
        self.assertAlmostEqual(summary.loc["Host Player", "Overall Second Serve %"], 80.0)
        self.assertAlmostEqual(summary.loc["Host Player", "First Serve Win %"], 40.0)
        self.assertAlmostEqual(summary.loc["Host Player", "Second Serve Win %"], 40.0)
        self.assertEqual(summary.loc["Host Player", "First Serve Attempts"], 10)
        self.assertEqual(summary.loc["Host Player", "Second Serve Attempts"], 5)
        self.assertEqual(summary.loc["Host Player", "First Serve In"], 6)
        self.assertEqual(summary.loc["Host Player", "Second Serve In"], 4)

    def test_aggregate_season_summaries_includes_all_players(self) -> None:
        match_one = pd.DataFrame(
            {
                "Overall First Serve %": [70.0, 40.0],
                "Overall Second Serve %": [80.0, 50.0],
                "First Serve In": [7, 4],
                "Second Serve In": [8, 5],
                "First Serve Attempts": [10, 10],
                "Second Serve Attempts": [10, 10],
                "First Serve Wins": [6, 3],
                "Second Serve Wins": [7, 4],
                "First Serve Win %": [60.0, 30.0],
                "Second Serve Win %": [70.0, 40.0],
            },
            index=["Alice", "Bob"],
        )
        match_one.index.name = "Player"

        match_two = pd.DataFrame(
            {
                "Overall First Serve %": [60.0, 55.0],
                "Overall Second Serve %": [75.0, 65.0],
                "First Serve In": [6, 5],
                "Second Serve In": [9, 6],
                "First Serve Attempts": [10, 10],
                "Second Serve Attempts": [12, 9],
                "First Serve Wins": [5, 4],
                "Second Serve Wins": [8, 5],
                "First Serve Win %": [50.0, 40.0],
                "Second Serve Win %": [66.67, 55.56],
            },
            index=["Alice", "Cara"],
        )
        match_two.index.name = "Player"

        season = aggregate_season_summaries([match_one, match_two])

        self.assertListEqual(sorted(season.index.tolist()), ["Alice", "Bob", "Cara"])
        self.assertEqual(season.loc["Alice", "First Serve Attempts"], 20)
        self.assertEqual(season.loc["Bob", "First Serve Attempts"], 10)
        self.assertEqual(season.loc["Cara", "Second Serve Attempts"], 9)

    def test_normalize_summary_players_collapses_generic_opponent_labels(self) -> None:
        summary = pd.DataFrame(
            {
                "First Serve In": [5, 3],
                "Second Serve In": [4, 2],
                "First Serve Attempts": [10, 8],
                "Second Serve Attempts": [6, 4],
                "First Serve Wins": [4, 2],
                "Second Serve Wins": [3, 1],
                "First Serve Win %": [40.0, 25.0],
                "Second Serve Win %": [50.0, 25.0],
                "Overall First Serve %": [50.0, 37.5],
                "Overall Second Serve %": [66.7, 50.0],
            },
            index=["Opponent", "Player 2"],
        )
        summary.index.name = "Player"

        normalized = normalize_summary_players(summary)

        self.assertListEqual(normalized.index.tolist(), ["Opponent"])
        self.assertEqual(normalized.loc["Opponent", "First Serve Attempts"], 18)
        self.assertEqual(normalized.loc["Opponent", "Second Serve Attempts"], 10)

    def test_canonicalize_player_label_variants(self) -> None:
        self.assertEqual(canonicalize_player_label("Opponent"), "Opponent")
        self.assertEqual(canonicalize_player_label("player 2"), "Opponent")
        self.assertEqual(canonicalize_player_label("PLAYER_2"), "Opponent")
        self.assertEqual(canonicalize_player_label("unknown"), "Opponent")
        self.assertEqual(canonicalize_player_label("  "), "Opponent")
        self.assertEqual(canonicalize_player_label("Bailey Bell"), "Bailey Bell")

    def test_get_excel_sheet_names_invalid_excel_raises(self) -> None:
        with self.assertRaises(DataLoadError):
            get_excel_sheet_names(b"not-a-real-excel", file_name="broken.xlsx")

    def test_load_df_drops_rows_without_point(self) -> None:
        with_missing_points = pd.DataFrame(
            [
                {"Player": "A", "Shot": "Serve", "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": "Return", "Type": "return", "Result": "Out", "Point": None},
                {"Player": "A", "Shot": "Serve", "Type": "second_serve", "Result": "In", "Point": "2"},
            ]
        )
        buffer = io.BytesIO()
        with_missing_points.to_csv(buffer, index=False)
        buffer.seek(0)

        df = load_df(buffer.getvalue(), file_name="points.csv")

        self.assertEqual(len(df), 2)
        self.assertTrue(df["Point"].notna().all())

    def test_point_length_outcomes(self) -> None:
        # Build a dataset with varying rally lengths
        df = pd.DataFrame(
            [
                # Point 1: 2 shots (0-4 bucket), A serves, B returns Out → A wins
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "Out", "Point": "1"},
                # Point 2: 6 shots (5-10 bucket), A serves, rally, B wins (A hits Out)
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "2"},
                {"Player": "A", "Shot": 3, "Type": "serve_plus_one", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 4, "Type": "return_plus_one", "Result": "In", "Point": "2"},
                {"Player": "A", "Shot": 5, "Type": "in_play", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 6, "Type": "in_play", "Result": "In", "Point": "2"},
                # Point 3: 3 shots (0-4 bucket), B serves, A returns In (last shot In = A wins)
                {"Player": "B", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "3"},
                {"Player": "A", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "3"},
                {"Player": "B", "Shot": 3, "Type": "serve_plus_one", "Result": "Out", "Point": "3"},
            ]
        )
        outcomes = calculate_point_length_outcomes(df)

        self.assertIn("A", outcomes.index)
        self.assertIn("B", outcomes.index)

        # 0-4 bucket: 2 points total, A wins both (pt1 and pt3)
        self.assertEqual(outcomes.loc["A", "0-4_shots_Wins"], 2)
        self.assertEqual(outcomes.loc["A", "0-4_shots_Total"], 2)
        self.assertAlmostEqual(outcomes.loc["A", "0-4_shots_Win%"], 100.0)
        self.assertEqual(outcomes.loc["B", "0-4_shots_Wins"], 0)

        # 5-10 bucket: 1 point, B wins (last shot In by B)
        self.assertEqual(outcomes.loc["B", "5-10_shots_Wins"], 1)
        self.assertEqual(outcomes.loc["A", "5-10_shots_Wins"], 0)

        # 11+ bucket: no points
        self.assertEqual(outcomes.loc["A", "11plus_shots_Total"], 0)

    def test_return_stats(self) -> None:
        df = pd.DataFrame(
            [
                # Point 1: A serves first, B returns first_return In, A Out → B wins
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "1"},
                {"Player": "A", "Shot": 3, "Type": "serve_plus_one", "Result": "Out", "Point": "1"},
                # Point 2: A serves first (fault), then second serve, B second_return Out → A wins
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "Out", "Point": "2"},
                {"Player": "A", "Shot": 2, "Type": "second_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 3, "Type": "second_return", "Result": "Out", "Point": "2"},
                # Point 3: B serves first, A first_return Net → B wins
                {"Player": "B", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "3"},
                {"Player": "A", "Shot": 2, "Type": "first_return", "Result": "Net", "Point": "3"},
            ]
        )

        ret_in = calculate_return_in_counts(df)
        self.assertEqual(ret_in.loc["B", "First Return In"], 1)
        self.assertEqual(ret_in.loc["B", "Second Return In"], 0)
        # A's only return was Net, so A may not appear in ret_in (no "In" returns)
        if "A" in ret_in.index:
            self.assertEqual(ret_in.loc["A", "First Return In"], 0)

        ret_att = calculate_return_attempts(df)
        self.assertEqual(ret_att.loc["B", "First Return Attempts"], 1)
        self.assertEqual(ret_att.loc["B", "Second Return Attempts"], 1)
        self.assertEqual(ret_att.loc["A", "First Return Attempts"], 1)

        ret_pcts = calculate_return_percentages(df)
        self.assertAlmostEqual(ret_pcts.loc["B", "First Return In %"], 100.0)
        self.assertAlmostEqual(ret_pcts.loc["B", "Second Return In %"], 0.0)
        self.assertAlmostEqual(ret_pcts.loc["A", "First Return In %"], 0.0)

        ret_wins = calculate_return_win_counts(df)
        self.assertEqual(ret_wins.loc["B", "First Return Wins"], 1)
        self.assertEqual(ret_wins.loc["B", "Second Return Wins"], 0)

        ret_win_pcts = calculate_return_win_percentages(df)
        self.assertAlmostEqual(ret_win_pcts.loc["B", "First Return Win %"], 100.0)
        self.assertAlmostEqual(ret_win_pcts.loc["B", "Second Return Win %"], 0.0)

    def test_summarize_all_includes_return_and_point_length_columns(self) -> None:
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "Out", "Point": "1"},
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "Out", "Point": "2"},
                {"Player": "A", "Shot": 2, "Type": "second_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 3, "Type": "second_return", "Result": "In", "Point": "2"},
                {"Player": "A", "Shot": 4, "Type": "serve_plus_one", "Result": "Out", "Point": "2"},
                {"Player": "B", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "3"},
                {"Player": "A", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "3"},
                {"Player": "B", "Shot": 3, "Type": "serve_plus_one", "Result": "Out", "Point": "3"},
            ]
        )

        summary = summarize_all(df)

        # Return columns should exist
        self.assertIn("First Return In", summary.columns)
        self.assertIn("First Return Attempts", summary.columns)
        self.assertIn("First Return In %", summary.columns)
        self.assertIn("First Return Win %", summary.columns)

        # Point-length columns should exist
        self.assertIn("0-4_shots_Win%", summary.columns)
        self.assertIn("5-10_shots_Win%", summary.columns)
        self.assertIn("11plus_shots_Win%", summary.columns)

    def test_empty_dataframe(self) -> None:
        """summarize_all on an empty DataFrame with correct columns should not crash."""
        df = pd.DataFrame(columns=["Player", "Shot", "Type", "Result", "Point"])
        summary = summarize_all(df)
        self.assertTrue(summary.empty)

    def test_single_point_match(self) -> None:
        """A match with exactly one point should produce valid percentages."""
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "Out", "Point": "1"},
            ]
        )
        summary = summarize_all(df)
        self.assertIn("A", summary.index)
        self.assertAlmostEqual(summary.loc["A", "First Serve Win %"], 100.0)
        self.assertEqual(summary.loc["A", "First Serve Attempts"], 1)

    def test_missing_stats_sheet_falls_back(self) -> None:
        """An Excel file without a Stats sheet should still be analyzable via Shots."""
        shots = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "Out", "Point": "1"},
            ]
        )
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            shots.to_excel(writer, sheet_name="Shots", index=False)
        buffer.seek(0)

        df = load_df(buffer.getvalue(), sheet="Shots", file_name="no_stats.xlsx")
        summary = summarize_all(df)
        self.assertIn("A", summary.index)

    def test_malformed_result_column(self) -> None:
        """Non-standard Result values (numbers, blanks) should not crash analysis."""
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": 123, "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "", "Point": "1"},
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "Out", "Point": "2"},
            ]
        )
        df = validate_and_rename(df)
        df = df.dropna(subset=["Point"])
        summary = summarize_all(df)
        self.assertIn("A", summary.index)

    def test_single_player_only(self) -> None:
        """Data with only one player should not cause index errors in pivots."""
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "A", "Shot": 2, "Type": "in_play", "Result": "In", "Point": "1"},
            ]
        )
        summary = summarize_all(df)
        self.assertIn("A", summary.index)
        self.assertEqual(len(summary.index), 1)

    def test_no_returns_in_data(self) -> None:
        """Data with no return types should produce zero return stats, not errors."""
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "1"},
                {"Player": "B", "Shot": 2, "Type": "in_play", "Result": "Out", "Point": "1"},
            ]
        )
        summary = summarize_all(df)
        self.assertEqual(summary.loc["A", "First Return In"], 0)
        self.assertEqual(summary.loc["A", "First Return Attempts"], 0)

    def test_build_point_facts_captures_double_fault_and_plus_one_sequences(self) -> None:
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "Out", "Point": "1"},
                {"Player": "A", "Shot": 2, "Type": "second_serve", "Result": "Net", "Point": "1"},
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "2"},
                {"Player": "A", "Shot": 3, "Type": "serve_plus_one", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 4, "Type": "return_plus_one", "Result": "Out", "Point": "2"},
            ]
        )

        point_facts = build_point_facts(df)

        self.assertEqual(len(point_facts), 2)
        self.assertTrue(point_facts.loc[point_facts["Point"] == "1", "Double_Fault"].iloc[0])
        self.assertEqual(point_facts.loc[point_facts["Point"] == "1", "Winner"].iloc[0], "B")
        self.assertEqual(point_facts.loc[point_facts["Point"] == "1", "Returner"].iloc[0], "B")
        self.assertTrue(point_facts.loc[point_facts["Point"] == "2", "Serve_Plus_One_Attempted"].iloc[0])
        self.assertTrue(point_facts.loc[point_facts["Point"] == "2", "Return_Plus_One_Attempted"].iloc[0])

    def test_new_transition_and_double_fault_metrics(self) -> None:
        df = pd.DataFrame(
            [
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "Out", "Point": "1"},
                {"Player": "A", "Shot": 2, "Type": "second_serve", "Result": "Net", "Point": "1"},
                {"Player": "A", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "2"},
                {"Player": "A", "Shot": 3, "Type": "serve_plus_one", "Result": "In", "Point": "2"},
                {"Player": "B", "Shot": 4, "Type": "return_plus_one", "Result": "Out", "Point": "2"},
                {"Player": "B", "Shot": 1, "Type": "first_serve", "Result": "In", "Point": "3"},
                {"Player": "A", "Shot": 2, "Type": "first_return", "Result": "In", "Point": "3"},
                {"Player": "B", "Shot": 3, "Type": "serve_plus_one", "Result": "Out", "Point": "3"},
            ]
        )

        summary = summarize_all(df)
        double_faults = calculate_double_fault_stats(df)
        serve_plus_one = calculate_plus_one_stats(df, "serve")
        return_plus_one = calculate_plus_one_stats(df, "return")

        self.assertEqual(double_faults.loc["A", "Double Faults"], 1)
        self.assertAlmostEqual(double_faults.loc["A", "Double Fault Rate"], 100.0)
        self.assertEqual(serve_plus_one.loc["A", "Serve +1 Attempts"], 1)
        self.assertEqual(serve_plus_one.loc["A", "Serve +1 In"], 1)
        self.assertAlmostEqual(serve_plus_one.loc["A", "Serve +1 Win %"], 100.0)
        self.assertEqual(return_plus_one.loc["B", "Return +1 Attempts"], 1)
        self.assertAlmostEqual(return_plus_one.loc["B", "Return +1 In %"], 0.0)
        self.assertIn("Double Fault Rate", summary.columns)
        self.assertIn("Serve +1 Win %", summary.columns)
        self.assertIn("Return +1 Win %", summary.columns)


if __name__ == "__main__":
    unittest.main()
