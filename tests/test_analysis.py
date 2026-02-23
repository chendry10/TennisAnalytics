import io
import unittest

import pandas as pd

from core.analysis import (
    aggregate_season_summaries,
    canonicalize_player_label,
    get_excel_sheet_names,
    load_df,
    normalize_summary_players,
    summarize_all,
    summarize_from_stats,
)
from core.errors import DataLoadError

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


if __name__ == "__main__":
    unittest.main()
