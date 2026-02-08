import io
import unittest

import pandas as pd

from core.analysis import load_df, summarize_all, summarize_from_stats

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


if __name__ == "__main__":
    unittest.main()
