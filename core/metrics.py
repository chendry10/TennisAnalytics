from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

MetricKind = Literal["count", "percent"]


@dataclass(frozen=True)
class MetricDefinition:
    key: str
    label: str
    description: str
    kind: MetricKind
    chart_group: str | None = None
    timeline: bool = False
    numerator: str | None = None
    denominator: str | None = None


POINT_LENGTH_BUCKETS = ("0-4 shots", "5-10 shots", "11+ shots")


def _bucket_tag(bucket: str) -> str:
    return bucket.replace(" ", "_").replace("+", "plus")


RALLY_COUNT_KEYS = [
    f"{_bucket_tag(bucket)}_{suffix}"
    for bucket in POINT_LENGTH_BUCKETS
    for suffix in ("Wins", "Losses", "Total")
]

RALLY_RATE_KEYS = [f"{_bucket_tag(bucket)}_Win%" for bucket in POINT_LENGTH_BUCKETS]


METRIC_DEFINITIONS = {
    "Overall First Serve %": MetricDefinition(
        key="Overall First Serve %",
        label="1st Serve In %",
        description="Share of first-serve attempts that landed in.",
        kind="percent",
        chart_group="serve",
        timeline=True,
        numerator="First Serve In",
        denominator="First Serve Attempts",
    ),
    "Overall Second Serve %": MetricDefinition(
        key="Overall Second Serve %",
        label="2nd Serve In %",
        description="Share of second-serve attempts that landed in.",
        kind="percent",
        chart_group="serve",
        timeline=True,
        numerator="Second Serve In",
        denominator="Second Serve Attempts",
    ),
    "First Serve In": MetricDefinition(
        key="First Serve In",
        label="1st Serves In",
        description="Count of first serves that landed in.",
        kind="count",
    ),
    "Second Serve In": MetricDefinition(
        key="Second Serve In",
        label="2nd Serves In",
        description="Count of second serves that landed in.",
        kind="count",
    ),
    "First Serve Attempts": MetricDefinition(
        key="First Serve Attempts",
        label="1st Serve Attempts",
        description="Count of first-serve attempts.",
        kind="count",
    ),
    "Second Serve Attempts": MetricDefinition(
        key="Second Serve Attempts",
        label="2nd Serve Attempts",
        description="Count of second-serve attempts.",
        kind="count",
    ),
    "First Serve Wins": MetricDefinition(
        key="First Serve Wins",
        label="1st Serve Wins",
        description="Points won when the point was played on a first serve.",
        kind="count",
    ),
    "Second Serve Wins": MetricDefinition(
        key="Second Serve Wins",
        label="2nd Serve Wins",
        description="Points won when the point was played on a second serve.",
        kind="count",
    ),
    "First Serve Win %": MetricDefinition(
        key="First Serve Win %",
        label="1st Serve Win %",
        description="Point win rate on first-serve points.",
        kind="percent",
        chart_group="serve",
        timeline=True,
        numerator="First Serve Wins",
        denominator="First Serve Attempts",
    ),
    "Second Serve Win %": MetricDefinition(
        key="Second Serve Win %",
        label="2nd Serve Win %",
        description="Point win rate on second-serve points.",
        kind="percent",
        chart_group="serve",
        timeline=True,
        numerator="Second Serve Wins",
        denominator="Second Serve Attempts",
    ),
    "Double Faults": MetricDefinition(
        key="Double Faults",
        label="Double Faults",
        description="Count of points lost on a missed second serve.",
        kind="count",
        chart_group="serve",
    ),
    "Double Fault Rate": MetricDefinition(
        key="Double Fault Rate",
        label="Double Fault Rate",
        description="Double faults divided by second-serve attempts.",
        kind="percent",
        chart_group="serve",
        timeline=True,
        numerator="Double Faults",
        denominator="Second Serve Attempts",
    ),
    "First Return In": MetricDefinition(
        key="First Return In",
        label="1st Returns In",
        description="Count of first returns that landed in.",
        kind="count",
    ),
    "Second Return In": MetricDefinition(
        key="Second Return In",
        label="2nd Returns In",
        description="Count of second returns that landed in.",
        kind="count",
    ),
    "First Return Attempts": MetricDefinition(
        key="First Return Attempts",
        label="1st Return Attempts",
        description="Count of first-return attempts.",
        kind="count",
    ),
    "Second Return Attempts": MetricDefinition(
        key="Second Return Attempts",
        label="2nd Return Attempts",
        description="Count of second-return attempts.",
        kind="count",
    ),
    "First Return Wins": MetricDefinition(
        key="First Return Wins",
        label="1st Return Wins",
        description="Points won when the first return was put in play.",
        kind="count",
    ),
    "Second Return Wins": MetricDefinition(
        key="Second Return Wins",
        label="2nd Return Wins",
        description="Points won when the second return was put in play.",
        kind="count",
    ),
    "First Return In %": MetricDefinition(
        key="First Return In %",
        label="1st Return In %",
        description="Share of first returns that landed in.",
        kind="percent",
        chart_group="return",
        timeline=True,
        numerator="First Return In",
        denominator="First Return Attempts",
    ),
    "Second Return In %": MetricDefinition(
        key="Second Return In %",
        label="2nd Return In %",
        description="Share of second returns that landed in.",
        kind="percent",
        chart_group="return",
        timeline=True,
        numerator="Second Return In",
        denominator="Second Return Attempts",
    ),
    "First Return Win %": MetricDefinition(
        key="First Return Win %",
        label="1st Return Win %",
        description="Point win rate on first-return points.",
        kind="percent",
        chart_group="return",
        timeline=True,
        numerator="First Return Wins",
        denominator="First Return In",
    ),
    "Second Return Win %": MetricDefinition(
        key="Second Return Win %",
        label="2nd Return Win %",
        description="Point win rate on second-return points.",
        kind="percent",
        chart_group="return",
        timeline=True,
        numerator="Second Return Wins",
        denominator="Second Return In",
    ),
    "Serve +1 Attempts": MetricDefinition(
        key="Serve +1 Attempts",
        label="Serve +1 Attempts",
        description="Times the server hit the shot immediately after the return.",
        kind="count",
    ),
    "Serve +1 In": MetricDefinition(
        key="Serve +1 In",
        label="Serve +1 In",
        description="Serve +1 shots that landed in.",
        kind="count",
    ),
    "Serve +1 Wins": MetricDefinition(
        key="Serve +1 Wins",
        label="Serve +1 Wins",
        description="Points won when the player played a serve +1 shot.",
        kind="count",
    ),
    "Serve +1 In %": MetricDefinition(
        key="Serve +1 In %",
        label="Serve +1 In %",
        description="Share of serve +1 shots that landed in.",
        kind="percent",
        chart_group="transition",
        timeline=True,
        numerator="Serve +1 In",
        denominator="Serve +1 Attempts",
    ),
    "Serve +1 Win %": MetricDefinition(
        key="Serve +1 Win %",
        label="Serve +1 Win %",
        description="Point win rate when the player reached a serve +1 shot.",
        kind="percent",
        chart_group="transition",
        timeline=True,
        numerator="Serve +1 Wins",
        denominator="Serve +1 Attempts",
    ),
    "Return +1 Attempts": MetricDefinition(
        key="Return +1 Attempts",
        label="Return +1 Attempts",
        description="Times the returner hit the shot immediately after the serve +1 ball.",
        kind="count",
    ),
    "Return +1 In": MetricDefinition(
        key="Return +1 In",
        label="Return +1 In",
        description="Return +1 shots that landed in.",
        kind="count",
    ),
    "Return +1 Wins": MetricDefinition(
        key="Return +1 Wins",
        label="Return +1 Wins",
        description="Points won when the player played a return +1 shot.",
        kind="count",
    ),
    "Return +1 In %": MetricDefinition(
        key="Return +1 In %",
        label="Return +1 In %",
        description="Share of return +1 shots that landed in.",
        kind="percent",
        chart_group="transition",
        timeline=True,
        numerator="Return +1 In",
        denominator="Return +1 Attempts",
    ),
    "Return +1 Win %": MetricDefinition(
        key="Return +1 Win %",
        label="Return +1 Win %",
        description="Point win rate when the player reached a return +1 shot.",
        kind="percent",
        chart_group="transition",
        timeline=True,
        numerator="Return +1 Wins",
        denominator="Return +1 Attempts",
    ),
}

for bucket in POINT_LENGTH_BUCKETS:
    tag = _bucket_tag(bucket)
    METRIC_DEFINITIONS[f"{tag}_Wins"] = MetricDefinition(
        key=f"{tag}_Wins",
        label=f"{bucket} Wins",
        description=f"Points won in {bucket} rallies.",
        kind="count",
    )
    METRIC_DEFINITIONS[f"{tag}_Losses"] = MetricDefinition(
        key=f"{tag}_Losses",
        label=f"{bucket} Losses",
        description=f"Points lost in {bucket} rallies.",
        kind="count",
    )
    METRIC_DEFINITIONS[f"{tag}_Total"] = MetricDefinition(
        key=f"{tag}_Total",
        label=f"{bucket} Total",
        description=f"Total points in {bucket} rallies.",
        kind="count",
    )
    METRIC_DEFINITIONS[f"{tag}_Win%"] = MetricDefinition(
        key=f"{tag}_Win%",
        label=f"{bucket} Win %",
        description=f"Win rate in {bucket} rallies.",
        kind="percent",
        chart_group="rally",
        timeline=True,
        numerator=f"{tag}_Wins",
        denominator=f"{tag}_Total",
    )


SUMMARY_RATE_COMPONENTS = {
    key: (definition.numerator, definition.denominator)
    for key, definition in METRIC_DEFINITIONS.items()
    if definition.numerator and definition.denominator
}

SUMMARY_COUNT_KEYS = sorted(
    {
        definition.key
        for definition in METRIC_DEFINITIONS.values()
        if definition.kind == "count"
    }
)

SERVE_TABLE_KEYS = [
    "First Serve In",
    "First Serve Attempts",
    "Overall First Serve %",
    "Second Serve In",
    "Second Serve Attempts",
    "Overall Second Serve %",
    "First Serve Wins",
    "Second Serve Wins",
    "First Serve Win %",
    "Second Serve Win %",
    "Double Faults",
    "Double Fault Rate",
]

RETURN_TABLE_KEYS = [
    "First Return In",
    "First Return Attempts",
    "First Return In %",
    "Second Return In",
    "Second Return Attempts",
    "Second Return In %",
    "First Return Wins",
    "Second Return Wins",
    "First Return Win %",
    "Second Return Win %",
]

TRANSITION_TABLE_KEYS = [
    "Serve +1 Attempts",
    "Serve +1 In",
    "Serve +1 In %",
    "Serve +1 Wins",
    "Serve +1 Win %",
    "Return +1 Attempts",
    "Return +1 In",
    "Return +1 In %",
    "Return +1 Wins",
    "Return +1 Win %",
]

KEY_METRIC_KEYS = [
    "First Serve Win %",
    "Second Serve Win %",
    "Overall First Serve %",
    "Overall Second Serve %",
    "First Return Win %",
    "Second Return Win %",
    "Serve +1 Win %",
    "Return +1 Win %",
    "Double Fault Rate",
]

TIMELINE_METRIC_KEYS = [
    key for key, definition in METRIC_DEFINITIONS.items() if definition.timeline
]
