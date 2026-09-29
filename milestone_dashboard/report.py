from dataclasses import dataclass, field
from datetime import date


@dataclass(frozen=True)
class Metric:
    name: str
    value: str
    target: str = ""
    status: str = "On track"


@dataclass(frozen=True)
class Milestone:
    name: str
    owner: str = "Unassigned"
    due: str = ""
    status: str = "Not started"
    progress: int = 0


@dataclass(frozen=True)
class MilestoneReport:
    title: str = "Milestone Report"
    period: str = field(default_factory=lambda: date.today().isoformat())
    summary: str = ""
    metrics: tuple[Metric, ...] = ()
    milestones: tuple[Milestone, ...] = ()
    risks: tuple[str, ...] = ()


def _parts(value: str, count: int) -> list[str]:
    values = [part.strip() for part in value.split("|")]
    return (values + [""] * count)[:count]


def _progress(value: str) -> int:
    try:
        return max(0, min(100, int(value.rstrip("%").strip())))
    except ValueError:
        return 0


def parse_report(prompt: str) -> MilestoneReport:
    """Parse a line-oriented user prompt into a milestone report.

    Supported fields are title, period, summary, metric, milestone, and risk.
    Fields are case-insensitive. Unlabelled text is included in the summary.
    """
    if not prompt or not prompt.strip():
        raise ValueError("The report prompt cannot be empty.")

    title = "Milestone Report"
    period = date.today().isoformat()
    summary_parts: list[str] = []
    metrics: list[Metric] = []
    milestones: list[Milestone] = []
    risks: list[str] = []

    for raw_line in prompt.splitlines():
        line = raw_line.strip().lstrip("-* ").strip()
        if not line:
            continue

        key, separator, value = line.partition(":")
        if not separator:
            summary_parts.append(line)
            continue

        key = key.strip().lower()
        value = value.strip()
        if key == "title" and value:
            title = value
        elif key == "period" and value:
            period = value
        elif key == "summary" and value:
            summary_parts.append(value)
        elif key == "metric" and value:
            name, metric_value, target, status = _parts(value, 4)
            if name and metric_value:
                metrics.append(Metric(name, metric_value, target, status or "On track"))
        elif key == "milestone" and value:
            name, owner, due, status, progress = _parts(value, 5)
            if name:
                milestones.append(
                    Milestone(
                        name,
                        owner or "Unassigned",
                        due,
                        status or "Not started",
                        _progress(progress),
                    )
                )
        elif key in {"risk", "risks"} and value:
            risks.append(value)
        else:
            summary_parts.append(line)

    return MilestoneReport(
        title=title,
        period=period,
        summary=" ".join(summary_parts),
        metrics=tuple(metrics),
        milestones=tuple(milestones),
        risks=tuple(risks),
    )
