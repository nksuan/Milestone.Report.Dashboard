# Milestone Report Dashboard

Generate a self-contained HTML dashboard and a PowerPoint presentation from a
plain-text milestone (MV) report prompt.

## Setup

Python 3.10 or newer is required.

```bash
python -m pip install -r requirements.txt
```

## Report format

Create a text file using the following case-insensitive fields. Repeat `Metric`,
`Milestone`, and `Risk` lines as needed.

```text
Title: Q3 Platform Migration
Period: July–September 2026
Summary: Migration is progressing to plan.
Metric: Services migrated | 18 | 24 | On track
Milestone: Complete pilot | Ada | 2026-08-15 | Completed | 100
Milestone: Production cutover | Sam | 2026-09-20 | At risk | 65
Risk: Vendor access may delay the cutover.
```

Metric columns are `name | value | target | status`. Milestone columns are
`name | owner | due date | status | progress percentage`. Unlabelled text is
included in the executive summary, so short free-form prompts are also valid.

## Generate dashboards

```bash
python -m milestone_dashboard report.txt --output output
```

The command creates `output/dashboard.html` and `output/dashboard.pptx`. Omit
the file name to read the report from standard input:

```bash
cat report.txt | python -m milestone_dashboard --output output
```

## Tests

```bash
python -m unittest discover -s tests -v
```
