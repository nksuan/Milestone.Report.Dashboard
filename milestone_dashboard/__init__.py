"""Generate HTML and PowerPoint dashboards from milestone report prompts."""

from .dashboard import generate_dashboard
from .report import MilestoneReport, parse_report

__all__ = ["MilestoneReport", "generate_dashboard", "parse_report"]
