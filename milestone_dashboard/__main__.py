import argparse
import sys
from pathlib import Path

from .dashboard import generate_dashboard
from .report import parse_report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate HTML and PowerPoint dashboards from an MV report prompt."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="-",
        help="Prompt text file, or - to read from standard input (default: -).",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="output",
        help="Directory for dashboard.html and dashboard.pptx (default: output).",
    )
    args = parser.parse_args()

    prompt = sys.stdin.read() if args.input == "-" else Path(args.input).read_text(encoding="utf-8")
    try:
        report = parse_report(prompt)
    except ValueError as error:
        parser.error(str(error))

    html_path, powerpoint_path = generate_dashboard(report, args.output)
    print(f"Created {html_path}")
    print(f"Created {powerpoint_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
