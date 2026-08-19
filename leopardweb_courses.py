#!/usr/bin/env python3
"""
WIT Course Catalog Fetcher

Fetches course schedule data for a given semester from the WIT Calendar
feed at calendar.witcc.dev.

Earlier versions of this script scraped LeopardWeb directly. They no longer
do. The calendar system already ingests LeopardWeb on a schedule, cleans the
result, and joins it to room and faculty records, so reading its feed is both
faster and more reliable than scraping. One HTTP request now replaces the
several hundred the scraper used to make.

Author: Jasper Mayone
Copyright (c) 2025 Jasper Mayone
Source: https://github.com/WITCodingClub/calendar-backend

Usage:
    python leopardweb_courses.py <term_code>
    python leopardweb_courses.py 202710              # Fall 2026 (Excel)
    python leopardweb_courses.py 202710 --format csv # CSV output
    python leopardweb_courses.py --list-terms        # Show available terms

Output:
    Saves courses to courses_{term_code}.xlsx (default), .csv, or .json
"""

import argparse
import csv
import io
import json
import sys
from typing import Dict, List, Optional

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from colorama import Fore, Style, init

# Initialize colorama for cross-platform colored output
init(autoreset=True)

DEFAULT_BASE_URL = "https://calendar.witcc.dev"

# Feed columns renamed for the spreadsheet. Order here is column order out.
COLUMN_LABELS = {
    "term": "Term",
    "crn": "CRN",
    "subject": "Subject",
    "course_number": "Course Number",
    "section_number": "Section",
    "title": "Title",
    "credit_hours": "Credit Hours",
    "schedule_type": "Schedule Type",
    "status": "Status",
    "faculty": "Faculty",
    "day": "Day",
    "begin_time": "Begin Time",
    "end_time": "End Time",
    "meeting_type": "Meeting Type",
    "building": "Building",
    "building_name": "Building Name",
    "room_number": "Room",
    "room": "Location",
    "seats_capacity": "Enrollment Max",
    "enrollment_current": "Enrollment Current",
    "seats_available": "Seats Available",
}


class CalendarFeedError(RuntimeError):
    """Raised when the calendar feed cannot be read."""


class CalendarFeedClient:
    """Client for the WIT Calendar public CSV reports."""

    def __init__(self, base_url: str = DEFAULT_BASE_URL, timeout: int = 60):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers["User-Agent"] = "leopardweb-python"

    def _get_csv(self, path: str, params: Optional[Dict] = None) -> List[Dict[str, str]]:
        url = f"{self.base_url}{path}"
        try:
            response = self.session.get(url, params=params, timeout=self.timeout)
            response.raise_for_status()
        except requests.RequestException as e:
            raise CalendarFeedError(f"Could not read {url}: {e}") from e

        # The feed always answers 200 with a CSV body. An HTML body means a
        # proxy or error page got in the way, which would otherwise parse as
        # a single nonsense row.
        content_type = response.headers.get("Content-Type", "")
        if "csv" not in content_type:
            raise CalendarFeedError(
                f"Expected CSV from {url}, got {content_type or 'an unknown type'}"
            )

        return list(csv.DictReader(io.StringIO(response.text)))

    def get_available_terms(self) -> List[Dict]:
        """Fetch the terms that have schedule data."""
        return [
            {
                "code": row["term_uid"],
                "description": row["term"],
                "meeting_times": row.get("meeting_times", ""),
                "start_date": row.get("start_date", ""),
                "end_date": row.get("end_date", ""),
            }
            for row in self._get_csv("/reports/terms")
        ]

    def get_meeting_times(self, term: str) -> List[Dict[str, str]]:
        """Fetch every scheduled meeting time for a term."""
        rows = self._get_csv("/reports/meeting_times", params={"term_uid": term})
        if not rows:
            raise CalendarFeedError(
                f"No schedule data for term {term}. "
                f"Run with --list-terms to see the terms that have data."
            )
        return rows


def to_output_row(row: Dict[str, str]) -> Dict:
    """Rename and reorder feed columns for tabular output."""
    out = {}
    for key, label in COLUMN_LABELS.items():
        value = row.get(key, "")
        # Keep numbers numeric so Excel sorts and sums them correctly.
        if key in ("crn", "credit_hours", "seats_capacity",
                   "seats_available", "enrollment_current"):
            out[label] = int(value) if value not in ("", None) else ""
        else:
            out[label] = value
    return out


def save_as_excel(rows: List[Dict], term: str, output_file: str, verbose: bool = True):
    """Save rows to an Excel workbook with a formatted, frozen header."""
    if verbose:
        print(f"{Fore.CYAN}📊 Creating Excel workbook...")

    wb = Workbook()
    ws = wb.active
    ws.title = f"Courses {term}"

    headers = list(COLUMN_LABELS.values())
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_num, row in enumerate(rows, 2):
        for col_num, header in enumerate(headers, 1):
            ws.cell(row=row_num, column=col_num, value=row.get(header, ""))

    # Size each column to its widest value, within reason.
    for col_num, header in enumerate(headers, 1):
        longest = max(
            [len(str(header))] + [len(str(r.get(header, ""))) for r in rows]
        )
        ws.column_dimensions[get_column_letter(col_num)].width = min(longest + 2, 50)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(output_file)

    if verbose:
        print(f"{Fore.GREEN}✓ Saved {Style.BRIGHT}{len(rows)}{Style.NORMAL} rows to {Style.BRIGHT}{output_file}")


def save_as_csv(rows: List[Dict], term: str, output_file: str, verbose: bool = True):
    """Save rows to a CSV file."""
    if verbose:
        print(f"{Fore.CYAN}📄 Creating CSV file...")

    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(COLUMN_LABELS.values()))
        writer.writeheader()
        writer.writerows(rows)

    if verbose:
        print(f"{Fore.GREEN}✓ Saved {Style.BRIGHT}{len(rows)}{Style.NORMAL} rows to {Style.BRIGHT}{output_file}")


def save_as_json(rows: List[Dict], term: str, output_file: str, verbose: bool = True):
    """Save rows to a JSON file."""
    if verbose:
        print(f"{Fore.CYAN}📝 Creating JSON file...")

    with open(output_file, "w") as f:
        json.dump({"term": term, "total_count": len(rows), "meeting_times": rows}, f, indent=2)

    if verbose:
        print(f"{Fore.GREEN}✓ Saved {Style.BRIGHT}{len(rows)}{Style.NORMAL} rows to {Style.BRIGHT}{output_file}")


def list_terms(base_url: str = DEFAULT_BASE_URL):
    """List all terms that have schedule data."""
    try:
        terms = CalendarFeedClient(base_url).get_available_terms()
    except CalendarFeedError as e:
        print(f"{Fore.RED}❌ {e}", file=sys.stderr)
        sys.exit(1)

    if not terms:
        print(f"{Fore.RED}No terms found", file=sys.stderr)
        return

    print(f"\n{Fore.CYAN}{Style.BRIGHT}Available Terms:")
    print(f"{Fore.CYAN}" + "-" * 60)
    for term in terms:
        print(f"{Fore.YELLOW}{term['code']:10} {Fore.WHITE}{term['description']:16} "
              f"{Fore.CYAN}{term['meeting_times']:>6} meeting times")
    print()


def fetch_courses(term: str, output_file: Optional[str] = None,
                  format: str = "excel", verbose: bool = True,
                  base_url: str = DEFAULT_BASE_URL):
    """Fetch a term's schedule from the feed and save it."""
    try:
        if verbose:
            print(f"{Fore.CYAN}🔎 Fetching term {Style.BRIGHT}{term}{Style.NORMAL} from {base_url}...")

        feed_rows = CalendarFeedClient(base_url).get_meeting_times(term)
        rows = [to_output_row(r) for r in feed_rows]

        if verbose:
            sections = len({r["CRN"] for r in rows})
            print(f"{Fore.GREEN}✓ Got {Style.BRIGHT}{len(rows)}{Style.NORMAL} meeting times "
                  f"across {Style.BRIGHT}{sections}{Style.NORMAL} sections")

        if not output_file:
            extensions = {"excel": ".xlsx", "csv": ".csv", "json": ".json"}
            output_file = f"courses_{term}{extensions.get(format, '.xlsx')}"

        if format == "excel":
            save_as_excel(rows, term, output_file, verbose)
        elif format == "csv":
            save_as_csv(rows, term, output_file, verbose)
        elif format == "json":
            save_as_json(rows, term, output_file, verbose)
        else:
            raise ValueError(f"Unsupported format: {format}")

    except (CalendarFeedError, ValueError) as e:
        print(f"{Fore.RED}❌ {e}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Fetch WIT course schedule data from calendar.witcc.dev",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # List terms that have data
  python leopardweb_courses.py --list-terms

  # Fetch Fall 2026 (Excel)
  python leopardweb_courses.py 202710

  # Fetch as CSV
  python leopardweb_courses.py 202710 --format csv

  # Fetch as JSON with a custom output file
  python leopardweb_courses.py 202710 --format json -o fall2026.json

Note:
  Each row is one meeting time in one room, so a course that meets three
  times a week has three rows. Group by CRN to get one row per section.
        """
    )

    parser.add_argument("term", nargs="?", help="Term code (e.g., 202710 for Fall 2026)")
    parser.add_argument("--list-terms", action="store_true", help="List all terms that have data")
    parser.add_argument("-f", "--format", choices=["excel", "csv", "json"],
                        default="excel", help="Output format (default: excel)")
    parser.add_argument("-o", "--output", help="Output filename (default: courses_{term}.{ext})")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL,
                        help=f"Calendar server to read from (default: {DEFAULT_BASE_URL})")
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress progress messages")

    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)

    if args.list_terms:
        list_terms(args.base_url)
        return

    if not args.term:
        print(f"{Fore.RED}Error: term code is required (or use --list-terms)", file=sys.stderr)
        parser.print_help()
        sys.exit(1)

    fetch_courses(args.term, args.output, args.format,
                  verbose=not args.quiet, base_url=args.base_url)


if __name__ == "__main__":
    main()
