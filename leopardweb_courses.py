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

One row is one section. Pass --by-meeting for one row per meeting day in
one room, which is the shape to use for room and hour questions.

Usage:
    python leopardweb_courses.py <term_code>
    python leopardweb_courses.py 202710              # Fall 2026 (Excel)
    python leopardweb_courses.py 202710 --format csv # CSV output
    python leopardweb_courses.py 202710 --by-meeting # One row per meeting day
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

# One row per section. This is the default, and the grain most people expect:
# a section you take or teach appears once. Order here is column order out.
SECTION_COLUMN_LABELS = {
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
    "meeting_days": "Meeting Days",
    "meeting_times": "Meeting Times",
    "meeting_type": "Meeting Type",
    "room": "Location",
    "room_capacity": "Room Capacity",
    "seats_capacity": "Enrollment Max",
    "enrollment_current": "Enrollment Current",
    "seats_available": "Seats Available",
    "meeting_count": "Meeting Count",
}

# One row per meeting time in one room, for --by-meeting. Use this grain to
# ask room and hour questions without reading a day string.
MEETING_COLUMN_LABELS = {
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
    "room_capacity": "Room Capacity",
    "seats_capacity": "Enrollment Max",
    "enrollment_current": "Enrollment Current",
    "seats_available": "Seats Available",
}

# Columns Excel should hold as numbers, so sorting and summing work.
NUMERIC_KEYS = frozenset({
    "crn", "credit_hours", "seats_capacity", "seats_available",
    "enrollment_current", "room_capacity", "meeting_count",
})


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
        except requests.HTTPError as e:
            # A 404 on a report path means the server predates the report, not
            # that the term is wrong. Say so, because the two read alike.
            if e.response is not None and e.response.status_code == 404:
                raise CalendarFeedError(
                    f"{self.base_url} has no {path} report. That server is older "
                    f"than this script. Update the server, or use --by-meeting, "
                    f"which reads the older /reports/meeting_times report."
                ) from e
            raise CalendarFeedError(f"Could not read {url}: {e}") from e
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

    def get_sections(self, term: str) -> List[Dict[str, str]]:
        """Fetch one row per course section for a term."""
        return self._term_report("/reports/sections", term)

    def get_meeting_times(self, term: str) -> List[Dict[str, str]]:
        """Fetch every scheduled meeting time for a term."""
        return self._term_report("/reports/meeting_times", term)

    def _term_report(self, path: str, term: str) -> List[Dict[str, str]]:
        rows = self._get_csv(path, params={"term_uid": term})
        if not rows:
            raise CalendarFeedError(
                f"No schedule data for term {term}. "
                f"Run with --list-terms to see the terms that have data."
            )
        return rows


def to_output_row(row: Dict[str, str], labels: Dict[str, str]) -> Dict:
    """Rename and reorder feed columns for tabular output."""
    out = {}
    for key, label in labels.items():
        value = row.get(key, "")
        # Keep numbers numeric so Excel sorts and sums them correctly. A blank
        # cell stays blank rather than becoming a wrong 0.
        if key in NUMERIC_KEYS and value not in ("", None):
            try:
                out[label] = int(value)
            except ValueError:
                out[label] = value
        else:
            out[label] = value
    return out


def save_as_excel(rows: List[Dict], term: str, output_file: str,
                  labels: Dict[str, str], verbose: bool = True):
    """Save rows to an Excel workbook with a formatted, frozen header."""
    if verbose:
        print(f"{Fore.CYAN}📊 Creating Excel workbook...")

    wb = Workbook()
    ws = wb.active
    ws.title = f"Courses {term}"

    headers = list(labels.values())
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


def save_as_csv(rows: List[Dict], term: str, output_file: str,
                labels: Dict[str, str], verbose: bool = True):
    """Save rows to a CSV file."""
    if verbose:
        print(f"{Fore.CYAN}📄 Creating CSV file...")

    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(labels.values()))
        writer.writeheader()
        writer.writerows(rows)

    if verbose:
        print(f"{Fore.GREEN}✓ Saved {Style.BRIGHT}{len(rows)}{Style.NORMAL} rows to {Style.BRIGHT}{output_file}")


def save_as_json(rows: List[Dict], term: str, output_file: str,
                 key: str, verbose: bool = True):
    """Save rows to a JSON file."""
    if verbose:
        print(f"{Fore.CYAN}📝 Creating JSON file...")

    with open(output_file, "w") as f:
        json.dump({"term": term, "total_count": len(rows), key: rows}, f, indent=2)

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
                  base_url: str = DEFAULT_BASE_URL, by_meeting: bool = False):
    """Fetch a term's schedule from the feed and save it."""
    try:
        if verbose:
            print(f"{Fore.CYAN}🔎 Fetching term {Style.BRIGHT}{term}{Style.NORMAL} from {base_url}...")

        client = CalendarFeedClient(base_url)
        labels = MEETING_COLUMN_LABELS if by_meeting else SECTION_COLUMN_LABELS
        feed_rows = client.get_meeting_times(term) if by_meeting else client.get_sections(term)
        rows = [to_output_row(r, labels) for r in feed_rows]

        if verbose:
            sections = len({r["CRN"] for r in rows})
            if by_meeting:
                print(f"{Fore.GREEN}✓ Got {Style.BRIGHT}{len(rows)}{Style.NORMAL} meeting times "
                      f"across {Style.BRIGHT}{sections}{Style.NORMAL} sections")
            else:
                print(f"{Fore.GREEN}✓ Got {Style.BRIGHT}{sections}{Style.NORMAL} sections")

        if not output_file:
            extensions = {"excel": ".xlsx", "csv": ".csv", "json": ".json"}
            output_file = f"courses_{term}{extensions.get(format, '.xlsx')}"

        if format == "excel":
            save_as_excel(rows, term, output_file, labels, verbose)
        elif format == "csv":
            save_as_csv(rows, term, output_file, labels, verbose)
        elif format == "json":
            save_as_json(rows, term, output_file,
                         "meeting_times" if by_meeting else "sections", verbose)
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

  # One row per meeting day instead of one row per section
  python leopardweb_courses.py 202710 --by-meeting

Row grain:
  By default one row is one section, the way it appears on a schedule.
  Meeting days are collapsed into Banner day codes, so a Tuesday and
  Thursday lecture reads as "TR". R is Thursday and U is Sunday.

  --by-meeting gives one row per meeting day in one room instead. Use it
  for room and hour questions, such as what is in WENTW 206 on Tuesday at
  08:00. In that shape a section that meets twice a week has two rows.
        """
    )

    parser.add_argument("term", nargs="?", help="Term code (e.g., 202710 for Fall 2026)")
    parser.add_argument("--list-terms", action="store_true", help="List all terms that have data")
    parser.add_argument("-f", "--format", choices=["excel", "csv", "json"],
                        default="excel", help="Output format (default: excel)")
    parser.add_argument("-o", "--output", help="Output filename (default: courses_{term}.{ext})")
    parser.add_argument("--by-meeting", action="store_true",
                        help="One row per meeting day in one room, instead of one per section")
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
                  verbose=not args.quiet, base_url=args.base_url,
                  by_meeting=args.by_meeting)


if __name__ == "__main__":
    main()
