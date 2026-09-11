# WIT Course Catalog Fetcher

A Python script to fetch the course schedule for a WIT semester and save it as
Excel, CSV, or JSON.

The script reads the public schedule feed at
[calendar.witcc.dev](https://calendar.witcc.dev). It does not scrape
LeopardWeb.

## Why the feed and not LeopardWeb

Earlier versions of this script scraped LeopardWeb directly. The WIT Calendar
system already does that scrape on a schedule, cleans the result, and joins it
to room and faculty records. Reading its feed gives you:

- **Speed.** One HTTP request, not several hundred. A full term takes about a
  second.
- **Clean rooms.** Building abbreviations and room numbers come as separate
  columns, plus one joined `Location` label. LeopardWeb gives you one text blob.
- **No session handling.** There is no JSESSIONID, no pagination, and no rate
  limit to respect.
- **No load on LeopardWeb.** Many people running the old script hit the
  registrar's server many times each.

## Requirements

- Python 3.7 or higher
- `requests` library
- `openpyxl` library (for Excel output)
- `colorama` library (for colored output)

## Installation

Install Python dependencies:
```bash
pip install -r requirements.txt
```

Or install libraries directly:
```bash
pip install requests openpyxl colorama
```

## Usage

### List Available Terms

To see the terms that have schedule data:
```bash
python leopardweb_courses.py --list-terms
```

Example output:
```
Available Terms:
------------------------------------------------------------
202610     Fall 2025          2061 meeting times
202620     Spring 2026        1898 meeting times
202710     Fall 2026          2120 meeting times
```

Only terms with schedule data appear here. A term that the calendar has not
ingested yet is not listed.

### Fetch Courses for a Term

To fetch the schedule for a term (defaults to Excel format):
```bash
python leopardweb_courses.py <term_code>
```

Example:
```bash
python leopardweb_courses.py 202710
```

This saves the result to `courses_202710.xlsx`.

### Output Formats

Choose between Excel (default), CSV, or JSON:

```bash
# Excel format (default) - formatted spreadsheet
python leopardweb_courses.py 202710

# CSV format - plain text, comma-separated
python leopardweb_courses.py 202710 --format csv

# JSON format
python leopardweb_courses.py 202710 --format json
```

### Custom Output File

To specify a custom output filename:
```bash
python leopardweb_courses.py 202710 -o fall_2026_courses.xlsx
python leopardweb_courses.py 202710 --format csv -o courses.csv
```

### Quiet Mode

To suppress progress messages:
```bash
python leopardweb_courses.py 202710 -q
```

### One Row Per Meeting Day

To get one row per meeting day in one room, instead of one row per section:
```bash
python leopardweb_courses.py 202710 --by-meeting
```

### Custom Server

To read from a different calendar server (for example, a local development
instance):
```bash
python leopardweb_courses.py 202710 --base-url http://localhost:3000
```

## Output Format Details

### Row grain

**One row is one section.** A section appears once, the way it appears on a
schedule. Meeting days are collapsed into Banner day codes, so a Tuesday and
Thursday lecture reads as `TR`.

```
CRN    Subject  Course Number  Section  Meeting Days  Meeting Times  Location
16036  MATH     2300           3        TR            10:00-11:45    WENTW 214
17309  MATH     1525           1A       MW            08:00-09:15    WENTW 206
```

Day codes start on Monday. `R` is Thursday and `U` is Sunday, because `T` and
`S` are already taken.

| Code | M | T | W | R | F | S | U |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Day | Mon | Tue | Wed | Thu | Fri | Sat | Sun |

Most sections meet on one pattern. A section with more than one pattern, such
as a lecture that also has a Friday lab in another room, lists one part per
pattern separated by `; `. The parts line up across every meeting column:

```
Meeting Days  Meeting Times              Location
MW; F         09:00-10:15; 13:00-14:50   ANX 305; DOB 005
```

### One row per meeting day

Pass `--by-meeting` for the other shape: **one row is one meeting time in one
room.** A course that meets Monday, Wednesday, and Friday has three rows. A
course booked into two rooms at the same hour has one row per room.

```bash
python leopardweb_courses.py 202710 --by-meeting
```

Use it for room and hour questions, such as "what is in Annex 305 on Tuesday
at 10:00". You can filter instead of parsing a day string.

Use the default for a course list. In the `--by-meeting` shape a reader
scanning for courses sees each one more than once, and reads the extra rows as
duplicates.

The `Meeting Count` column on a section row says how many `--by-meeting` rows
collapsed into it, so the two shapes reconcile.

### Columns

| Column | Notes |
| --- | --- |
| Term | For example, `Fall 2026` |
| CRN | Course Reference Number |
| Subject | For example, `CS`, `MATH` |
| Course Number | |
| Section | |
| Title | |
| Credit Hours | |
| Schedule Type | `lecture`, `lab`, and so on |
| Status | `active` or `cancelled` |
| Faculty | Team-taught sections list every teacher, comma separated |
| Meeting Days | Banner day codes, for example `TR` |
| Meeting Times | 24-hour `HH:MM-HH:MM` |
| Meeting Type | The meeting's own type, which can differ from the section's |
| Location | Building and room joined, for example `ANX 305` |
| Room Capacity | Largest room the section is scheduled into. Blank when unknown |
| Enrollment Max | Section seat cap |
| Enrollment Current | Seats taken |
| Seats Available | Seats left |
| Meeting Count | How many `--by-meeting` rows this section collapses |

With `--by-meeting`, the four meeting columns are replaced by these:

| Column | Notes |
| --- | --- |
| Day | `monday` through `sunday` |
| Begin Time | 24-hour `HH:MM` |
| End Time | 24-hour `HH:MM` |
| Meeting Type | The meeting's own type, which can differ from the section's |
| Building | Abbreviation, for example `ANX` |
| Building Name | Full name, for example `Annex` |
| Room | Room number, padded to 3 digits when numeric |
| Location | Building and room joined, for example `ANX 305` |
| Room Capacity | Capacity of that room. Blank when unknown |

Excel files include:
- Formatted header row (blue background, white text)
- Frozen header row for easy scrolling
- An auto-filter on every column
- Auto-adjusted column widths
- Numeric columns kept numeric, so sums and sorts work

### Columns that are gone

The old scraper produced two columns the feed does not have:

- **Instructional Method** — the calendar does not store this yet.
- **Campus** — the calendar does not store this yet.
- **Waitlist Current / Waitlist Max** — the calendar does not expose these in
  the public feed.

## How It Works

1. `GET /reports/terms` returns the terms that have schedule data.
2. `GET /reports/sections?term_uid=<term>` returns one row per section for that
   term as CSV. With `--by-meeting`, `GET /reports/meeting_times?term_uid=<term>`
   returns one row per meeting time instead.
3. The script renames the columns and writes your chosen format.

The feed does the collapsing, so this script and any other tool reading the
same report always agree on the shape.

All three endpoints are public, read-only, and cached for one hour. They carry
course schedule data only — never user data.

You can point any tool at them, not just this script. Power BI and Excel can
read them directly with a Web connector.

## Troubleshooting

### Import Error
If you see `ModuleNotFoundError`:
```bash
pip install -r requirements.txt
```

### Connection Error
If the script fails to connect:
- Check your internet connection.
- Check that https://calendar.witcc.dev is up.

### No Schedule Data For Term
- Check the term code with `--list-terms`.
- The calendar lists only the terms it has ingested. If a term is missing, the
  calendar has not imported it yet.

## Examples

```bash
# List the terms that have data
python leopardweb_courses.py --list-terms

# Fetch Fall 2026 as Excel (default)
python leopardweb_courses.py 202710

# Fetch as CSV
python leopardweb_courses.py 202710 --format csv

# Fetch with a custom filename
python leopardweb_courses.py 202710 -o fall_2026.xlsx

# Quiet mode (no progress messages)
python leopardweb_courses.py 202710 -q --format csv
```

## Technical Details

The upstream scrape lives in the [WIT Calendar
Backend](https://github.com/WITCodingClub/calendar-backend):
- Scraper: `app/services/leopard_web_service.rb`
- Feed: `app/controllers/reports_controller.rb`

## Author & License

Copyright © 2025 Jasper Mayone

This script is provided as-is for educational and research purposes.
