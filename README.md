# LeopardWeb Course Catalog Fetcher

A Python script to fetch all courses for a given semester from WIT's LeopardWeb system.

## Requirements

- Python 3.7 or higher
- `requests` library
- `openpyxl` library (for Excel output)
- `tqdm` library (for progress bars)
- `colorama` library (for colored output)

## Installation

Install Python dependencies:
```bash
pip install -r requirements.txt
```

Or install libraries directly:
```bash
pip install requests openpyxl tqdm colorama
```

## Usage

### List Available Terms

To see all available academic terms:
```bash
python leopardweb_courses.py --list-terms
```

Example output:
```
Available Terms:
------------------------------------------------------------
202610     Fall 2025
202510     Spring 2025
202410     Fall 2024
```

### Fetch Courses for a Term

To fetch all courses for a specific term (defaults to Excel format):
```bash
python leopardweb_courses.py <term_code>
```

Example:
```bash
python leopardweb_courses.py 202510
```

This will:
1. Initialize a session with LeopardWeb
2. Fetch all courses for the specified term
3. Fetch detailed meeting times and faculty information for each course (with progress bar)
4. Save the results to `courses_202510.xlsx`

**Note:** The script fetches detailed information for every course by default, which may take a few minutes for large course catalogs (e.g., 500+ courses). A progress bar shows real-time status.

### Output Formats

Choose between Excel (default), CSV, or JSON:

```bash
# Excel format (default) - formatted spreadsheet
python leopardweb_courses.py 202510

# CSV format - plain text, comma-separated
python leopardweb_courses.py 202510 --format csv

# JSON format - raw API data with full details
python leopardweb_courses.py 202510 --format json
```

### Custom Output File

To specify a custom output filename:
```bash
python leopardweb_courses.py 202510 -o spring_2025_courses.xlsx
python leopardweb_courses.py 202510 --format csv -o courses.csv
```

### Quick Mode (Skip Detailed Fetch)

For faster execution without detailed meeting times:
```bash
python leopardweb_courses.py 202510 --quick
```

### Quiet Mode

To suppress progress messages:
```bash
python leopardweb_courses.py 202510 -q
```

## Output Format Details

### Excel/CSV Output
The Excel and CSV formats include the following columns:
- CRN (Course Reference Number)
- Subject (e.g., CS, MATH)
- Course Number
- Section
- Title
- Credit Hours
- Schedule Type (Lecture, Lab, etc.)
- Instructional Method
- Faculty (professor names)
- Meeting Days (e.g., MWF, TR)
- Meeting Times (e.g., 09:00-09:50)
- Location (Building and room)
- Campus
- Enrollment Current/Max/Available
- Waitlist Current/Max

Excel files include:
- Formatted header row (blue background, white text)
- Frozen header row for easy scrolling
- Auto-adjusted column widths
- Text wrapping for long content

### JSON Output
The JSON format preserves the complete raw API response:

```json
{
  "term": "202510",
  "total_count": 450,
  "courses": [
    {
      "courseReferenceNumber": "12345",
      "subject": "CS",
      "courseNumber": "101",
      "courseTitle": "Introduction to Computer Science",
      "creditHours": 3,
      "faculty": [...],
      "meetingsFaculty": [...],
      ...
    }
  ]
}
```

## How It Works

The script replicates the functionality of the Rails backend service:

1. **Session Initialization**: Creates a JSESSIONID cookie by posting term selection to LeopardWeb
2. **Pagination**: Fetches courses in batches (500 per page) to handle large course catalogs
3. **Data Export**: Saves course data in your chosen format

## Troubleshooting

### Import Error
If you see `ModuleNotFoundError`:
```bash
pip install requests openpyxl
```

### Connection Error
If the script fails to connect:
- Check your internet connection
- Verify that https://selfservice.wit.edu is accessible
- The LeopardWeb system may be temporarily down

### No Courses Found
- Verify the term code is correct using `--list-terms`
- Some terms may not have courses published yet

## Examples

```bash
# List all available terms
python leopardweb_courses.py --list-terms

# Fetch Spring 2025 courses as Excel (default)
python leopardweb_courses.py 202510

# Fetch as CSV
python leopardweb_courses.py 202510 --format csv

# Fetch Fall 2025 courses with custom filename
python leopardweb_courses.py 202610 -o fall_2025.xlsx

# Quiet mode (no progress messages)
python leopardweb_courses.py 202510 -q --format csv

# Quick mode (skip detailed fetch for faster execution)
python leopardweb_courses.py 202510 --quick
```

## Technical Details

This script is based on the Ruby implementation from the [WIT Calendar Backend](https://github.com/WITCodingClub/calendar-backend) project:
- Source: [`app/services/leopard_web_service.rb`](https://github.com/WITCodingClub/calendar-backend/blob/main/app/services/leopard_web_service.rb)
- Uses the same API endpoints and authentication flow
- Maintains session cookies for authenticated requests
- Handles pagination automatically

## Author & License

Copyright © 2025 Jasper Mayone

This script is provided as-is for educational and research purposes.
