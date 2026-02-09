# LeopardWeb Course Catalog Fetcher

A Python script to fetch all courses for a given semester from WIT's LeopardWeb system.

## Requirements

- Python 3.7 or higher
- `requests` library

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

Or install the requests library directly:
```bash
pip install requests
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

To fetch all courses for a specific term:
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
3. Save the results to `courses_202510.json`

### Custom Output File

To specify a custom output filename:
```bash
python leopardweb_courses.py 202510 -o spring_2025_courses.json
```

### Quiet Mode

To suppress progress messages:
```bash
python leopardweb_courses.py 202510 -q
```

## Output Format

The script saves courses in JSON format:

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
    },
    ...
  ]
}
```

Each course object contains detailed information including:
- Course identification (CRN, subject, number, title)
- Credit hours
- Faculty information
- Meeting times and locations
- Enrollment information
- And more...

## How It Works

The script replicates the functionality of the Rails backend service:

1. **Session Initialization**: Creates a JSESSIONID cookie by posting term selection to LeopardWeb
2. **Pagination**: Fetches courses in batches (500 per page) to handle large course catalogs
3. **JSON Export**: Saves all course data to a JSON file for easy processing

## Troubleshooting

### Import Error
If you see `ModuleNotFoundError: No module named 'requests'`:
```bash
pip install requests
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

# Fetch Spring 2025 courses
python leopardweb_courses.py 202510

# Fetch Fall 2025 courses with custom filename
python leopardweb_courses.py 202610 -o fall_2025.json

# Quiet mode (no progress messages)
python leopardweb_courses.py 202510 -q
```

## Technical Details

This script is based on the Ruby implementation from the WIT Calendar Backend project:
- Source: `app/services/leopard_web_service.rb`
- Uses the same API endpoints and authentication flow
- Maintains session cookies for authenticated requests
- Handles pagination automatically

## License

This script is provided as-is for educational and research purposes.
