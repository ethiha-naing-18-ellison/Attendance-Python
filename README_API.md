# Attendance Report Generator API

A FastAPI-based web service that generates Excel attendance reports from ZK.db SQLite database files.

## Features

- 📤 Upload ZK.db database files
- 📅 Specify exact date ranges (calendar dates in YYYY-MM-DD format)
- 🎉 Support for multiple public holidays (comma-separated) with automatic row highlighting
- 📊 Generate formatted Excel reports with attendance data
- 🎨 **Produces IDENTICAL output to the original saya.py script** with complete feature parity:
  - ✅ Complex SQL queries with all CTEs (punches_per_day, ranked_punches, final, with_flags, with_work_time, final_with_ot)
  - ✅ All header color coding: light yellow (#FFF2CC), orange (#FFCC99), blue (#ADD8E6), purple (#E6E6FA), green (#E2EFDA)
  - ✅ Sunday row highlighting with yellow background (#FFFF00)
  - ✅ Late/Early clock-in/out detection with red highlighting (#FFCCCB)
  - ✅ Suspicious punch pattern detection with orange highlighting (#FFA500)
  - ✅ Suspicious early clock-in detection with bright red font
  - ✅ Employee grouping with comprehensive totals calculations
  - ✅ Outline levels for data grouping (collapsed by default)
  - ✅ Frozen panes for header visibility
  - ✅ Complete business logic for penalties, overtime, night shifts, allowances
  - ✅ Time format conversions and decimal calculations
  - ✅ Borders, fonts (Tahoma), and all styling elements
  - ✅ Date range display under subtitle (e.g., "2025-06-01 to 2025-06-30")
  - ✅ Public holiday support with blue-gray row highlighting (#B0C4DE)
  - ✅ **Public Holiday OT Redistribution**: OT1 and OT2 values automatically moved to OT3 on public holidays
- ⚡ Fast and efficient processing

## Installation

1. **Install Python dependencies:**

```bash
pip install -r requirements.txt
```

2. **Run the API server:**

```bash
python main.py
```

Or directly with uvicorn:

```bash
uvicorn attendance_api:app --host 0.0.0.0 --port 8000 --reload
```

## Usage

### Starting the Server

Run the server using:

```bash
python main.py
```

The API will be available at:

- **Main API**: http://localhost:8000
- **Interactive Documentation**: http://localhost:8000/docs
- **Alternative Documentation**: http://localhost:8000/redoc

### API Endpoints

#### POST `/generate-attendance-report`

Generate an Excel attendance report from a ZK.db file.

**Parameters:**

- `db_file` (File): ZK.db SQLite database file
- `start_date` (Form): Start date in YYYY-MM-DD format (e.g., "2025-06-01")
- `end_date` (Form, Optional): End date in YYYY-MM-DD format (defaults to start_date if not provided)
- `public_holidays` (Form, Optional): Public holiday dates in YYYY-MM-DD format, comma-separated (e.g., "2025-06-01,2025-06-15")

**Example using cURL:**

```bash
curl -X POST "http://localhost:8000/generate-attendance-report" \
  -F "db_file=@/path/to/your/ZK.db" \
  -F "start_date=2025-06-01" \
  -F "end_date=2025-06-30" \
  -F "public_holidays=2025-06-05,2025-06-15" \
  --output attendance_report.xlsx
```

**Example using Python requests:**

```python
import requests

url = "http://localhost:8000/generate-attendance-report"
files = {'db_file': open('ZK.db', 'rb')}
data = {
    'start_date': '2025-06-01',
    'end_date': '2025-06-30',
    'public_holidays': '2025-06-05,2025-06-15'  # Comma-separated public holidays
}

response = requests.post(url, files=files, data=data)

if response.status_code == 200:
    with open('attendance_report.xlsx', 'wb') as f:
        f.write(response.content)
    print("Report generated successfully!")
else:
    print(f"Error: {response.json()}")
```

### Punch Time Smart Adjustment (1-Hour Gap Rule)

The system now includes intelligent punch time adjustment to handle accidental quick double-punches for BOTH Clock-In/Clock-Out AND In/Out pairs:

**How it works:**

- **Clock-In** is always the first punch of the day
- **Clock-Out** must be at least 1 hour after Clock-In, otherwise skip to next punch
- **In** and **Out** must be at least 1 hour apart, otherwise skip the Out to next punch
- All adjustments cascade to maintain logical sequence

**Example scenarios:**

**Scenario 1 - Normal Case (no adjustment needed):**

```
Raw punches: 8:00 AM, 12:00 PM, 1:00 PM, 5:00 PM
Result: Clock-In: 8:00 AM, Clock-Out: 12:00 PM, In: 1:00 PM, Out: 5:00 PM
(4-hour gap between Clock-In and Clock-Out ✅)
```

**Scenario 2 - Clock-Out too close (adjustment applied):**

```
Raw punches: 8:00 AM, 8:03 AM, 12:00 PM, 1:00 PM, 5:00 PM
Before: Clock-In: 8:00 AM, Clock-Out: 8:03 AM (❌ only 3 minutes!)
After:  Clock-In: 8:00 AM, Clock-Out: 12:00 PM, In: 1:00 PM, Out: 5:00 PM
```

**Scenario 3 - In/Out too close (adjustment applied):**

```
Raw punches: 8:00 AM, 12:00 PM, 2:00 PM, 2:02 PM, 6:00 PM
Before: In: 2:00 PM, Out: 2:02 PM (❌ only 2 minutes!)
After:  Clock-In: 8:00 AM, Clock-Out: 12:00 PM, In: 2:00 PM, Out: 6:00 PM
```

**Scenario 4 - Multiple adjustments:**

```
Raw punches: 8:00 AM, 8:03 AM, 12:00 PM, 2:00 PM, 2:02 PM, 6:00 PM
Result: Clock-In: 8:00 AM, Clock-Out: 12:00 PM, In: 2:00 PM, Out: 6:00 PM
(Both Clock-Out and Out were adjusted)
```

This ensures realistic work patterns for all punch types while preserving all punch data.

### Using the Web Interface

1. Open http://localhost:8000 in your browser for the user-friendly web interface
2. Upload your ZK.db file using the file selector
3. Enter the start date (and optionally end date) using the date pickers
4. Optionally enter public holiday dates (comma-separated format, e.g., "2025-06-05,2025-06-15")
5. Click the **"Generate Report"** button
6. Wait for processing to complete (progress bar will show status)
7. Click the **"Download Report"** button when it appears

### Using the Interactive API Documentation

1. Open http://localhost:8000/docs in your browser
2. Click on the **POST /generate-attendance-report** endpoint
3. Click "Try it out"
4. Upload your ZK.db file
5. Enter the start date (and optionally end date) in YYYY-MM-DD format
6. Optionally enter public holiday dates (comma-separated format, e.g., "2025-06-05,2025-06-15")
7. Click "Execute"
8. Download the generated Excel file from the response

## Date Format

- **Format**: YYYY-MM-DD (e.g., "2025-06-01" for June 1st, 2025)
- **start_date**: Required - The starting date for the report
- **end_date**: Optional - The ending date for the report (defaults to start_date)
- **public_holidays**: Optional - Multiple public holiday dates in YYYY-MM-DD format, comma-separated (e.g., "2025-06-05,2025-06-15")
- **Date Range Display**: The exact date range will be displayed under the "Monthly Statement Report" subtitle in the Excel file (e.g., "2025-06-01 to 2025-06-30")
- **Public Holiday Highlighting**: Rows matching public holiday dates will be highlighted with blue-gray background (#B0C4DE)
- **Public Holiday OT Redistribution**: On public holiday dates, any OT1 and OT2 values are automatically moved to OT3. This applies to both individual rows and TOTAL calculations.

### Public Holiday OT Redistribution Example

**Normal Day:**

- OT1: 2.5 hours
- OT2: 0.0 hours
- OT3: 1.0 hours

**Same Day if it's a Public Holiday:**

- OT1: 0.0 hours (moved to OT3)
- OT2: 0.0 hours (moved to OT3)
- OT3: 3.5 hours (2.5 + 0.0 + 1.0)

## Response

- **Success**: Returns an Excel file (.xlsx) with the attendance report
- **Error**: Returns JSON with error details

## Error Handling

The API handles various error scenarios:

- Invalid date formats
- Invalid file types (non-.db files)
- Database connection errors
- No data found for specified date range
- Internal server errors

## Security Note

This API is designed for internal use. In production environments, consider adding:

- Authentication and authorization
- Rate limiting
- File size limits
- Input validation and sanitization
- HTTPS encryption

## Dependencies

- FastAPI: Web framework
- Uvicorn: ASGI server
- Pandas: Data manipulation
- OpenPyXL: Excel file generation
- SQLite3: Database connectivity (built-in)

## File Structure

```
├── attendance_api.py    # Main FastAPI application
├── main.py          # Simple script to run the API
├── requirements.txt    # Python dependencies
├── README_API.md      # This documentation
└── saya.py            # Original script (for reference)
multi_employee_attendance_converter.py         # Convery Exel to Json File (For Reference)
```
