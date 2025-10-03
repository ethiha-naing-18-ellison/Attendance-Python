# Data Sheet Implementation Summary

## What Was Changed

### File Modified: `attendance_api.py`

#### 1. Added Raw Punch Data Query (Lines 430-470)
- Fetches raw punch data for up to 7 punches per employee per day
- Returns: Employee ID, Employee Name, Date, and all punch times

#### 2. Created `generate_data_sheet()` Function (Lines 492-605)
- Generates the "Data" sheet with raw punch information
- **Columns**: Date, Employee ID, Employee Name, In, Out, In, Out, In, Out
- **Features**:
  - Company name as title
  - "Raw Punch Data" subtitle
  - Date range display
  - Yellow highlighting for Sundays
  - Professional Tahoma font styling
  - Bordered cells
  - Frozen header row
  - Optimized column widths

#### 3. Modified `generate_excel_report()` Function (Lines 607-621)
- Now accepts `raw_punch_df` parameter
- Creates two sheets in order:
  1. **"Data"** sheet (first) - Raw punch data
  2. **"Attendance"** sheet (second) - Processed attendance report

## Excel Output Structure

Your Excel file will now contain **2 sheets**:

### Sheet 1: "Data"
```
| Date       | Employee ID | Employee Name    | In    | Out   | In    | Out   | In    | Out   |
|------------|-------------|------------------|-------|-------|-------|-------|-------|-------|
| 2025-01-15 | 101         | John Doe         | 08:00 | 12:00 | 13:00 | 17:00 |       |       |
| 2025-01-16 | 101         | John Doe         | 08:15 | 12:05 | 13:00 | 17:10 |       |       |
| 2025-01-17 | 101         | John Doe         | 07:55 | 11:58 | 13:02 | 17:05 |       |       |
```

### Sheet 2: "Attendance"
(Your existing processed attendance report with all calculations)

## How to Use

### Starting the API Server

**Option 1: Using run_api.py (Recommended)**
```bash
python run_api.py
```

**Option 2: Direct uvicorn command**
```bash
uvicorn attendance_api:app --host 0.0.0.0 --port 8000 --reload
```

**Option 3: Running main.py (NOT recommended - uses different file)**
```bash
# DON'T USE THIS - main.py is a different implementation
python main.py
```

### Testing the API

**1. Access the Web Interface:**
- Open: http://localhost:8000
- Upload your ZK.db file
- Select date range
- Download the report

**2. Use the API Documentation:**
- Open: http://localhost:8000/docs
- Try the `/generate-attendance-report` endpoint

**3. Use cURL:**
```bash
curl -X POST "http://localhost:8000/generate-attendance-report" \
  -F "db_file=@ZKTimeNet.db" \
  -F "start_date=2025-01-01" \
  -F "end_date=2025-01-31" \
  -F "public_holidays=2025-01-15,2025-01-25" \
  --output attendance_report.xlsx
```

## Which File is Actually Used?

✅ **`attendance_api.py`** - This is the file being used by the API
- Contains the complete SQL logic
- Handles punch time adjustments
- Generates both Data and Attendance sheets

❌ **`main.py`** - NOT used by the API
- Different implementation
- Only used if you run `python main.py` directly

✅ **`run_api.py`** - Entry point script
- Runs `attendance_api.py` when executed
- Line 14: `"attendance_api:app"`

## Feature Highlights

### Data Sheet Features:
- ✅ Shows raw punch times as they appear in the database
- ✅ Employee-friendly format (Date, ID, Name, punch times)
- ✅ Sunday rows highlighted in yellow
- ✅ Up to 6 punch times per day (3 In/Out pairs)
- ✅ Professional formatting matching Attendance sheet
- ✅ Frozen headers for easy scrolling

### Attendance Sheet Features (Unchanged):
- ✅ All existing calculations (Late, Early, OT, etc.)
- ✅ Public holiday support with OT redistribution
- ✅ Suspicious punch detection
- ✅ Employee grouping with totals
- ✅ All styling and conditional formatting

## Verification

To verify the implementation works:

1. Start the API: `python run_api.py`
2. Open: http://localhost:8000/docs
3. Upload a ZK.db file with date range
4. Download the generated Excel file
5. Check that it has 2 sheets: "Data" and "Attendance"

## Notes

- The "Data" sheet always appears first (leftmost tab)
- The "Attendance" sheet appears second
- Both sheets use the same company name and date range
- Raw punch data shows actual database values (not processed)
- Sunday highlighting is consistent across both sheets

