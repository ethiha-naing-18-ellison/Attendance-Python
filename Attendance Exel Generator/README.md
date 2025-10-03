# 📊 Attendance Data Sheet Generator

Simple FastAPI app that generates Excel files with **raw punch data** from ZK.db database files.

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Server
```bash
python run_data_generator.py
```

### 3. Open Browser
```
http://localhost:8001
```

Upload your ZK.db file, select dates, and generate!

## 📋 What You Get

Excel file with one sheet called "Data":

| Employee ID | Name        | In    | Out   | In    | Out   | In    | Out   |
|-------------|-------------|-------|-------|-------|-------|-------|-------|
| 101         | John Doe    | 08:00 | 12:00 | 13:00 | 17:00 |       |       |
| 102         | Jane Smith  | 08:15 | 12:05 | 13:10 | 17:15 |       |       |
| 103         | Bob Johnson | 09:00 | 12:30 | 13:30 | 18:00 |       |       |

**Features:**
- ✅ Raw punch times (no calculations)
- ✅ **Smart 10-Minute Rule**: First half keeps earlier time, last half keeps later time
- ✅ Sunday rows highlighted in yellow
- ✅ Professional formatting

## 🔧 API Endpoints

### POST `/generate-data-sheet`
Upload ZK.db file and generate Excel report.

**Parameters:**
- `db_file` (File): Your ZK.db database file
- `start_date` (String): Start date (YYYY-MM-DD, e.g., 2025-01-01)
- `end_date` (String, optional): End date (defaults to start_date)

**Example:**
```bash
curl -X POST "http://localhost:8001/generate-data-sheet" \
  -F "db_file=@ZKTimeNet.db" \
  -F "start_date=2025-01-01" \
  -F "end_date=2025-01-31" \
  --output attendance_data.xlsx
```

### GET `/`
Web interface for file upload

### GET `/docs`
Interactive API documentation (Swagger UI)

## 🕐 10-Minute Rule (Duplicate Filter)

**Checks ALL consecutive punches!** Uses smart logic based on position:

### Logic:

**FIRST HALF (Punch 1-4):** Keep **FIRST** punch (earlier time)
- ✅ Punch 1 → Punch 2: If < 10 min, keep Punch 1, remove Punch 2
- ✅ Punch 2 → Punch 3: If < 10 min, keep Punch 2, remove Punch 3  
- ✅ Punch 3 → Punch 4: If < 10 min, keep Punch 3, remove Punch 4

**SECOND HALF (Punch 5-6):** Keep **SECOND** punch (later time)
- ✅ Punch 4 → Punch 5: If < 10 min, keep Punch 5, remove Punch 4
- ✅ Punch 5 → Punch 6: If < 10 min, keep Punch 6, remove Punch 5

**Why different logic?**
- First punches = Morning clock-in/out → Keep earlier time (actual start)
- Last punches = Evening in/out → Keep later time (actual end)

### Examples:

**Example 1: Duplicate at beginning (Keep FIRST)**
```
Input:  08:30, 08:32, 12:00, 14:00
        ^^^^^^^^^^^^^ Only 2 minutes!
Output: 08:30, 12:00, 14:00 ✅ (kept 08:30, removed 08:32)
```

**Example 2: Duplicate in middle (Keep FIRST)**
```
Input:  08:30, 12:00, 14:10, 14:12, 18:00
                      ^^^^^^^^^^^^^ Only 2 minutes!
Output: 08:30, 12:00, 14:10, 18:00 ✅ (kept 14:10, removed 14:12)
```

**Example 3: Duplicate at end (Keep SECOND)**
```
Input:  08:30, 12:00, 14:10, 20:06, 20:07
                            ^^^^^^^^^^^^^ Only 1 minute!
Output: 08:30, 12:00, 14:10, 20:07 ✅ (kept 20:07, removed 20:06)
```

**Example 4: Multiple duplicates**
```
Input:  08:30, 08:32, 12:00, 20:06, 20:07
        ^^^^^^^^^^^^^ first   ^^^^^^^^^^^^^ second
Output: 08:30, 12:00, 20:07 ✅
        (kept 08:30 first, kept 20:07 second)
```

**Why 10 minutes?**
- Real work sessions are rarely < 10 minutes
- Catches accidental double-punches (usually 1-5 min)
- Safe and lenient threshold
- Works for duplicates ANYWHERE in the sequence

**To change threshold**, edit TWO places in `data_generator_api.py`:

1. **Line 169** (SQL - first check): Change `< 600`
2. **Line 75** (Python - all other checks): Change `< 600`

```
600 = 10 minutes
300 = 5 minutes
900 = 15 minutes
```

## 📦 Project Files

```
Attendance Exel Generator/
├── data_generator_api.py    # Main FastAPI application
├── run_data_generator.py    # Server startup script
├── requirements.txt         # Python dependencies
├── upload_form.html         # Web interface
└── README.md               # This file
```

## 💡 Tips

- **Single Day**: Only provide `start_date`
- **Month**: Provide `start_date` and `end_date`
- **Port Conflict**: Change port 8001 in `data_generator_api.py` (last line)
- **View API Docs**: http://localhost:8001/docs

## 🆚 vs Main Attendance API

| Feature | Main API (Port 8000) | This API (Port 8001) |
|---------|---------------------|---------------------|
| Sheets | 2 (Data + Attendance) | 1 (Data only) |
| Calculations | Yes (OT, Late, etc.) | No |
| Purpose | Complete reports | Raw data export |
| Complexity | High | Low |

**Both can run at the same time!**

---

**Version**: 1.4.0 | **Port**: 8001 | **Updated**: October 2025
