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
- ✅ **10-Minute Rule**: Filters duplicate punches (if first In/Out < 10 min apart, skips duplicate)
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

## 🕐 10-Minute Rule

Automatically filters duplicate punches:
- If first **In** and **Out** are less than 10 minutes apart → Skip that Out (duplicate!)
- Shifts remaining punches forward
- Example: `In: 08:30, Out: 08:32` (2 min) → Filters 08:32 and uses next punch

**Why 10 minutes?**
- Real work sessions are rarely < 10 minutes
- Catches accidental double-punches (usually 1-5 min)
- Safe and lenient threshold

To change threshold, edit line 99 in `data_generator_api.py`:
```sql
AND (strftime('%s', full_punch_2) - strftime('%s', full_punch_1)) < 600
                                                                    ^^^
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

**Version**: 1.1.0 | **Port**: 8001 | **Created**: October 2025
