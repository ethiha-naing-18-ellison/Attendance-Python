# 🚀 Quick Start Guide

## Get Started in 3 Steps!

### Step 1: Install Dependencies
```bash
cd "Attendance Exel Generator"
pip install -r requirements.txt
```

### Step 2: Run the Server
```bash
python run_data_generator.py
```

You'll see:
```
🚀 Starting Attendance Data Sheet Generator API...
📋 API Documentation: http://localhost:8001/docs
🔗 API Root: http://localhost:8001/
📊 Generate endpoint: http://localhost:8001/generate-data-sheet

✨ This API generates ONLY the Data sheet with raw punch times
   Columns: Employee ID | Name | In | Out | In | Out | In | Out

Press Ctrl+C to stop the server
```

### Step 3: Generate Your Report

**Option A - Web Interface (Easiest!)**
1. Open: http://localhost:8001
2. Upload your ZK.db file
3. Pick dates
4. Click "Generate Data Sheet"
5. Done! 🎉

**Option B - API Docs**
1. Open: http://localhost:8001/docs
2. Click "POST /generate-data-sheet"
3. Click "Try it out"
4. Upload & generate

## 📋 What You Get

An Excel file with ONE sheet called "Data":

```
┌─────────────┬─────────────┬───────┬───────┬───────┬───────┬───────┬───────┐
│ Employee ID │ Name        │  In   │  Out  │  In   │  Out  │  In   │  Out  │
├─────────────┼─────────────┼───────┼───────┼───────┼───────┼───────┼───────┤
│    101      │ John Doe    │ 08:00 │ 12:00 │ 13:00 │ 17:00 │       │       │
│    102      │ Jane Smith  │ 08:15 │ 12:05 │ 13:10 │ 17:15 │       │       │
│    103      │ Bob Johnson │ 09:00 │ 12:30 │ 13:30 │ 18:00 │       │       │
└─────────────┴─────────────┴───────┴───────┴───────┴───────┴───────┴───────┘
```

**Features:**
- ✅ Company name at top
- ✅ Date range shown
- ✅ Raw punch times (no calculations)
- ✅ Sunday rows in yellow
- ✅ Professional formatting
- ✅ Clean and simple

## ⚡ Pro Tips

1. **Single Day**: Just enter start date
2. **Month**: Enter first and last day
3. **Custom**: Any date range works!
4. **Port Conflict?**: Change 8001 to 8002 in `data_generator_api.py`

## 🎯 Difference from Main Project

| Feature | Main Project | This Project |
|---------|-------------|--------------|
| Port | 8000 | 8001 |
| Output | 2 sheets | 1 sheet |
| Data | Calculations | Raw only |

**Both can run at the same time!** 🎉

## 📞 Need Help?

- API Docs: http://localhost:8001/docs
- Full README: See README.md
- Test API: http://localhost:8001/api

---

That's it! Simple as 1-2-3! 🚀

