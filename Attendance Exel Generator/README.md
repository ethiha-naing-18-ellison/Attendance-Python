# 📊 Attendance Data Sheet Generator

A standalone FastAPI application that generates Excel files with **raw punch data only** from ZK.db database files.

## ✨ Features

- 📄 **Single Sheet Output**: Generates only the "Data" sheet
- 🎯 **Simple Format**: Employee ID, Name, and up to 6 punch times (3 In/Out pairs)
- 🕐 **5-Minute Rule**: Automatically filters duplicate punches (removes false punches < 5 min apart)
- 🌟 **Clean Design**: No calculations, just raw punch data
- 🎨 **Professional Styling**: 
  - Tahoma font
  - Bordered cells
  - Header highlighting
  - Sunday rows in yellow
  - Frozen header row
- 🚀 **Easy to Use**: Web interface + API endpoints
- ⚡ **Fast Processing**: Lightweight and efficient

## 📋 Output Format

### Excel Sheet: "Data"

| Employee ID | Name        | In    | Out   | In    | Out   | In    | Out   |
|-------------|-------------|-------|-------|-------|-------|-------|-------|
| 101         | John Doe    | 08:00 | 12:00 | 13:00 | 17:00 |       |       |
| 102         | Jane Smith  | 08:15 | 12:05 | 13:10 | 17:15 |       |       |
| 103         | Bob Johnson | 09:00 | 12:30 | 13:30 | 18:00 |       |       |

**Note**: Sunday rows are highlighted in yellow

## 🚀 Installation

### 1. Install Dependencies

```bash
cd "Attendance Exel Generator"
pip install -r requirements.txt
```

### 2. Run the API Server

```bash
python run_data_generator.py
```

The server will start on **http://localhost:8001**

## 📖 Usage

### Option 1: Web Interface (Easiest)

1. Open your browser: http://localhost:8001
2. Upload your ZK.db file
3. Select date range (start date required, end date optional)
4. Click "Generate Data Sheet"
5. Download the Excel file

### Option 2: Interactive API Documentation

1. Open: http://localhost:8001/docs
2. Click on **POST /generate-data-sheet**
3. Click "Try it out"
4. Upload file and enter dates
5. Click "Execute"
6. Download the generated file

### Option 3: Command Line (cURL)

```bash
curl -X POST "http://localhost:8001/generate-data-sheet" \
  -F "db_file=@ZKTimeNet.db" \
  -F "start_date=2025-01-01" \
  -F "end_date=2025-01-31" \
  --output attendance_data.xlsx
```

### Option 4: Python Script

```python
import requests

url = "http://localhost:8001/generate-data-sheet"
files = {'db_file': open('ZKTimeNet.db', 'rb')}
data = {
    'start_date': '2025-01-01',
    'end_date': '2025-01-31'  # Optional
}

response = requests.post(url, files=files, data=data)

if response.status_code == 200:
    with open('attendance_data.xlsx', 'wb') as f:
        f.write(response.content)
    print("✅ Data sheet generated successfully!")
else:
    print(f"❌ Error: {response.json()}")
```

## 🔧 API Endpoints

### POST `/generate-data-sheet`

Generate Excel file with raw punch data.

**Parameters:**
- `db_file` (File, required): ZK.db SQLite database file
- `start_date` (string, required): Start date in YYYY-MM-DD format
- `end_date` (string, optional): End date in YYYY-MM-DD format

**Response:**
- Success: Excel file download
- Error: JSON with error details

### GET `/`

Web interface for file upload

### GET `/docs`

Interactive API documentation (Swagger UI)

### GET `/api`

API information endpoint

## 📁 Project Structure

```
Attendance Exel Generator/
├── data_generator_api.py      # Main FastAPI application
├── run_data_generator.py      # Script to start the API server
├── requirements.txt           # Python dependencies
├── upload_form.html           # Web interface
└── README.md                  # This file
```

## 🎯 What Makes This Different?

This is a **standalone project** separate from the main attendance system:

| Feature | Main Attendance API | Data Generator API |
|---------|--------------------|--------------------|
| Port | 8000 | **8001** |
| Sheets | Data + Attendance | **Data only** |
| Calculations | Full calculations | **None** |
| Purpose | Complete reports | **Raw punch data** |
| Complexity | High | **Low** |

## 📊 Column Details

1. **Employee ID**: Employee's ID number
2. **Name**: Full name of employee
3. **In** (1st): First punch-in time
4. **Out** (1st): First punch-out time
5. **In** (2nd): Second punch-in time (after lunch)
6. **Out** (2nd): Second punch-out time
7. **In** (3rd): Third punch-in time (if exists)
8. **Out** (3rd): Third punch-out time (if exists)

## 🎨 Excel Features

- **Company Name**: Displayed at the top in large font
- **Subtitle**: "Raw Punch Data Report"
- **Date Range**: Shown under subtitle
- **Headers**: Light yellow background with bold font
- **Sunday Rows**: Yellow background for easy identification
- **Time Format**: HH:MM (no seconds)
- **Borders**: All data cells have borders
- **Frozen Panes**: Header row stays visible when scrolling
- **Column Widths**: Optimized for readability

## 🔍 Troubleshooting

### Port Already in Use

If port 8001 is already in use, you can change it:

1. Open `data_generator_api.py`
2. Change line: `uvicorn.run(app, host="0.0.0.0", port=8001)`
3. Update port number to something else (e.g., 8002)
4. Restart the server

### No Data Found

- Check that your date range is correct
- Verify that the database has punch data for that period
- Ensure the database file is a valid ZK.db file

### File Upload Fails

- Maximum file size depends on your system
- Ensure the file has `.db` extension
- Check that the file is not corrupted

## 📝 Date Format

- **Format**: YYYY-MM-DD
- **Example**: 2025-01-15
- **start_date**: Required
- **end_date**: Optional (defaults to start_date if not provided)

## 🚦 Status Codes

- **200**: Success - Excel file generated
- **400**: Bad Request - Invalid input or database error
- **404**: Not Found - No punch data found for date range
- **500**: Internal Server Error

## 💡 Tips

1. **Single Day Report**: Only provide start_date
2. **Month Report**: Set start_date to first day, end_date to last day
3. **Custom Range**: Any date range is supported
4. **Large Files**: The API handles large datasets efficiently
5. **Multiple Employees**: All employees in the date range are included

## 🔒 Security Notes

This API is designed for **internal use**. For production:

- Add authentication
- Implement rate limiting
- Add file size limits
- Use HTTPS
- Validate all inputs

## 🕐 5-Minute Rule (Duplicate Filter)

The system automatically detects and removes duplicate punches that occur within 5 minutes of each other.

### Example:

**Before filtering:**
```
In: 08:30, Out: 08:32 (only 2 minutes - duplicate!)
```

**After filtering:**
```
In: 08:30, Out: (shifts to next valid punch)
```

This prevents false duplicate entries when employees accidentally punch multiple times.

📖 **For detailed explanation**, see: `5_MINUTE_RULE_EXPLAINED.md`

## 📦 Dependencies

- **FastAPI**: Web framework
- **Uvicorn**: ASGI server
- **Pandas**: Data processing
- **OpenPyXL**: Excel generation
- **SQLite3**: Database connectivity (built-in)

## 🆚 Comparison with Main Project

### This Project (Data Generator)
- ✅ Simple and focused
- ✅ Only raw punch data
- ✅ No calculations
- ✅ Single sheet output
- ✅ Lightweight

### Main Attendance API
- ✅ Comprehensive reports
- ✅ Full calculations (OT, Late, etc.)
- ✅ Multiple sheets
- ✅ Public holiday handling
- ✅ Advanced features

**Use this project when you need**: Simple raw punch data export  
**Use main project when you need**: Complete attendance reports with calculations

## 📞 Support

For issues or questions, refer to the API documentation at:
- http://localhost:8001/docs (when server is running)

## 🎓 Example Output

When you generate a report, you'll get an Excel file with:
- Company name at top
- Date range displayed
- Clean table with employee data
- All punch times for the period
- Professional formatting
- Easy to read and share

---

**Version**: 1.0.0  
**Port**: 8001  
**License**: Use as needed  
**Created**: October 2025

