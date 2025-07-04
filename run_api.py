"""
Simple script to run the FastAPI Attendance Report Generator
"""
import uvicorn

if __name__ == "__main__":
    print("🚀 Starting Attendance Report Generator API...")
    print("📋 API Documentation will be available at: http://localhost:8000/docs")
    print("🔗 API Root endpoint: http://localhost:8000/")
    print("📊 Upload endpoint: http://localhost:8000/generate-attendance-report")
    print("\nPress Ctrl+C to stop the server")
    
    uvicorn.run(
        "attendance_api:app", 
        host="0.0.0.0", 
        port=8000, 
        reload=True
    ) 