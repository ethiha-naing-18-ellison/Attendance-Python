"""
Simple script to run the Attendance Data Sheet Generator API
"""
import uvicorn

if __name__ == "__main__":
    print("🚀 Starting Attendance Data Sheet Generator API...")
    print("📋 API Documentation: http://localhost:8001/docs")
    print("🔗 API Root: http://localhost:8001/")
    print("📊 Generate endpoint: http://localhost:8001/generate-data-sheet")
    print("\n✨ This API generates ONLY the Data sheet with raw punch times")
    print("   Columns: Employee ID | Name | In | Out | In | Out | In | Out")
    print("\nPress Ctrl+C to stop the server\n")
    
    uvicorn.run(
        "data_generator_api:app", 
        host="0.0.0.0", 
        port=8001, 
        reload=True
    )

