import pandas as pd
import json
import re
import os

def is_valid_date(date_str):
    """
    Check if a string represents a valid date in MM/DD/YYYY format
    """
    if pd.isna(date_str) or str(date_str).strip() == '':
        return False
    
    try:
        date_pattern = r'\d{1,2}/\d{1,2}/\d{4}'
        return bool(re.match(date_pattern, str(date_str).strip()))
    except:
        return False

def clean_value(value):
    """
    Clean and format values for JSON output
    """
    if pd.isna(value) or str(value).strip() == '' or str(value).lower() == 'nan':
        return None
    return str(value).strip()

def is_total_row(row):
    """
    Check if this is a total/summary row that ends the attendance data
    """
    first_col = str(row.iloc[0]).strip().lower()
    return 'total' in first_col or 'checked by' in first_col

def extract_employee_info(df, header_row_idx):
    """
    Extract employee information from the row immediately below the "Employee ID" header
    """
    try:
        # Get the row immediately below the header row
        if header_row_idx + 1 >= len(df):
            return None
            
        info_row = df.iloc[header_row_idx + 1]
        
        # Extract employee data from specific columns as specified
        employee_id = clean_value(info_row.iloc[1])    # Column B (index 1)
        full_name = clean_value(info_row.iloc[3])      # Column D (index 3)
        department = clean_value(info_row.iloc[5])     # Column F (index 5)
        
        # Try to convert employee_id to integer if it's numeric
        try:
            if employee_id and employee_id.isdigit():
                employee_id = int(employee_id)
        except:
            pass
        
        return {
            'employeeid': employee_id,
            'fullname': full_name,
            'Department': department
        }
        
    except Exception as e:
        print(f"Error extracting employee info at row {header_row_idx}: {e}")
        return None

def extract_attendance_data(df, start_row, end_row):
    """
    Extract attendance data for one employee from start_row to end_row
    """
    attendance_records = []
    
    print(f"   Extracting attendance data from rows {start_row} to {end_row}")
    
    for i in range(start_row, min(end_row, len(df))):
        try:
            row = df.iloc[i]
            
            # Check if this is a total row - if so, stop processing
            if is_total_row(row):
                print(f"   Found total row at {i}, stopping attendance extraction")
                break
            
            # Check if this row has a valid date in column A
            date_val = row.iloc[0]
            if not is_valid_date(date_val):
                continue
            
            # Extract attendance data from specified columns
            attendance_record = {
                'Date': clean_value(row.iloc[0]),           # Column A
                'Workday': clean_value(row.iloc[1]),        # Column B
                'Timetable': clean_value(row.iloc[2]),      # Column C
                'Col_3': None,                              # Always null as specified
                'Clock-In': clean_value(row.iloc[4]) if len(row) > 4 else None,  # Column E
                'Col_5': None,                              # Always null as specified
                'Clock-Out': clean_value(row.iloc[6]) if len(row) > 6 else None, # Column G
                'Col_7': None,                              # Always null as specified
                'In': clean_value(row.iloc[8]) if len(row) > 8 else None,        # Column I
                'Out': clean_value(row.iloc[9]) if len(row) > 9 else None        # Column J
            }
            
            attendance_records.append(attendance_record)
            
        except Exception as e:
            print(f"   Error processing attendance row {i}: {e}")
            continue
    
    return attendance_records

def find_employee_blocks(df):
    """
    Find all employee block starting positions by looking for "Employee ID" in column A
    """
    employee_blocks = []
    
    for i, row in df.iterrows():
        first_col = str(row.iloc[0]).strip()
        
        # Look for "Employee ID" in the first column
        if 'Employee ID' in first_col:
            employee_blocks.append(i)
    
    return employee_blocks

def find_attendance_start(df, employee_header_row):
    """
    Find where attendance data starts for this employee (after employee info row)
    """
    # Start looking from 2 rows after the employee header
    start_search = employee_header_row + 2
    
    # Look for the first row that contains "Date" in column A (attendance header)
    for i in range(start_search, min(start_search + 10, len(df))):
        if i < len(df):
            first_col = str(df.iloc[i].iloc[0]).strip().lower()
            if 'date' in first_col:
                return i + 1  # Return the row after the header
    
    # If no "Date" header found, assume attendance starts 3 rows after employee header
    return employee_header_row + 3

def convert_multi_employee_excel_to_json(file_path):
    """
    Convert multi-employee Excel attendance file to structured JSON
    """
    print(f"📖 Reading Excel file: {file_path}")
    
    # Read the Excel file without assuming headers
    df = pd.read_excel(file_path, header=None)
    
    print(f"📊 Total rows in file: {len(df)}")
    
    # Find all employee blocks
    employee_blocks = find_employee_blocks(df)
    print(f"🔍 Found {len(employee_blocks)} employee blocks at rows: {employee_blocks}")
    
    employees_data = []
    
    # Process each employee block
    for block_idx, employee_header_row in enumerate(employee_blocks):
        try:
            print(f"\n👤 Processing employee block {block_idx + 1} (header at row {employee_header_row})...")
            
            # Extract employee information
            employee_info = extract_employee_info(df, employee_header_row)
            
            if employee_info is None:
                print(f"   ❌ Failed to extract employee info")
                continue
            
            print(f"   ✅ Employee: {employee_info['fullname']} (ID: {employee_info['employeeid']}, Dept: {employee_info['Department']})")
            
            # Find where attendance data starts
            attendance_start = find_attendance_start(df, employee_header_row)
            
            # Determine where this employee's data ends
            if block_idx < len(employee_blocks) - 1:
                # Next employee block starts
                attendance_end = employee_blocks[block_idx + 1]
            else:
                # Last employee, go to end of file
                attendance_end = len(df)
            
            # Extract attendance data
            attendance_data = extract_attendance_data(df, attendance_start, attendance_end)
            
            print(f"   📅 Extracted {len(attendance_data)} attendance records")
            
            # Create employee record
            employee_record = {
                'employeeid': employee_info['employeeid'],
                'fullname': employee_info['fullname'],
                'Department': employee_info['Department'],
                'Attendance': attendance_data
            }
            
            employees_data.append(employee_record)
            
        except Exception as e:
            print(f"   ❌ Error processing employee block {block_idx + 1}: {e}")
            continue
    
    return employees_data

def save_employees_to_json(employees_data, output_file):
    """
    Save employees data to JSON file with specified formatting
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(employees_data, f, indent=4, ensure_ascii=False)
        
        print(f"💾 Successfully saved {len(employees_data)} employees to: {output_file}")
        return True
        
    except Exception as e:
        print(f"❌ Error saving JSON file: {e}")
        return False

def display_summary(employees_data):
    """
    Display a comprehensive summary of the processed data
    """
    print(f"\n📈 PROCESSING SUMMARY:")
    print("=" * 70)
    print(f"Total employees processed: {len(employees_data)}")
    
    total_attendance_records = sum(len(emp['Attendance']) for emp in employees_data)
    print(f"Total attendance records: {total_attendance_records}")
    
    print(f"\n👥 EMPLOYEE DETAILS:")
    print("-" * 70)
    
    for i, employee in enumerate(employees_data, 1):
        attendance_count = len(employee['Attendance'])
        print(f"{i:2d}. ID: {employee['employeeid']} | {employee['fullname']} | {employee['Department']} | {attendance_count} records")
        
        # Show sample attendance record if available
        if attendance_count > 0:
            sample = employee['Attendance'][0]
            date_range = f"{employee['Attendance'][0]['Date']}"
            if attendance_count > 1:
                date_range += f" to {employee['Attendance'][-1]['Date']}"
            print(f"     Period: {date_range}")
            
            # Show a sample with actual times if available
            if sample.get('Clock-In') or sample.get('In'):
                times = []
                if sample.get('Clock-In'):
                    times.append(f"Clock-In: {sample['Clock-In']}")
                if sample.get('In'):
                    times.append(f"In: {sample['In']}")
                if times:
                    print(f"     Sample times: {' | '.join(times)}")

def main():
    """
    Main function to convert multi-employee attendance Excel to JSON
    """
    # Input and output files
    excel_file = "attendance_report_6551.xlsx"
    output_file = "employees_attendance.json"
    
    print("🚀 MULTI-EMPLOYEE ATTENDANCE CONVERTER")
    print("=" * 70)
    
    # Check if input file exists
    if not os.path.exists(excel_file):
        print(f"❌ Error: Input file '{excel_file}' not found!")
        return
    
    try:
        # Convert the Excel file
        employees_data = convert_multi_employee_excel_to_json(excel_file)
        
        if not employees_data:
            print("❌ No employee data found or processed!")
            return
        
        # Display summary
        display_summary(employees_data)
        
        # Save to JSON
        if save_employees_to_json(employees_data, output_file):
            print(f"\n🎉 Conversion completed successfully!")
            print(f"📁 Output file: {output_file}")
            
            # Show sample JSON structure
            print(f"\n📋 SAMPLE JSON STRUCTURE:")
            print("-" * 50)
            if employees_data:
                sample_employee = employees_data[0].copy()
                # Show only first 2 attendance records for preview
                if len(sample_employee['Attendance']) > 2:
                    original_count = len(sample_employee['Attendance'])
                    sample_employee['Attendance'] = sample_employee['Attendance'][:2]
                    sample_employee['Attendance'].append({
                        "...": f"and {original_count - 2} more attendance records"
                    })
                
                print(json.dumps(sample_employee, indent=2, ensure_ascii=False))
        
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 