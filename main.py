import pandas as pd
from datetime import datetime, time
from openpyxl import Workbook

# Load the Excel file containing the attendance data
file_path = "OriginalDatabase.xlsx"
data = pd.read_excel(file_path)

# Convert 'Punch Date' to a datetime format and extract the date part
# Convert 'Punch Time' to a time format
data['Punch Date'] = pd.to_datetime(data['Punch Date']).dt.date
data['Punch Time'] = pd.to_datetime(data['Punch Time'], format='%H:%M:%S').dt.time

# Initializing a dictionary to store the report data for each employee
report = {}

# Processing attendance records for each unique employee
for employee in data['Employee ID'].unique():
    # Filtering records for the current employee
    emp_data = data[data['Employee ID'] == employee]
    
    # Initializing counters for the current employee
    total_present = 0
    half_day = 0
    late_in = 0
    early_out = 0

    # Processing records for each unique date for the current employee
    for date in emp_data['Punch Date'].unique():
        # Filtering records for the current date
        day_records = emp_data[emp_data['Punch Date'] == date]
        
        # Extracting punch-in and punch-out times
        punch_in_records = day_records[day_records['Directionality'] == 'In']['Punch Time']
        punch_out_records = day_records[day_records['Directionality'] == 'Out']['Punch Time']
        
        # Ensuring that there are both punch-in and punch-out records
        if not punch_in_records.empty and not punch_out_records.empty:
            punch_in_time = punch_in_records.iloc[0]  # Get the first punch-in time
            punch_out_time = punch_out_records.iloc[0]  # Get the first punch-out time

            # Combining punch-in and punch-out times with the current date to calculate working hours
            punch_in_datetime = datetime.combine(datetime.today(), punch_in_time)
            punch_out_datetime = datetime.combine(datetime.today(), punch_out_time)
            working_hours = (punch_out_datetime - punch_in_datetime).total_seconds() / 3600.0

            # Applying attendance rules
            if working_hours < 4:  # Less than 4 hours worked is considered a half day
                half_day += 1
            else:
                # Check if the punch-in time is after 11:30 AM or 11:00 AM (Late In)
                if punch_in_time > time(11, 30):
                    half_day += 1
                elif punch_in_time > time(11, 0):
                    late_in += 1
                
                # Check if the punch-out time is before 5:00 PM or 6:00 PM (Early Out)
                if punch_out_time < time(17, 0):
                    half_day += 1
                elif punch_out_time < time(18, 0):
                    early_out += 1

                # If all conditions for a full day are met, increment the total present counter
                total_present += 1

    # Storing the calculated statistics for the current employee in the report dictionary
    report[employee] = {
        'Total Present': total_present,
        'Half Days': half_day,
        'Late In': late_in,
        'Early Out': early_out
    }

# Defining the output file path for the attendance report
output_file_path = "Report.xlsx"
wb = Workbook()  # Create a new Excel workbook
ws = wb.active  # Get the active worksheet
ws.title = "Report"  # Set the title of the worksheet

# Defining the headers for the report and append them to the worksheet
headers = ['Employee ID', 'Total Present', 'Half Days', 'Late In', 'Early Out']
ws.append(headers)

# Appending the report data for each employee to the worksheet
for employee, data in report.items():
    row = [employee, data['Total Present'], data['Half Days'], data['Late In'], data['Early Out']]
    ws.append(row)

# Saving the workbook
wb.save(output_file_path)

print("Attendance report has been successfully generated and saved to Report.xlsx")
