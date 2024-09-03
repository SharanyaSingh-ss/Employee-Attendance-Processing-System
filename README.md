# Employee Attendance Processing System

This project provides a Python-based solution to process employee attendance data from an Excel sheet.
The script generates a report with details like total present days, number of half-days, late arrivals, and early departures for each employee.

## Features

- **Automated Attendance Calculation**: Processes employee punch-in and punch-out times.
- **Detailed Report**: Generates a comprehensive report summarizing each employee's attendance.
- **Customizable Rules**: Easily adaptable to different attendance rules and conditions.

## Rules Implemented

- **Late In**: Punch-in after 11:00 AM.
- **Half Day (First Half)**: Punch-in after 11:30 AM.
- **Half Day (Second Half)**: Punch-out before 5:00 PM.
- **Early Out**: Punch-out before 6:00 PM.
- **Full Day Leave**: Total working hours less than 4 hours.

## Prerequisites

- **Python 3.x**: Ensure Python is installed on your system.
- **Pandas Library**: Install pandas using pip.
- **OpenPyXL Library**: Used for Excel file handling.
