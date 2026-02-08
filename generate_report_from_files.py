"""
Script to generate a report from files in the 2-1 files folder
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# Add current directory to path to import app functions
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import generate_report
import pandas as pd
from openpyxl import Workbook

# Path to the files folder
files_folder = Path("../2-7 files")

# Map files
doxy_file_path = files_folder / "meeting_history_02_07_2026_America_New_York.csv"
account_file_path = files_folder / "AccountDetailReport_Feb.01.2026-Feb.08.2026.csv"
gusto_file_path = files_folder / "duval-medical-p-a-time-tracking-hours-2026-02-01-to-2026-02-07.csv"
booking_file_path = files_folder / "BookingPageSummaryReport_Feb.01.2026-Feb.08.2026.xls"

print("Generating report from files in 2-7 files folder...")
print(f"Doxy Report: {doxy_file_path.name}")
print(f"Account Detail: {account_file_path.name}")
print(f"Gusto Hours: {gusto_file_path.name}")
print(f"Booking Summary: {booking_file_path.name}")
print()

# Open files as file-like objects
try:
    with open(doxy_file_path, 'rb') as doxy_file, \
         open(account_file_path, 'rb') as account_file, \
         open(gusto_file_path, 'rb') as gusto_file:
        
        # Create file-like objects that mimic Flask's request.files
        class FileWrapper:
            def __init__(self, file_path, file_obj):
                self.filename = file_path.name
                self.file_obj = file_obj
            
            def read(self, size=-1):
                if size == -1:
                    return self.file_obj.read()
                return self.file_obj.read(size)
            
            def seek(self, pos):
                return self.file_obj.seek(pos)
            
            def tell(self):
                return self.file_obj.tell()
        
        doxy_wrapped = FileWrapper(doxy_file_path, doxy_file)
        account_wrapped = FileWrapper(account_file_path, account_file)
        gusto_wrapped = FileWrapper(gusto_file_path, gusto_file)
        
        # Handle optional booking file
        booking_wrapped = None
        if booking_file_path.exists():
            with open(booking_file_path, 'rb') as booking_file:
                booking_wrapped = FileWrapper(booking_file_path, booking_file)
        
        # Generate the report
        print("Processing files...")
        output, stats, preview_data = generate_report(
            doxy_wrapped,
            account_wrapped,
            gusto_wrapped,
            booking_wrapped
        )
        
        # Save the report
        generation_date = datetime.now().strftime("%m-%d-%Y")
        report_name = f'Weekly Report ({generation_date})'
        output_path = files_folder / f"{report_name}.xlsx"
        
        with open(output_path, 'wb') as f:
            f.write(output.getvalue())
        
        print(f"\nReport generated successfully!")
        print(f"File saved to: {output_path}")
        print(f"\nReport Statistics:")
        print(f"  - Providers: {stats['providers']}")
        print(f"  - Total Visits: {stats['total_visits']}")
        print(f"  - Sheets: {stats['sheets']}")
        
except FileNotFoundError as e:
    print(f"Error: File not found - {e}")
    sys.exit(1)
except Exception as e:
    print(f"Error generating report: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)


