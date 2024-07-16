from docx import Document
from datetime import datetime
import holidays
from calendar import monthrange

# Specify the year and month for the timesheet
year = 2024  # Example year
month = 7    # Example month

# Create a holiday list for Singapore
sg_holidays = holidays.SG(years=year)

# Get the number of days in the specified month
num_days = monthrange(year, month)[1]

path = '/Users/zhaoyan.x/Documents/timesheets/'
# Load your timesheet Word document
doc_filename = 'timesheet.docx'  # Change this to the path of your timesheet
doc = Document(path + doc_filename)

def is_public_holiday(date):
    return date in sg_holidays

# Process each table that is for time logging
for table_index, table in enumerate(doc.tables):
    if table_index not in [1]:  # Assuming the 2nd table is used for time logging
        continue
    for i, row in enumerate(table.rows):
        if i >= 2 and i < 30:  # Skip header row and handle only valid dates
            date_obj_1 = datetime(year, month, i-1)  # Correct the day number
            weekday_1 = date_obj_1.weekday()
            
            date_cell = i
            start_cell_index_1 = 1
            end_cell_index_1 = 2
            start_cell_index_2 = 11
            end_cell_index_2 = 12

            # Avoid calculating a second date if the day exceeds the number of days in the month
            if i <= num_days:
                date_obj_2 = datetime(year, month, i + 14) if i + 14 <= num_days else None
                weekday_2 = date_obj_2.weekday() if date_obj_2 else None
            
            # Set times for the first set of dates
            if i <= 16:
                if weekday_1 == 5:  # Saturday
                    table.rows[date_cell].cells[start_cell_index_1].text = 'Sat'
                    table.rows[date_cell].cells[end_cell_index_1].text = ''
                elif weekday_1 == 6:  # Sunday
                    table.rows[date_cell].cells[start_cell_index_1].text = 'Sun'
                    table.rows[date_cell].cells[end_cell_index_1].text = ''
                else:
                    table.rows[date_cell].cells[start_cell_index_1].text = '0830' if not is_public_holiday(date_obj_1) else ''
                    table.rows[date_cell].cells[end_cell_index_1].text = '1800' if weekday_1 < 4 and not is_public_holiday(date_obj_1) else '1730' if weekday_1 == 4 else ''
            
            # Set times for the second set of dates if valid
            if date_obj_2:
                if weekday_2 == 5:  # Saturday
                    table.rows[date_cell].cells[start_cell_index_2].text = 'Sat'
                    table.rows[date_cell].cells[end_cell_index_2].text = ''
                elif weekday_2 == 6:  # Sunday
                    table.rows[date_cell].cells[start_cell_index_2].text = 'Sun'
                    table.rows[date_cell].cells[end_cell_index_2].text = ''
                else:
                    table.rows[date_cell].cells[start_cell_index_2].text = '0830' if not is_public_holiday(date_obj_2) else ''
                    table.rows[date_cell].cells[end_cell_index_2].text = '1800' if weekday_2 < 4 and not is_public_holiday(date_obj_2) else '1730' if weekday_2 == 4 else ''

# Save the modified document
doc.save(path + 'modified_' + doc_filename)  # Change this to where you want to save the modified file
