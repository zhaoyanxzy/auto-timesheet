from docx import Document
from datetime import datetime
import holidays
from calendar import monthrange

# Auto update the month and year
now = datetime.now()
year = now.year
month = now.month

# Create a holiday list for Singapore
sg_holidays = holidays.SG(years=year)

# Get the number of days in the current month
num_days = monthrange(year, month)[1]

# Path to your timesheet document
path = "/Users/zhaoyan.x/Documents/Timesheets/"
doc_filename = "timesheet.docx"  # Adjust if needed
doc = Document(path + doc_filename)


def is_public_holiday(date):
    return date in sg_holidays


def set_cell_text(cell, text):
    cell.text = text
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_after = 0  # Remove extra space after text


# Process each table that is for time logging
for table_index, table in enumerate(doc.tables):
    if table_index not in [1]:  # Assuming the 2nd table is for time logging
        continue
    rows = table.rows
    for i, row in enumerate(rows):
        print(f"Processing row {i}")
        date = row.cells[0].text
        if not date:
            continue
        # Assuming that the first date starts in row 2
        if 1 <= i <= num_days:  # Ensure we handle valid dates only
            row_index_offset = i + 1
            day_num_1 = i  # Dates start from 1 in the table
            print(f"Day 1: {day_num_1}")
            start_cell_index_1 = 1
            end_cell_index_1 = 3

            if day_num_1 <= num_days:
                date_obj = datetime(year, month, day_num_1)
                weekday = date_obj.weekday()

            if day_num_1 < 16:
                # Mark public holidays
                if is_public_holiday(date_obj):
                    set_cell_text(
                        rows[row_index_offset].cells[start_cell_index_1], "PH"
                    )
                    set_cell_text(rows[row_index_offset].cells[end_cell_index_1], "PH")
                else:
                    # Set regular work hours based on weekday
                    if weekday == 5:  # Saturday
                        set_cell_text(
                            rows[row_index_offset].cells[start_cell_index_1], "Sat"
                        )
                        set_cell_text(
                            rows[row_index_offset].cells[end_cell_index_1], "Sat"
                        )
                    elif weekday == 6:  # Sunday
                        set_cell_text(
                            rows[row_index_offset].cells[start_cell_index_1], "Sun"
                        )
                        set_cell_text(
                            rows[row_index_offset].cells[end_cell_index_1], "Sun"
                        )
                    else:
                        # Weekday work hours
                        set_cell_text(
                            rows[row_index_offset].cells[start_cell_index_1], "0830"
                        )
                        set_cell_text(
                            rows[row_index_offset].cells[end_cell_index_1],
                            "1800" if weekday < 4 else "1730",
                        )

            day_num_2 = i + 15
            print(f"Day 2: {day_num_2}")

            start_cell_index_2 = 11
            end_cell_index_2 = 13

            if day_num_2 <= num_days:
                date_obj = datetime(year, month, day_num_2)
                weekday = date_obj.weekday()

            if day_num_2 < 32:
                # Mark public holidays
                if is_public_holiday(date_obj):
                    set_cell_text(
                        rows[row_index_offset].cells[start_cell_index_2], "PH"
                    )
                    set_cell_text(rows[row_index_offset].cells[end_cell_index_2], "PH")
                else:
                    # Set regular work hours based on weekday
                    if weekday == 5:  # Saturday
                        set_cell_text(
                            rows[row_index_offset].cells[start_cell_index_2], "Sat"
                        )
                        set_cell_text(
                            rows[row_index_offset].cells[end_cell_index_2], "Sat"
                        )
                    elif weekday == 6:  # Sunday
                        set_cell_text(
                            rows[row_index_offset].cells[start_cell_index_2], "Sun"
                        )
                        set_cell_text(
                            rows[row_index_offset].cells[end_cell_index_2], "Sun"
                        )
                    else:
                        # Weekday work hours
                        set_cell_text(
                            rows[row_index_offset].cells[start_cell_index_2], "0830"
                        )
                        set_cell_text(
                            rows[row_index_offset].cells[end_cell_index_2],
                            "1800" if weekday < 4 else "1730",
                        )


# Save the modified document
doc.save(path + "modified_" + doc_filename)
