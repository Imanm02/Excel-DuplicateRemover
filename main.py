import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

def apply_styles_to_excel(filename):
    # Load the Excel file back into memory
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    # Define the styles to apply to the cells
    font = Font(name='Vazirmatn')
    alignment = Alignment(horizontal='center', vertical='center')
    light_yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    light_green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    # Apply the styles to the cells
    for row in sheet:
        for cell in row:
            cell.font = font
            cell.alignment = alignment
            cell.number_format = '@'  # Set number format to text
            # Apply different background colors to the first row and the rest
            if cell.row == 1:
                cell.fill = light_yellow_fill
            else:
                cell.fill = light_green_fill

    # Adjust the width of the columns
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length

    # Save the changes made to the Excel file
    wb.save(filename)

# Load all sheets of the Excel file into a list of dataframes
all_sheets = pd.read_excel("input.xlsx", sheet_name=None, dtype=str)

# Concatenate all the sheets into one dataframe
full_df = pd.concat(all_sheets.values(), ignore_index=True)

# Drop duplicate rows based on 'Name' and 'Surname' columns
unique_df = full_df.drop_duplicates(subset=['Name', 'Surname'])

# Write the dataframe to an Excel file
output_filename = "output.xlsx"
unique_df.to_excel(output_filename, index=False)

# Apply the styles and formatting to the Excel file
apply_styles_to_excel(output_filename)