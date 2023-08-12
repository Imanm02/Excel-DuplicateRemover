# Excel Duplicate Remover

This repository contains a Python script tailored to efficiently process Excel sheets by removing duplicates based on 'Name' and 'Surname' columns. The tool applies custom formatting to enhance the visual clarity of the resulting Excel file.

## Features

- **Duplicate Removal**: The script identifies and removes duplicate rows from the Excel sheets based on the 'Name' and 'Surname' columns.
- **Custom Formatting**: Applying unique styles, such as font choice and cell color, allows for enhanced readability.
- **Automation**: With this script, the tedious task of manually looking for duplicates and then formatting the file is transformed into a seamless automated process.

## Code Walkthrough

Here's a brief walkthrough of the script's sections:

1. **Importing Libraries**: The initial section imports essential Python libraries. We use `pandas` for handling and processing Excel data, and `openpyxl` for applying specific formatting and styles to the Excel sheets.

```python
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
```

2. **Reading and Processing Excel File**: The script starts by reading the input Excel file, followed by removing duplicates based on 'Name' and 'Surname'.

```python
data = pd.read_excel("Micro.xlsx")
data.drop_duplicates(subset=['Name', 'Surname'], inplace=True, keep="last")
```

3. **Applying Formatting to Excel File**: After processing, the data undergoes formatting where specific styles are applied for enhanced readability.

```python
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
```

4. **Saving Processed Excel File**: Once all processing and formatting steps are complete, the Excel file is saved with the applied changes.

```python
data.to_excel('output.xlsx', index=False)
```

## Usage

To utilize this tool, ensure the input Excel file named `input.xlsx` is placed in the same directory as the script. Run the `main.py` script to process the data and retrieve the resulting Excel file as `output.xlsx` in the root directory.

For successful execution, the required Python libraries `pandas` and `openpyxl` must be installed in your Python environment.

# Maintainer
- [Iman Mohammadi](https://github.com/Imanm02)
