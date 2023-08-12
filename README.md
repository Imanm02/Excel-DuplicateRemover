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








