# Allergen Data Extraction and Visualization

This script extracts allergen data and date values from Excel files, then creates visualizations showing allergen values over time. Supports filtering by city, selecting specific allergens, and customizing the time range.

## Features

- **Batch Processing**: Automatically finds and processes all Excel files matching a city name
- **Flexible Allergen Selection**: Choose any allergen column by name or index
- **Smart Allergen Detection**: Automatically detects available allergens and suggests alternatives if not found
- **Customizable Time Range**: Display data for any number of years (default: 10)
- **Data Extraction**: Extracts date (column A) and allergen values from any column
- **Weekly Aggregation**: Groups data by weeks to reduce data points for better visualization
- **Professional Visualization**: Creates clean scatter plots with line overlay
- **Month-Year Labels**: X-axis displays month-year format (e.g., Jan-26, Feb-26)

## Dependencies

- **pandas**: Data manipulation and Excel file reading
- **matplotlib**: Creating graphs and visualizations
- **openpyxl**: Reading .xlsx Excel files (required by pandas)
- **xlrd**: Reading .xls Excel files (required by pandas)

## Installation

1. Make sure you have Python 3.7+ installed
2. Install required packages:

```bash
pip install pandas openpyxl xlrd matplotlib
```

Or install all at once:

```bash
pip install -r requirements.txt
```

## Usage

1. Prepare your Excel files:
   - Files must have the city name in the filename (e.g., `NICE_data_2024.xlsx`)
   - Column A should contain dates
   - Subsequent columns should contain allergen values
   - Supports both .xlsx and .xls formats

2. Run the script with command-line arguments:

**Basic usage (defaults to ALNUS for last 10 years):**
```bash
python extract_alnus.py "C:/path/to/files" -c NICE
```

**Specify a different allergen by name:**
```bash
python extract_alnus.py "C:/path/to/files" -c NICE -a BETULA
```

**Specify allergen by column index (0-based, column A is date):**
```bash
python extract_alnus.py "C:/path/to/files" -c NICE -a 6  # Column G (7th column)
```

**Specify number of years to plot:**
```bash
python extract_alnus.py "C:/path/to/files" -c NICE -y 5   # Last 5 years
python extract_alnus.py "C:/path/to/files" -c NICE --years 15  # Last 15 years
```

**Combine multiple options:**
```bash
python extract_alnus.py "C:/path/to/files" -c NICE -a CORYLUS -y 3
```

**Auto-detect available allergens:**
```bash
python extract_alnus.py "C:/path/to/files" -c NICE -a NONEXISTENT
# Will show available allergens and suggest correct usage
```

**Interactive mode (prompts for folder path):**
```bash
python extract_alnus.py -c NICE -a ALNUS
```

**Display help:**
```bash
python extract_alnus.py -h
```

3. The script will:
   - Find all Excel files matching the city name
   - Extract and process the data
   - If allergen not found, display available options
   - Display a summary of extracted records and year range
   - Generate and save a plot (filename based on allergen name)
   - Display the plot in a window

## Output

The script generates:
- **Console Output**: Summary of processed files, record count, and year range
- **Plot File**: A high-resolution (300 DPI) scatter plot with dynamic filename based on allergen (e.g., `alnus_plot.png`, `betula_plot.png`):
  - X-axis: Time (in month-year format)
  - Y-axis: Mean allergen values per week
  - Data points: Weekly averaged allergen values
  - Time Period: Customizable number of years (default: last 10 years)

## Command-Line Arguments

- **folder_path** (optional): Path to the folder containing Excel files. If not provided, you will be prompted to enter it.
- **-c, --city** (optional): City name to search for in filenames. Default is 'NICE'.
- **-a, --allergen** (optional): Allergen column name or index (0-based) to plot. If not found, available allergens will be displayed. Default is column 6 (ALNUS).
- **-y, --years** (optional): Number of years to plot. Default is 10.

### Allergen Column Index Reference

Column indices are 0-based:
- Index 0: Column A (Date)
- Index 1: Column B (First allergen column)
- Index 2: Column C (Second allergen column)
- ...
- Index 6: Column G (Often ALNUS)

Example: `-a 2` selects column C

## Script Parameters

To modify the script behavior by editing `extract_alnus.py`:

- **File filter**: Change `'*NICE*'` to `'*YOUR_PATTERN*'` in line 19
- **Column indices**: Change `[0, 6]` to select different columns (0-indexed: A=0, B=1, ... G=6)
- **Time range**: Change `max_year - 9` to `max_year - X` to show last X years
- **Output file**: Change `output_file='alnus_plot.png'` to a different filename

## Troubleshooting

**No files found**: Ensure files have the correct city name in the filename and are in the correct folder

**Allergen not found**: Run the script with a non-existent allergen to see available options, e.g.:
```bash
python extract_alnus.py "C:/path/to/files" -c NICE -a LIST
```

**Column index out of range**: Use `-a LIST` to detect available columns, or check that your Excel files have enough columns

**Date parsing errors**: Ensure column A contains valid date values that pandas can parse

**Missing values in output**: The script removes rows with missing date or allergen values

## Requirements File

To create a `requirements.txt` for easy installation:

```
pandas>=1.3.0
matplotlib>=3.4.0
openpyxl>=3.6.0
xlrd>=2.0.0
```

## Example

```bash
# Run the script
python extract_alnus.py

# Output example:
# Found 3 Excel file(s)
# Processing: NICE_data_2024.xlsx
# Processing: NICE_data_2023.xlsx
# Processing: NICE_data_2022.xlsx
# 
# Extracted Data Summary:
#         date  alnus  year
# 0 2016-01-04    5.2  2016
# 1 2016-01-11    6.1  2016
# 2 2016-01-18    5.8  2016
# 
# Total records: 3456
# Year range: 2016 - 2026
# Plot saved to: alnus_plot.png
```

## License

This script is provided as-is for data analysis purposes.
