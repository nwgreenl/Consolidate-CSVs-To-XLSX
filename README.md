## Consolidate CSVs to XLSX
Simple python script that consolidates multiple CSVs into a single Excel file (with multiple sheets)

**Example**

- Input: `hello.csv` and `world.csv` 
- Output: `consolidated_{date}.xlsx` with `hello` and `world` sheets.

### Requirements
- Python3
- Pandas (`pip3 install pandas`)
- xlsxwriter (`pip3 install xlsxwriter`)

### Usage
Run `main.py` and select the CSVs you wish to consolidate via the file dialog. 

- Customize output directory & file name via `get_outputfile(filename, dir_name, append_date)`
- Defaults are `("consolidated", "output", True)`