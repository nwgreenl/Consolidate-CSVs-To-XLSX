## Consolidate CSVs to XLSX
Simple python script that consolidates multiple CSVs into a single Excel file (with multiple sheets).

**Example**

- Input: `hello.csv` and `world.csv` 
- Output: `consolidated_{date}.xlsx` with `hello` and `world` sheets.

### Requirements
- Python3
- Pandas (`pip3 install pandas`)
- xlsxwriter (`pip3 install xlsxwriter`)

### Usage
- Run `main.py` and select the CSVs you wish to consolidate via the file dialog. 
- Assumes first row is the header.

### Customization
Output directory & file name via the final call to `get_outputfile()`

- Defaults are `("consolidated", "output", True)`

Header styling via `header_format`. The spec can be found [here](https://xlsxwriterlua.readthedocs.io/working_with_formats.html)

- Defaults are set to be center alignment (vertical and horizontal), no borders, bold text, & text wrapping.