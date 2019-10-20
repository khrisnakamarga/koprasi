# Koprasi
Generates balance reports for PT. Kamarga Kurnia Textile's employee. Sorts the input file to its respective NIK (employee number), and records transaction's date, amount, type, and current employee balance.

## Installing requirements
1. Python version: Python 3.6
2. Libraries: `pip install -r requirements.txt`

## Generating the report
1. Navigate to the repository folder
2. Put the excel sheet in the repository with the name: "toBeParsed.xlsx"
3. Run the script: `python transaction_parser.py`
4. The sheet: "nik_report.xlsx" will be generated
