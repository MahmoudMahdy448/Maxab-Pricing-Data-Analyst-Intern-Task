import pandas as pd

def extract_sheets_to_csv(input_file):
    # Read the Excel file
    xls = pd.ExcelFile(input_file)

    # Get the sheet names (last 3 sheets in our case contains the data we need)
    sheet_names = xls.sheet_names[-3:]

    # Extract each sheet to a separate CSV file
    for sheet_name in sheet_names:
        df = pd.read_excel(xls, sheet_name)
        output_file = f"{sheet_name}.csv"
        df.to_csv(output_file, index=False, encoding="utf-8")
        print(f"Saved {output_file}")

if __name__ == "__main__":
    input_excel_file = "Data.xlsx"  
    extract_sheets_to_csv(input_excel_file)
