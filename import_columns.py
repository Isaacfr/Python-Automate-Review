import pandas as pd
from openpyxl import load_workbook

def import_columns_function(source_file, raw_sheet, columns_to_copy, columns_to_modify, new_sheet_name):
    """
    Copies specified columns from a raw sheet in a source Excel file to a new sheet in the same file.

    Parameters:
    - source_file: str, path to the source Excel file
    - raw_sheet_name: str, name of the raw sheet to copy from
    - columns_to_copy: list of str, names of the columns to copy
    - new_sheet_name: str, name of the new sheet to create
    """

    # Load the raw sheet into a DataFrame
    try:
        df = pd.read_excel(source_file, sheet_name=raw_sheet)
    except Exception as e:
        print(f"Error reading the raw sheet: {e}")
        return

    # Check if the specified columns exist in the DataFrame
    missing_columns = [col for col in columns_to_copy if col not in df.columns]
    if missing_columns:
        print(f"Missing columns in the raw sheet: {missing_columns}")
        return

    # Create a new DataFrame with the specified columns
    new_df = df[columns_to_copy].copy()

    #Modify specified column by subtracting 1
    new_df[columns_to_modify] = new_df[columns_to_modify] - pd.Timedelta(days=1)

    # Load the existing Excel file and write the new DataFrame to a new sheet
    with pd.ExcelWriter(source_file, engine='openpyxl', mode='a') as writer:
        new_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

    print(f"Columns {columns_to_copy} copied to new sheet '{new_sheet_name}' successfully.")

