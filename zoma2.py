import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from datetime import datetime
import sys
import time

def remove_null_columns(filename):
    start_time = time.time()  # Record start time

    # Check file type and read into DataFrame
    if filename.endswith('.csv'):
        df = pd.read_csv(filename)
    elif filename.endswith('.xlsx'):
        df = pd.read_excel(filename)
    else:
        print("Unsupported file format. Please provide a CSV or XLSX file.")
        return

    # Check if the number of columns is more than 300
    if len(df.columns) > 300:
        print("The file has more than 300 columns. Exiting the program.")
        return

    # Find 'eventTime' column and move it to the first column
    if 'eventTime' in df.columns:
        cols = list(df.columns)
        cols.insert(0, cols.pop(cols.index('eventTime')))
        df = df[cols]

    # Identify columns with only 'null' or NaN values and delete them, except 'productAction'
    columns_to_delete = []
    for col in df.columns:
        # Replace NaN with 'null' for comparison purposes
        df[col] = df[col].fillna('null')
        unique_values = df[col].unique()

        if len(unique_values) == 1 and str(unique_values[0]).strip().lower() == 'null':
            if col != 'productAction':
                columns_to_delete.append(col)

    # Debugging output
    print("Columns to delete:", columns_to_delete)

    # Drop identified columns
    df = df.drop(columns=columns_to_delete)

    # Additional processing for columns starting with 'aisaac' and ending with 'Name'
    i = 0
    while i < len(df.columns) - 1:
        current_col = df.columns[i]
        next_col = df.columns[i + 1]

        if current_col.startswith('aisaac') and current_col.endswith('Name'):
            if next_col + 'Name' != current_col:
                df = df.drop(columns=[current_col])
            else:
                if df[current_col].notna().all() and df[current_col].nunique() == 1 and str(df[current_col].iloc[0]).strip().lower() != 'null':
                    unique_value = df[current_col].iloc[0]
                    df.rename(columns={next_col: unique_value}, inplace=True)
                    df = df.drop(columns=[current_col])
                else:
                    i += 1
        else:
            i += 1

    # Create an Excel workbook and add the DataFrame
    wb = Workbook()
    ws = wb.active

    # Add DataFrame to worksheet, keeping 'null' as 'null'
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Save the modified workbook
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_filename = filename.replace('.csv', f'_{timestamp}.xlsx') if filename.endswith('.csv') else filename.replace('.xlsx', f'_{timestamp}.xlsx')
    try:
        wb.save(output_filename)
        elapsed_time = time.time() - start_time  # Calculate elapsed time
        print(f"Modified workbook saved as '{output_filename}'.\nElapsed Time: {elapsed_time:.2f} seconds")
    except PermissionError:
        print("Permission denied. Please make sure you have write access to the directory.")
        return

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script_name.py file_name.csv/xlsx")
    else:
        remove_null_columns(sys.argv[1])