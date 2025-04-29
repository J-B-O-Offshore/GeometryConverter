import excel as ex
import xlwings as xw
import pandas as pd
sheet_name_structure_loading = "BuildYourStructure"

#
excel_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/GeometrieConverter.xlsm"
workbook_name = "GeometrieConverter.xlsm"
sheet_name = "BuildYourStructure"
table_name = "MP_DATA"
message = "Duh?"


def read_excel_table(workbook_name, sheet_name, table_name):
    """
    Read an Excel Table into a Pandas DataFrame, using the Table's header as column names.

    Parameters:
        workbook_name (xw.Book): The name of the workbook
        sheet_name (str): The name of the sheet containing the table.
        table_name (str): The name of the Excel Table (not the range name).

    Returns:
        pd.DataFrame: DataFrame containing the table data with correct headers.
    """
    wb = xw.Book(workbook_name)

    sheet = wb.sheets[sheet_name]
    table = sheet.tables[table_name]

    # Read the data body into a DataFrame
    df = table.data_body_range.options(pd.DataFrame, header=True, index=False).value

    # Set the correct headers from the table's header row
    df.columns = [h.strip() for h in table.header_row_range.value]

    return df
# Example usage:
#result = read_table_to_df(workbook_name, sheet_name, table_name)


df = read_excel_table(workbook_name, sheet_name, table_name)

print(1)
