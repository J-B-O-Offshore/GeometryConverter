import xlwings as xw
import os
import logging
from datetime import datetime
import pandas as pd
import numpy as np
def setup_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        # Create logs/ directory if it doesn't exist#
        log_dir = resolve_path_relative_to_script("./logs")

        os.makedirs(log_dir, exist_ok=True)

        # Create a unique log file with timestamp down to the second
        log_file = os.path.join(log_dir, f"log_{datetime.now():%Y%m%d_%H%M%S}.log")

        # File handler
        fh = logging.FileHandler(log_file, mode='w', encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        logger.addHandler(fh)

        # Optional: Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        logger.addHandler(ch)

    return logger


def set_dropdown_values(workbook_name, sheet_name, dropdown_name, items):
    """
    Modifies a Form Control dropdown (not Data Validation) by its name in a specific open workbook.

    Parameters:
        workbook_name (str): Name of the open Excel workbook (e.g., "Budget2025.xlsx").
        sheet_name (str): Name of the sheet containing the dropdown.
        dropdown_name (str): Name of the Form Control dropdown (e.g., "Drop Down 6").
        items (list of str): List of options to populate the dropdown with.

    Returns:
        True if successful, False otherwise.
    """
    try:
        for app in xw.apps:
            for wb in app.books:
                if wb.name == workbook_name:
                    try:
                        sheet = wb.sheets[sheet_name]
                        dropdown = sheet.api.Shapes(dropdown_name).ControlFormat
                        dropdown.RemoveAllItems()
                        for item in items:
                            dropdown.AddItem(item)
                        return True
                    except Exception as e:
                        print(f"Error accessing dropdown in {workbook_name}: {e}")
                        continue
    except Exception as e:
        print(f"Error while setting dropdown values: {e}")

    print(f"Workbook '{workbook_name}' or dropdown '{dropdown_name}' not found.")
    return False


import xlwings as xw


def resolve_path_relative_to_script(path):
    """
    Resolves a given path relative to the script's directory.
    If the path is already absolute, it is returned unchanged.

    Parameters:
    - path: str, the path to resolve

    Returns:
    - str, the absolute path resolved relative to this script
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return path if os.path.isabs(path) else os.path.abspath(os.path.join(script_dir, path))


def write_df(workbook_name, sheet_name, upper_left_address, dataframe, include_headers=True):
    """
     Writes a Pandas DataFrame to an already open Excel workbook, starting at a given cell address or named range.

     Parameters:
     workbook_name (str): Name of the open Excel workbook (e.g. 'GeometrieConverter.xlsx').
     sheet_name (str): Name of the sheet within the workbook.
     upper_left_address (str): Excel address (e.g., 'A1') or named range where the DataFrame should be placed.
     dataframe (pd.DataFrame): The Pandas DataFrame to write to the Excel sheet.
     include_headers (bool): Whether to include column names as headers in Excel (default is True).
     """
    logger = setup_logger()

    try:
        # Connect to the already open workbook
        wb = xw.books[workbook_name]
        sheet = wb.sheets[sheet_name]

        # Check if upper_left_address is a named range
        try:
            upper_left_range = wb.names[upper_left_address].refers_to_range
        except KeyError:
            # Not a named range, treat it as a cell address
            upper_left_range = sheet.range(upper_left_address)

        # Prepare data
        if include_headers:
            data = [dataframe.columns.tolist()] + dataframe.values.tolist()
        else:
            data = dataframe.values.tolist()

        # Write data
        upper_left_range.value = data

    except Exception as e:
        print(f"Error writing DataFrame to Excel: {e}")
        logger.debug(f"Error!{e}")


def write_df_to_table(workbook_name, sheet_name, table_name, dataframe):
    """
    Replace the contents of an existing Excel table with a pandas DataFrame using xlwings.

    Parameters:
    - workbook_name: str, name of the open Excel workbook (no path needed if open).
    - sheet_name: str, name of the sheet containing the table.
    - table_name: str, name of the Excel table (ListObject) to manipulate.
    - dataframe: pandas DataFrame to replace the table's data (same number of columns).
    """
    # Connect to the open workbook
    wb = xw.books[workbook_name]
    ws = wb.sheets[sheet_name]

    # Find the table (ListObject)
    try:
        table = ws.api.ListObjects(table_name)
    except Exception as e:
        raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.") from e

    # Get the header range and data body range
    header_range = table.HeaderRowRange
    data_body_range = table.DataBodyRange

    # Clear the existing table data (keep headers)
    data_body_range.ClearContents()

    # Write the new DataFrame below the headers
    start_cell = ws.range((header_range.Row + 1, header_range.Column))
    start_cell.options(index=False, header=False).value = dataframe

    # Resize the table to match new data
    last_row = header_range.Row + dataframe.shape[0]
    last_col = header_range.Column + dataframe.shape[1] - 1
    new_range = ws.range(
        (header_range.Row, header_range.Column),
        (last_row, last_col)
    )
    table.Resize(new_range.api)


def show_message_box(workbook_name, message):
    """
    Displays a message box in the specified Excel workbook and returns a boolean value
    based on the user's choice (Yes or No).

    Args:
        workbook_name (str): The name of the workbook (with .xlsm extension) that is already open.
        message (str): The message to display in the message box.

    Returns:
        bool: True if the user clicks 'Yes', False if the user clicks 'No'.
              Returns None if the workbook is not open or if an error occurs.

    Raises:
        KeyError: If the specified workbook is not found among the open workbooks.
        AttributeError: If no active Excel instance is found.
    """

    # Attempt to connect to an existing Excel instance
    try:
        app = xw.apps.active  # Use the active Excel application instance
    except AttributeError:
        print("No active Excel instance found.")
        return None

    # Try to find the workbook by name
    try:
        wb = app.books[workbook_name]
    except KeyError:
        print(f"Workbook '{workbook_name}' is not open.")
        return None

    # Define the VBA code for the message box
    vba_code = f"""
    Function ShowMsgBox() As Boolean
        Dim answer As Integer
        answer = MsgBox("{message}", vbYesNo + vbQuestion, "Choice Box")
        If answer = vbYes Then
            ShowMsgBox = True
        Else
            ShowMsgBox = False
        End If
    End Function
    """

    # Access the VBA project and add the code
    vb_module = wb.api.VBProject.VBComponents.Add(1)  # Add a new module
    vb_module.CodeModule.AddFromString(vba_code)  # Add the VBA code to the module

    # Run the VBA function and capture the return value
    result = wb.macro('ShowMsgBox')()

    return result

def read_excel_table(workbook_name, sheet_name, table_name):
    """
    Read an Excel Table into a Pandas DataFrame, using the Table's header as column names.

    Parameters:
        workbook_name (str): The name of the workbook
        sheet_name (str): The name of the sheet containing the table.
        table_name (str): The name of the Excel Table (not the range name).

    Returns:
        pd.DataFrame: DataFrame containing the table data with correct headers.
    """
    wb = xw.Book(workbook_name)

    sheet = wb.sheets[sheet_name]
    table = sheet.tables[table_name]

    # Read the data body into a DataFrame
    df = table.data_body_range.options(pd.DataFrame, header=False, index=False).value

    # Set the correct headers from the table's header row
    df.columns = [h.strip() for h in table.header_row_range.value]

    return df

def add_unique_row(df1, df2, exclude_columns=None):
    """
    Adds a row from `df1` to `df2` if no equivalent row already exists,
    with optional exclusion of specific columns from the comparison.

    Two rows are considered equivalent if all non-excluded column values match,
    treating NaN and None as equal and attempting to convert numerical values to float
    before comparison.

    Parameters
    ----------
    df1 : pandas.DataFrame
        A single-row DataFrame representing the row to add.
    df2 : pandas.DataFrame
        The target DataFrame to which the row may be added.
    exclude_columns : list of str, optional
        List of column names to ignore when checking for matching rows.
        Defaults to an empty list.

    Returns
    -------
    df2 : pandas.DataFrame
        The updated DataFrame after adding the row if it was unique.
    matching_indices : list of int
        List of indices in `df2` where matching rows were found.
        Empty if the row was added.

    Notes
    -----
    - Missing values (NaN, None) are treated as equivalent during comparison.
    - If the columns of `df1` and `df2` differ, missing columns are filled with None.
    - All values are cast to string after float conversion to ensure robust comparison.
    """
    if exclude_columns is None:
        exclude_columns = []

    all_columns = df1.columns.union(df2.columns)
    df1 = df1.reindex(columns=all_columns, fill_value=None)
    df2 = df2.reindex(columns=all_columns, fill_value=None)

    # Drop excluded columns from both df1 and df2 for comparison
    df1_comp = df1.drop(columns=exclude_columns, errors='ignore')
    df2_comp = df2.drop(columns=exclude_columns, errors='ignore')

    # Convert both DataFrames to treat NaN and None as equal and convert numerical values (including strings) to float
    def convert_to_float(x):
        try:
            return float(x)
        except (ValueError, TypeError):
            return np.nan if pd.isna(x) else x

    df1_comp = df1_comp.map(convert_to_float)
    df2_comp = df2_comp.map(convert_to_float)

    # Ensure all columns are of the same type for proper comparison
    df1_comp = df1_comp.astype(str)
    df2_comp = df2_comp.astype(str)

    # Check for rows in df2 that match the row in df1
    matching_rows = df2_comp[df2_comp.eq(df1_comp.values[0]).all(axis=1)]


    if not matching_rows.empty:
        matching_indices = matching_rows.index.tolist()
    else:
        df2 = pd.concat([df2, df1], ignore_index=True)
        matching_indices = []

    return df2, matching_indices