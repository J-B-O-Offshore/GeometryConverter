import xlwings as xw
import os
import logging
from datetime import datetime


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


import xlwings as xw

import xlwings as xw

import xlwings as xw


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