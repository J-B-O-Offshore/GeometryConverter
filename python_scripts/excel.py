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
    logger.debug("hello")

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


def write_df_to_existing_table(workbook_name, sheet_name, table_name, dataframe):
    """
    Writes a DataFrame into an existing Excel table (ListObject), preserving the table formatting.

    Parameters:
    workbook_name (str): Name of the open Excel workbook (e.g. 'MyWorkbook.xlsx').
    sheet_name (str): Name of the sheet containing the table.
    table_name (str): Name of the Excel table (ListObject) to overwrite.
    dataframe (pd.DataFrame): The DataFrame to write into the Excel table.
    """
    try:
        # Connect to open workbook
        wb = next((wb for wb in xw.books if wb.name.lower() == workbook_name.lower()), None)
        if wb is None:
            raise ValueError(f"Workbook '{workbook_name}' not found.")

        sheet = next((s for s in wb.sheets if s.name.lower() == sheet_name.lower()), None)
        if sheet is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")

        # Get the Excel table object (ListObject)
        table = sheet.api.ListObjects(table_name)

        # Resize the table to fit the new DataFrame
        num_rows = len(dataframe)
        num_cols = len(dataframe.columns)

        # Set headers (if necessary)
        table.HeaderRowRange.Value = [dataframe.columns.tolist()]

        # Resize the data body (or clear it)
        if num_rows == 0:
            table.DataBodyRange.ClearContents()
        else:
            # Resize the table to the correct size
            table.Resize(table.Range.Resize(num_rows + 1, num_cols))
            table.DataBodyRange.Value = dataframe.values.tolist()

    except Exception as e:
        print(f"Error writing DataFrame to Excel table: {e}")
