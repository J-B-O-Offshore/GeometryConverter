import xlwings
import pandas as pd
import sqlite3

import excel as ex


class ConciveError(Exception):
    """
    Custom exception for errors related to database operations
    in the context of structure data handling.
    """
    pass


def load_db_table(db_path, table_name):
    """
    Load a specific table from an SQLite database into a pandas DataFrame.

    Args:
        db_path (str): Path to the SQLite database file.
        table_name (str): Name of the table to load.

    Returns:
        pd.DataFrame: The table content as a pandas DataFrame.

    Raises:
        ConciveError: If database connection fails or table does not exist.
    """
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to connect to the database: {e}")

    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    table_names = [table[0] for table in cursor.fetchall()]

    if table_name not in table_names:
        raise ConciveError(f"Table '{table_name}' does not exist in the database.")

    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql_query(query, conn)
    conn.close()

    return df


def load_MP_META(db_path):
    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the MP SQLite database.

    Returns:
        None
    """
    logger = ex.setup_logger()
    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    ex.set_dropdown_values(
        "GeometrieConverter.xlsm",
        sheet_name_structure_loading,
        "Dropdown_MP_Structures",
        list(META.loc[:, "table_name"].values)
    )


def load_MP_DATA(Structure_name, db_path):
    """
    Load metadata and structure-specific data from the MP database
    and write them to the Excel workbook.

    Args:
        Structure_name (str): Name of the structure (table) to load.
        db_path (str): Path to the MP SQLite database.

    Returns:
        None
    """
    logger = ex.setup_logger()
    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    DATA = load_db_table(db_path, Structure_name)

    META_relevant = META.loc[META["table_name"] == Structure_name]

    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "MP_META", META_relevant)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "MP_DATA_TRUE", DATA)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "MP_DATA", DATA)


def load_TP_META(db_path):
    """
    Load the META table from the TP database and update the TP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the TP SQLite database.

    Returns:
        None
    """
    logger = ex.setup_logger()
    sheet_name_structure_loading = "BuildYourStructure"

    logger.debug(f"TP_META db path: {db_path}")
    META = load_db_table(db_path, "META")
    logger.debug(f"META: {META.to_string()}")

    ex.set_dropdown_values(
        "GeometrieConverter.xlsm",
        sheet_name_structure_loading,
        "Dropdown_TP_Structures",
        list(META.loc[:, "table_name"].values)
    )


def load_TP_DATA(Structure_name, db_path):
    """
    Load metadata and structure-specific data from the TP database
    and write them to the Excel workbook.

    Args:
        Structure_name (str): Name of the structure (table) to load.
        db_path (str): Path to the TP SQLite database.

    Returns:
        None
    """
    logger = ex.setup_logger()
    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    DATA = load_db_table(db_path, Structure_name)

    META_relevant = META.loc[META["table_name"] == Structure_name]

    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "TP_META", META_relevant)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "TP_DATA_TRUE", DATA)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "TP_DATA", DATA)

# def save_MP_data(db_path):
#     logger = ex.setup_logger()
#
#     META_DB = load_db_table(db_path, "META")
#     META_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "MP_META")
#
#     DATA_DB = load_db_table(db_path, "Data")
#
#
#     META_DB = META_DB.drop(columns=["index"])
#     META_CURR = META_CURR.drop(columns=["index"])
#
#     match = ((META_DB == META_CURR.iloc[0]).all(axis=1)).any()
#
#     if match:
#
#         print(1)
#     logger.debug(f"match: {match}")
#
#
#
#     ex.add_unique_row(META_DB, META_CURR)
#
#     return
