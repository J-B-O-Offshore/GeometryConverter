import xlwings
import pandas as pd
import sqlite3

import excel as ex


class ConciveError(Exception):
    pass


def load_db_table(db_path, table_name):
    # Check if db_path is an actual database file
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to connect to the database: {e}")

    # Get the list of all table names in the database
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    table_names = [table[0] for table in cursor.fetchall()]

    # Check if table_name exists in the database
    if table_name not in table_names:
        raise ConciveError(f"Table '{table_name}' does not exist in the database.")

    # Load the table into a DataFrame
    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql_query(query, conn)

    # Close the connection
    conn.close()

    return df


def MP_META(db_path):
    logger = ex.setup_logger()

    sheet_name_structure_loading = "BuildYourStructure"

    logger.debug(db_path)
    META = load_db_table(db_path, "META")

    ex.set_dropdown_values("GeometrieConverter.xlsm", sheet_name_structure_loading, "Dropdown_MP_Structures", list(META.loc[:, "table_name"].values))

    return


def MP_DATA(Structure_name, db_path):
    logger = ex.setup_logger()

    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    DATA = load_db_table(db_path, Structure_name)

    META_relevant = META.loc[META["table_name"] == Structure_name]
    # ex.write_df_to_existing_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "MP_Data", META_relevant)

    ex.write_df("GeometrieConverter.xlsm", sheet_name_structure_loading, "Table_MP_META", META_relevant)
    ex.write_df("GeometrieConverter.xlsm", sheet_name_structure_loading, "Table_MP_DATA", DATA)

    return


def TP_META(db_path):
    logger = ex.setup_logger()

    sheet_name_structure_loading = "BuildYourStructure"

    logger.debug(db_path)
    META = load_db_table(db_path, "META")

    ex.set_dropdown_values("GeometrieConverter.xlsm", sheet_name_structure_loading, "Dropdown_TP_Structures", list(META.loc[:, "table_name"].values))

    return


def TP_DATA(Structure_name, db_path):
    logger = ex.setup_logger()

    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    DATA = load_db_table(db_path, Structure_name)

    META_relevant = META.loc[META["table_name"] == Structure_name]
    # ex.write_df_to_existing_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "MP_Data", META_relevant)

    ex.write_df("GeometrieConverter.xlsm", sheet_name_structure_loading, "Table_TP_META", META_relevant)
    ex.write_df("GeometrieConverter.xlsm", sheet_name_structure_loading, "Table_TP_DATA", DATA)

    return
