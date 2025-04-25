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




def fill_MP_section():

    return


def fill_TP_section():

    return

def fill_TOWER_section():

    return

def fill_TURBINE_section():

    return