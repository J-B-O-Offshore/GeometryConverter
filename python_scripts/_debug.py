import excel as ex
import xlwings as xw
import pandas as pd
sheet_name_structure_loading = "BuildYourStructure"
import sqlite3
#
excel_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/GeometrieConverter.xlsm"
workbook_name = "GeometrieConverter.xlsm"
sheet_name = "BuildYourStructure"
table_name = "MP_DATA"
message = "Duh?"
db_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/MP.db"
selected_structure = "MP02"
Structure = "MP"

class ConciveError(Exception):
    """
    Custom exception for errors related to database operations
    in the context of structure data handling.
    """
    pass

def drop_db_table(db_path, table_name):
    """
    Drop (delete) a table from an SQLite database completely.

    Args:
        db_path (str): Path to the SQLite database file.
        table_name (str): Name of the table to drop.

    Raises:
        ConciveError: If database connection fails or drop operation fails.
    """
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to connect to the database: {e}")

    cursor = conn.cursor()

    try:
        cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
        conn.commit()
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to drop table '{table_name}': {e}")
    finally:
        conn.close()
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

def create_db_table(db_path, table_name, df, if_exists='fail'):
    """
    Create a new table in an SQLite database from a pandas DataFrame.

    Args:
        db_path (str): Path to the SQLite database file.
        table_name (str): Name of the table to create.
        df (pd.DataFrame): DataFrame containing the data to write.
        if_exists (str): What to do if the table already exists.
                         Options: 'fail', 'replace', 'append'.

    Raises:
        ConciveError: If database connection fails or table creation fails.
    """
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to connect to the database: {e}")

    try:
        df.to_sql(table_name, conn, if_exists=if_exists, index=False)
    except Exception as e:
        raise ConciveError(f"Failed to create or write to the table '{table_name}': {e}")
    finally:
        conn.close()

def save_data(Structure, db_path, selected_structure):

    logger = ex.setup_logger()

    META_FULL = load_db_table(db_path, "META")

    META_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META")
    META_CURR_NEW = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_NEW")
    META_DB = META_FULL.loc[META_FULL["table_name"] == selected_structure]

    DATA_DB = load_db_table(db_path, selected_structure)
    DATA_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA")

    META_DB = META_DB.drop(columns=["index"])
    META_CURR = META_CURR.drop(columns=["index"])
    DATA_DB = DATA_DB.astype(float)
    DATA_CURR = DATA_CURR.astype(float)

    # changetracks
    data_changed = not(DATA_DB.equals(DATA_CURR))
    meta_loaded_changed = not (META_DB.values == META_CURR.values).all()


    meta_new_populated = (META_CURR_NEW.values != None).any()

    # create new database table

    # check, if there are values in meta_new first
    if meta_new_populated:
        # check if meta data is complete
        if ~((META_CURR_NEW.values != None).all()):
            ex.show_message_box("GeometrieConverter.xlsm", "Please fully populate the NEW Meta table to create a new DB entry or clear it of all data to overwrite the loaded Structure")
            return False
        # check if data has changed
        if ~data_changed:
            ex.show_message_box("GeometrieConverter.xlsm", "The provided data is the same as the data in the current Structure. Aborting.")
            return False
        # check if table name already exists
        new_table_name = META_CURR_NEW["table_name"].values[0]
        if new_table_name in META_DB["table_name"].values:
            ex.show_message_box("GeometrieConverter.xlsm", "The provided table_name of the new structure is already in in the database, please pride a unique name")
            return False

        # creating new table
        create_db_table(db_path, new_table_name, DATA_CURR)

        # adding meta info
        META_FULL = pd.concat([META_FULL, META_CURR_NEW])
        create_db_table(db_path, "META", META_FULL, if_exists='replace')

        ex.show_message_box("GeometrieConverter.xlsm", f"Data saved in new Database entry {new_table_name}")

        return True

    if meta_loaded_changed:
        curr_table_name = META_CURR["table_name"].values[0]

        # check if meta data is complete
        if ~((META_CURR.values != None).all()):
            ex.show_message_box("GeometrieConverter.xlsm", "Please fully populate the Current Meta table to modify the DB entry.")
            return False

        # check if new table name is valid
        if META_DB["table_name"].value_counts().get(curr_table_name, 0) > 1:
            ex.show_message_box("GeometrieConverter.xlsm", f"{curr_table_name} is already taken in database, please choose a different name.")
            return False

        # replace row in META table in db
        META_DB.loc[META_DB["table_name"] == selected_structure] = META_CURR.iloc[0]

        # write new table to database
        drop_db_table(db_path, selected_structure)
        create_db_table(db_path, curr_table_name, DATA_CURR)

        ex.show_message_box("GeometrieConverter", f"Data in {selected_structure} overwriten, now named {curr_table_name}")

        return True

    if data_changed:

        if any(pd.isna(DATA_CURR.values)):
            ex.show_message_box("GeometrieConverter.xlsm", f"Modified data contains invalid values, please correct")
            return False

        drop_db_table(db_path, selected_structure)
        create_db_table(db_path, selected_structure, DATA_CURR)

        return True

    ex.show_message_box("GeometrieConverter.xlsm", f"No changes detected.")
    return False

def save_MP_data(db_path, selected_structure):

    save_data("MP", db_path, selected_structure)

    return

save_MP_data(db_path, selected_structure)


print(1)
