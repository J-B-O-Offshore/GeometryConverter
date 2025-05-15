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

    # if table_name not in table_names:
    #     raise ConciveError(f"Table '{table_name}' does not exist in the database.")

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


def load_META(Structure, db_path):
    """
    Load the META table from the  database and update the structures dropdown
    in the Excel workbook.

    Args:
        Structure (str): Name of the structure to load (MP, TP,...)
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
        f"Dropdown_{Structure}_Structures",
        list(META.loc[:, "table_name"].values)
    )
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_META_FULL", META)
    return


def load_DATA(Structure, Structure_name, db_path):
    """
    Load metadata and structure-specific data from the database
    and write them to the Excel workbook.

    Args:
        Stucture (str): Name of the structure to load (MP, TP,...)
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

    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_META_TRUE", META_relevant)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_META", META_relevant)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_DATA_TRUE", DATA)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_DATA", DATA)


def save_data(Structure, db_path, selected_structure):
    """
    Saves data from an Excel sheet and updates a database based on changes in structure metadata and data.

    This function loads the full metadata and data for a given structure, compares the current data with the existing data in the
    database, and handles the saving of new data or updating existing entries in the database. The function follows a series of checks
    to validate data, handle new entries, overwrite existing entries, and ensure data integrity.

    The function performs the following:
    1. Validates and processes the current data and metadata.
    2. Checks if new metadata is fully populated and whether it represents a new structure or an update.
    3. Saves the data to a new database table if the structure is new, or overwrites the existing structure if the metadata has changed.
    4. Displays message boxes to inform the user of the progress, errors, or status of the operation.

    Parameters:
    -----------
    Structure : str
        The name of the structure whose data and metadata need to be saved or updated.
    db_path : str
        The path to the database where the data and metadata will be saved.
    selected_structure : str
        The name of the structure that is selected for saving or updating.

    Returns:
    --------
    bool
        A boolean indicating whether the data was successfully saved or updated.
    str
        The name of the structure after the operation (could be a new name or the same).
    """

    META_FULL = load_db_table(db_path, "META")
    META_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META")
    META_CURR_NEW = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_NEW")
    META_DB = META_FULL.loc[META_FULL["table_name"] == selected_structure]
    DATA_DB = load_db_table(db_path, selected_structure)
    DATA_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA")

    META_DB = META_DB.drop(columns=["index"])
    META_CURR = META_CURR.drop(columns=["index"])
    META_FULL = META_FULL.drop(columns=["index"])
    META_CURR_NEW = META_CURR_NEW.drop(columns=["index"])

    def saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR, selected_structure):

        def valid_data(data):
            if pd.isna(data.values).any():
                return False, data
            try:
                return True, data.astype(float)
            except (ValueError, TypeError):
                return False, data

        succes_1, DATA_DB = valid_data(DATA_DB)
        succes_2, DATA_CURR = valid_data(DATA_CURR)

        if not succes_2:
            _ = ex.show_message_box("GeometrieConverter.xlsm",
                                    "Invalid data found in Structure data! Aborting.")

            return False, _

        data_changed = not (DATA_DB.equals(DATA_CURR))
        meta_loaded_changed = not (META_DB.values == META_CURR.values).all()
        meta_new_populated = (META_CURR_NEW.values != None).any()

        if meta_new_populated:
            if not ((META_CURR_NEW.values != None).all()):
                _ = ex.show_message_box("GeometrieConverter.xlsm",
                                        "Please fully populate the NEW Meta table to create a new DB entry or clear it of all data to overwrite the loaded Structure")
                return False, _
            if not data_changed:
                _ = ex.show_message_box("GeometrieConverter.xlsm", "The provided data is the same as the data in the current Structure. Aborting.")
                return False, _

            new_table_name = META_CURR_NEW["table_name"].values[0]
            if new_table_name in META_FULL["table_name"].values:
                _ = ex.show_message_box("GeometrieConverter.xlsm", "The provided table_name of the new structure is already in the database, please provide a unique name")
                return False, _

            create_db_table(db_path, new_table_name, DATA_CURR)
            META_FULL = pd.concat([META_FULL, META_CURR_NEW])
            META_FULL.insert(0, 'index', META_FULL.index.values)
            create_db_table(db_path, "META", META_FULL, if_exists='replace')

            _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data saved in new Database entry {new_table_name}")
            structure_load_after = new_table_name
            return True, structure_load_after

        if meta_loaded_changed:
            curr_table_name = META_CURR["table_name"].values[0]
            if not ((META_CURR.values != None).all()):
                _ = ex.show_message_box("GeometrieConverter.xlsm", "Please fully populate the Current Meta table to modify the DB entry.")
                return False, _

            if (curr_table_name != selected_structure) and (curr_table_name in META_FULL["table_name"].values):
                _ = ex.show_message_box("GeometrieConverter.xlsm", f"{curr_table_name} is already taken in database, please choose a different name.")
                return False, _

            answer = ex.show_message_box("GeometrieConverter.xlsm", f"Sure you want to overwrite the data stored in {selected_structure}?", icon="vbYesNo", buttons="vbYesNo")
            if answer == "No":
                return False, None

            META_FULL.loc[META_FULL["table_name"] == selected_structure, :] = META_CURR.iloc[0].values
            drop_db_table(db_path, selected_structure)
            create_db_table(db_path, curr_table_name, DATA_CURR, if_exists='replace')

            META_FULL.insert(0, 'index', META_FULL.index.values)
            create_db_table(db_path, "META", META_FULL, if_exists='replace')

            _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data in {selected_structure} overwriten, now named {curr_table_name}")
            structure_load_after = curr_table_name
            return True, structure_load_after

        if data_changed:
            if pd.isna(DATA_CURR.values).any():
                _ = ex.show_message_box("GeometrieConverter.xlsm", f"Modified data contains invalid values, please correct")
                return False, _

            answer = ex.show_message_box("GeometrieConverter.xlsm", f"Sure you want to overwrite the data stored in {selected_structure}?", icon="vbYesNo", buttons="vbYesNo")
            if answer == "No":
                return False, None

            drop_db_table(db_path, selected_structure)
            create_db_table(db_path, selected_structure, DATA_CURR, if_exists='replace')

            _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data in current Structure saved to the database.")

            structure_load_after = selected_structure
            return True, structure_load_after

        _ = ex.show_message_box("GeometrieConverter.xlsm", f"No changes detected.")
        return False, _

    saved, structure_load_after = saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR, selected_structure)

    if saved:
        load_META(Structure, db_path)
        load_DATA(Structure, structure_load_after, db_path)


def delete_data(Structure, db_path, selected_structure):
    META_FULL = load_db_table(db_path, "META")
    answer = ex.show_message_box("GeometrieConverter.xlsm", f"Are you sure you want to delete the structure {selected_structure} from the database?", icon="vbYesNo",
                                 buttons="vbYesNo")
    logger = ex.setup_logger()
    logger.debug(answer)
    if answer == "Yes":
        META_FULL.drop(META_FULL[META_FULL["table_name"] == selected_structure].index, inplace=True)

        drop_db_table(db_path, selected_structure)
        create_db_table(db_path, "META", META_FULL, if_exists='replace')
        load_META(Structure, db_path)

        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_NEW")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_TRUE")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA_TRUE")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA")

    return


# %% MP
def load_MP_META(db_path):
    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the MP SQLite database.

    Returns:
        None
    """
    load_META("MP", db_path)


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

    load_DATA("MP", Structure_name, db_path)


def save_MP_data(db_path, selected_structure):
    save_data("MP", db_path, selected_structure)

    return


def delete_MP_data(db_path, selected_structure):
    delete_data("MP", db_path, selected_structure)

    return


# %% TP
def load_TP_META(db_path):
    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the TP SQLite database.

    Returns:
        None
    """
    load_META("TP", db_path)


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

    load_DATA("TP", Structure_name, db_path)


def save_TP_data(db_path, selected_structure):
    save_data("TP", db_path, selected_structure)

    return


def delete_TP_data(db_path, selected_structure):
    delete_data("TP", db_path, selected_structure)

    return


# %% TOWER
def load_TOWER_META(db_path):
    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the TOWER SQLite database.

    Returns:
        None
    """
    load_META("TOWER", db_path)


def load_TOWER_DATA(Structure_name, db_path):
    """
    Load metadata and structure-specific data from the TOWER database
    and write them to the Excel workbook.

    Args:
        Structure_name (str): Name of the structure (table) to load.
        db_path (str): Path to the TOWER SQLite database.

    Returns:
        None
    """

    load_DATA("TOWER", Structure_name, db_path)


def save_TOWER_data(db_path, selected_structure):
    save_data("TOWER", db_path, selected_structure)

    return


def delete_TOWER_data(db_path, selected_structure):
    delete_data("TOWER", db_path, selected_structure)

    return
