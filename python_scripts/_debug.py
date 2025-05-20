import pandas as pd

import excel as ex

sheet_name_structure_loading = "BuildYourStructure"
import sqlite3

#
excel_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/GeometrieConverter.xlsm"
workbook_name = "GeometrieConverter.xlsm"
sheet_name = "BuildYourStructure"
Identifier = "MP_DATA"
message = "Duh?"
db_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/MP.db"
selected_structure = "24A535_FEED_DP-A1_L0_G0_S0"
Structure = "MP"


class ConciveError(Exception):
    """
    Custom exception for errors related to database operations
    in the context of structure data handling.
    """
    pass


def drop_db_table(db_path, Identifier):
    """
    Drop (delete) a table from an SQLite database completely.

    Args:
        db_path (str): Path to the SQLite database file.
        Identifier (str): Name of the table to drop.

    Raises:
        ConciveError: If database connection fails or drop operation fails.
    """
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to connect to the database: {e}")

    cursor = conn.cursor()

    try:
        cursor.execute(f'DROP TABLE IF EXISTS "{Identifier}"')
        conn.commit()
    except sqlite3.Error as e:
        raise ConciveError(f"Failed to drop table '{Identifier}': {e}")
    finally:
        conn.close()


def load_db_table(db_path, Identifier):
    """
    Load a specific table from an SQLite database into a pandas DataFrame.

    Args:
        db_path (str): Path to the SQLite database file.
        Identifier (str): Name of the table to load.

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

    # if Identifier not in table_names:
    #     raise ConciveError(f"Table '{Identifier}' does not exist in the database.")

    query = f'SELECT * FROM "{Identifier}"'
    df = pd.read_sql_query(query, conn)
    conn.close()

    df.drop(columns=['index'], inplace=True)

    return df


def create_db_table(db_path, Identifier, df, if_exists='fail'):
    """
    Create a new table in an SQLite database from a pandas DataFrame.

    Args:
        db_path (str): Path to the SQLite database file.
        Identifier (str): Name of the table to create.
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
        df.to_sql(Identifier, conn, if_exists=if_exists, index=False)
    except Exception as e:
        raise ConciveError(f"Failed to create or write to the table '{Identifier}': {e}")
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
        list(META.loc[:, "Identifier"].values)
    )
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
    MASSES = load_db_table(db_path, Structure_name+"__ADDED_MASSES")

    META_relevant = META.loc[META["Identifier"] == Structure_name]

    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_META_TRUE", META_relevant)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_META", META_relevant)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_DATA_TRUE", DATA)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_DATA", DATA)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_MASSES_TRUE", MASSES)
    ex.write_df_to_table("GeometrieConverter.xlsm", sheet_name_structure_loading, f"{Structure}_MASSES", MASSES)


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
    META_DB = META_FULL.loc[META_FULL["Identifier"] == selected_structure]

    DATA_DB = load_db_table(db_path, selected_structure)
    DATA_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA")

    #DATA_DB = load_db_table(db_path, selected_structure)
    MASSES_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_MASSES")

    def saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR):

        def valid_data(data):
            if pd.isna(data.values).any():
                return False, data
            try:
                return True, data.astype(float)
            except (ValueError, TypeError):
                return False, data

        succes, DATA_CURR = valid_data(DATA_CURR)

        if not succes:
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

            sucess = add_db_element(db_path, DATA_CURR, MASSES_CURR, META_CURR_NEW)

            if sucess:
                return True, META_CURR_NEW["identifier"]
            else:
                return False, None

        if meta_loaded_changed:
            if not ((META_CURR.values != None).all()):
                _ = ex.show_message_box("GeometrieConverter.xlsm", "Please fully populate the Current Meta table to modify the DB entry.")
                return False, _

            sucess = replace_db_element(db_path, DATA_CURR, MASSES_CURR, META_CURR, selected_structure)

            if sucess:
                return True, META_CURR["identifier"].values[0]
            else:
                return False, None

        if data_changed:
            sucess = write_db_element_data(db_path, selected_structure, DATA_CURR, MASSES_CURR)

            if sucess:
                return True, META_CURR["identifier"].values[0]
            else:
                return False, None

        _ = ex.show_message_box("GeometrieConverter.xlsm", f"No changes detected.")
        return False, _

    saved, structure_load_after = saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR)

    if saved:
        load_META(Structure, db_path)
        load_DATA(Structure, structure_load_after, db_path)


# do that

# %% Database element handling

def add_db_element(db_path, Structure_data, added_masses_data, Meta_infos):

    META = load_db_table(db_path, "META")

    if Identifier in META["Identifier"].values:
        _ = ex.show_message_box("GeometrieConverter.xlsm", "The provided Identifier of the new structure is already in the database, please provide a unique name")
        return False

    create_db_table(db_path, Identifier, Structure_data, if_exists='fail')
    create_db_table(db_path, f"{Identifier}__ADDED_MASSES", added_masses_data, if_exists='fail')

    META = pd.concat([META, Meta_infos], axis=1)
    create_db_table(db_path, "META", META, if_exists='replace')

    _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data saved in new Database entry {Identifier}")

    return True

def delete_db_element():

    return

def replace_db_element(db_path, Structure_data, added_masses_data, Meta_infos, replace_id):

    META = load_db_table(db_path, "META")
    new_id = Meta_infos["Identifier"].values[0]

    # replace data in meta
    META.loc[META["Identifier"] == replace_id, :] = Meta_infos.iloc[0].values

    if META["Identifier"].value_counts().get(replace_id, 0) > 1:
        _ = ex.show_message_box("GeometrieConverter.xlsm", f"{replace_id} is already taken in database, please choose a different name.")
        return False

    drop_db_table(db_path, replace_id)
    drop_db_table(db_path, f"{replace_id}__ADDED_MASSES")

    sucess = write_db_element_data(db_path, new_id, Structure_data, added_masses_data)

    if not sucess:
        return False

    create_db_table(db_path, "META", META, if_exists="replace")

    _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data in {replace_id} overwriten (now named {new_id})")

    return True


def write_db_element_data(db_path, change_id, Structure_data, added_masses_data):
    if pd.isna(Structure_data.values).any():
        _ = ex.show_message_box("GeometrieConverter.xlsm", f"Structure data contains invalid values, please correct")
        return False

    if pd.isna(added_masses_data.values).any():
        _ = ex.show_message_box("GeometrieConverter.xlsm", f"Added Masses data contains invalid values, please correct")
        return False

    create_db_table(db_path, change_id, Structure_data, if_exists="replace")
    create_db_table(db_path, change_id+"__ADDED_MASSES", added_masses_data, if_exists="replace")

    return True

def check_db_integrity():

    return



def save_MP_data(db_path, selected_structure):
    save_data("MP", db_path, selected_structure)

    return


save_MP_data(db_path, selected_structure)
