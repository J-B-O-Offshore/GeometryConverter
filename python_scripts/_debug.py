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


def load_db_table(db_path, Identifier, dtype=None):
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

    if Identifier not in table_names:
        raise ConciveError(f"Table '{Identifier}' does not exist in the database.")

    query = f'SELECT * FROM "{Identifier}"'
    df = pd.read_sql_query(query, conn)
    conn.close()

    if "index" in list(df.columns):
        df.drop(columns=['index'], inplace=True)

    if dtype is not None:
        df = df.astype(dtype)

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


def add_db_element(db_path, Structure_data, added_masses_data, Meta_infos):

    log.debug(Meta_infos.to_string())
    META = load_db_table(db_path, "META")

    Identifier = Meta_infos["Identifier"].values[0]
    if Identifier in META["Identifier"].values:
        _ = ex.show_message_box("GeometrieConverter.xlsm", "The provided Identifier of the new structure is already in the database, please provide a unique name.")
        return False

    create_db_table(db_path, Identifier, Structure_data, if_exists='fail')
    create_db_table(db_path, f"{Identifier}__ADDED_MASSES", added_masses_data, if_exists='fail')

    META = pd.concat([META, Meta_infos], axis=0)
    create_db_table(db_path, "META", META, if_exists='replace')

    _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data saved in new database entry {Identifier}")

    return True

def delete_db_element(db_path, Identifier):
    META = load_db_table(db_path, "META")

    META.drop(META[META["Identifier"] == Identifier].index, inplace=True)

    drop_db_table(db_path, Identifier)
    drop_db_table(db_path, Identifier+"__ADDED_MASSES")

    create_db_table(db_path, "META", META, if_exists='replace')
    return


def replace_db_element(db_path, Structure_data, added_masses_data, Meta_infos, replace_id):

    META = load_db_table(db_path, "META")
    new_id = Meta_infos["Identifier"].values[0]

    # replace data in meta
    META.loc[META["Identifier"] == replace_id, :] = Meta_infos.iloc[0].values
    # ex.show_message_box("GeometrieConverter.xlsm", "counts_replaceID"+str(META["Identifier"].value_counts().get(replace_id, 0)))
    # ex.show_message_box("GeometrieConverter.xlsm", "replace id"+str(replace_id))
    # ex.show_message_box("GeometrieConverter.xlsm", "META NEw Value"+str(Meta_infos.iloc[0].values[0]))
    # ex.show_message_box("GeometrieConverter.xlsm", "counts_NEw Value"+str(META["Identifier"].value_counts().get(Meta_infos.iloc[0].values, 0)))


    if META["Identifier"].value_counts().get(Meta_infos.iloc[0].values[0], 0) > 1:
        _ = ex.show_message_box("GeometrieConverter.xlsm", f"{replace_id} is already used in database, please choose a different name.")
        return False

    drop_db_table(db_path, replace_id)
    drop_db_table(db_path, f"{replace_id}__ADDED_MASSES")

    sucess = write_db_element_data(db_path, new_id, Structure_data, added_masses_data)

    if not sucess:
        return False

    create_db_table(db_path, "META", META, if_exists="replace")

    _ = ex.show_message_box("GeometrieConverter.xlsm", f"Data in {replace_id} overwritten (now named {new_id})")

    return True


def write_db_element_data(db_path, change_id, Structure_data, added_masses_data):
    if pd.isna(Structure_data.values).any():
        _ = ex.show_message_box("GeometrieConverter.xlsm", f"Structure data contains invalid values, please correct.")
        return False

    # if pd.isna(added_masses_data.values).any():
    #     _ = ex.show_message_box("GeometrieConverter.xlsm", f"Added Masses data contains invalid values, please correct")
    #     return False

    create_db_table(db_path, change_id, Structure_data, if_exists="replace")
    create_db_table(db_path, change_id+"__ADDED_MASSES", added_masses_data, if_exists="replace")

    return True

def check_db_integrity():

    return


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
    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    ex.set_dropdown_values(
        "GeometrieConverter.xlsm",
        sheet_name_structure_loading,
        f"Dropdown_{Structure}_Structures",
        list(META.loc[:, "Identifier"].values)
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
    META_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META", dtype=str)
    META_CURR_NEW = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_NEW", dtype=str)

    META_DB = META_FULL.loc[META_FULL["Identifier"] == selected_structure]

    DATA_DB = load_db_table(db_path, selected_structure, dtype=float)
    DATA_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA", dtype=float)

    #DATA_DB = load_db_table(db_path, selected_structure)
    MASSES_DB = load_db_table(db_path, selected_structure+"__ADDED_MASSES")
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

        data_changed = not (DATA_DB.equals(DATA_CURR)) or not(MASSES_DB.equals(MASSES_CURR))
        meta_loaded_changed = not (META_DB.values[0][0:-2] == META_CURR.values[0][0:-2]).all()
        meta_new_populated = (META_CURR_NEW.values[0][0:-2] != 'None').any()


        if meta_new_populated:
            if not ((META_CURR_NEW.values[0][0:-1] != None).all()):
                _ = ex.show_message_box("GeometrieConverter.xlsm",
                                        "Please fully populate the NEW Meta table to create a new DB entry or clear it of all data to overwrite the loaded structure.")
                return False, _

            sucess = add_db_element(db_path, DATA_CURR, MASSES_CURR, META_CURR_NEW)

            if sucess:
                return True, META_CURR_NEW["Identifier"]
            else:
                return False, None

        if meta_loaded_changed:
            if not ((META_CURR.values[0][0:-1] != None).all()):
                _ = ex.show_message_box("GeometrieConverter.xlsm", "Please fully populate the current Meta table to modify the DB entry.")
                return False, _

            sucess = replace_db_element(db_path, DATA_CURR, MASSES_CURR, META_CURR, selected_structure)

            if sucess:
                return True, META_CURR["Identifier"].values[0]
            else:
                return False, None

        if data_changed:
            sucess = write_db_element_data(db_path, selected_structure, DATA_CURR, MASSES_CURR)

            if sucess:
                return True, META_CURR["Identifier"].values[0]
            else:
                return False, None

        _ = ex.show_message_box("GeometrieConverter.xlsm", f"No changes detected.")
        return False, _

    saved, structure_load_after = saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR)

    if saved:

        load_META(Structure, db_path)

  #      load_DATA(Structure, structure_load_after, db_path)


def delete_data(Structure, db_path, selected_structure):
    answer = ex.show_message_box("GeometrieConverter.xlsm", f"Are you sure you want to delete the structure {selected_structure} from the database?", icon="vbYesNo",
                                 buttons="vbYesNo")
    if answer == "Yes":
        delete_db_element(db_path, selected_structure)

        load_META(Structure, db_path)

        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_NEW")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META_TRUE")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA_TRUE")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_MASSES_TRUE")
        ex.clear_excel_table_contents("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_MASSES")

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



selected_structure = "24A535_FEED_DP-A1_L0_G0_S0"
save_MP_data("C:/temp/_dev/_checks/GeometrieConverter/databases/MP.db",selected_structure)