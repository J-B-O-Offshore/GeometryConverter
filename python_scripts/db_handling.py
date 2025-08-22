import os.path
import xlwings as xw
import xlwings
import pandas as pd
import sqlite3
import excel as ex

import plot as ex_plt

class ConciveError(Exception):
    """
    Custom exception for errors related to database operations
    in the context of structure data handling.
    """

    pass


def drop_db_table(excel_filename, db_path, Identifier):
    """
    Drops (deletes) a table from an SQLite database.

    Parameters
    ----------
    db_path : str
        Path to the SQLite database file.
    Identifier : str
        Name of the table to drop.

    Returns
    -------
    bool
        True if table was successfully dropped, False otherwise.

    Notes
    -----
    - If the table does not exist, no error is raised.
    - Errors are shown via message box and do not raise exceptions.
    """
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        ex.show_message_box(
            excel_filename,
            f"Failed to connect to the database '{db_path}'.\nPython Error: {e}"
        )
        return False

    try:
        conn.execute(f'DROP TABLE IF EXISTS "{Identifier}"')
        conn.commit()
    except sqlite3.Error as e:
        ex.show_message_box(
            excel_filename,
            f"Failed to drop table '{Identifier}'.\nPython Error: {e}"
        )
        return False
    finally:
        conn.close()

    return True

def create_db_table(excel_filename, db_path, Identifier, df, if_exists='fail'):
    """
    Creates or appends to a table in an SQLite database from a pandas DataFrame.

    Parameters
    ----------
    db_path : str
        Path to the SQLite database file.
    Identifier : str
        Name of the table to create or append to.
    df : pd.DataFrame
        DataFrame containing the data to write.
    if_exists : str
        Behavior if the table already exists. One of: 'fail', 'replace', 'append'.

    Returns
    -------
    bool
        True if table creation/appending was successful, False otherwise.

    Notes
    -----
    - Uses pandas `to_sql()` under the hood.
    - Shows message boxes via `ex.show_message_box()` on failure.
    """
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        ex.show_message_box(
            excel_filename,
            f"Failed to connect to the database '{db_path}'.\nPython Error: {e}"
        )
        return False

    try:
        df.to_sql(Identifier, conn, if_exists=if_exists, index=False)
    except Exception as e:
        ex.show_message_box(
            excel_filename,
            f"Failed to create or write to table '{Identifier}'.\nPython Error: {e}"
        )
        return False
    finally:
        conn.close()

    return True

def load_db_table(excel_filename, db_path, Identifier, dtype=None):
    """
    Load a specific table from an SQLite database into a pandas DataFrame.

    Args:
        excel_filename (str): Path to the Excel file (used for message box context).
        db_path (str): Path to the SQLite database file.
        Identifier (str): Name of the table to load.
        dtype (dict): Optional dictionary specifying column data types.

    Returns:
        pd.DataFrame or None: The table content as a pandas DataFrame if successful, otherwise None.

    Notes:
        Shows Excel warning boxes via `ex.show_message_box()` on failure.
    """
    if not isinstance(Identifier, str):
        ex.show_message_box(
            excel_filename,
            f"Invalid Identifier.\nExpected string but got {type(Identifier)}."
        )
        return None

    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.Error as e:
        ex.show_message_box(
            excel_filename,
            f"Failed to connect to the database '{db_path}'.\nPython Error: {e}"
        )
        return None

    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        table_names = [table[0] for table in cursor.fetchall()]
    except sqlite3.Error as e:
        conn.close()
        ex.show_message_box(
            excel_filename,
            f"Failed to retrieve table names from '{db_path}'.\nPython Error: {e}"
        )
        return None

    if Identifier not in table_names:
        conn.close()
        ex.show_message_box(
            excel_filename,
            f"Table '{Identifier}' does not exist in the database.\nAvailable tables: {table_names}"
        )
        return None

    try:
        query = f'SELECT * FROM "{Identifier}"'
        df = pd.read_sql_query(query, conn)
    except Exception as e:
        conn.close()
        ex.show_message_box(
            excel_filename,
            f"Failed to load table '{Identifier}' from database.\nPython Error: {e}"
        )
        return None
    finally:
        conn.close()

    # Optional: Remove 'index' column if it exists
    if 'index' in df.columns:
        df = df.drop(columns=['index'])

    # Optional: Apply dtype conversion
    if dtype is not None:
        try:
            df = df.astype(dtype)
        except Exception as e:
            ex.show_message_box(
                excel_filename,
                f"Failed to apply data types {dtype} to table '{Identifier}'.\nPython Error: {e}"
            )
            return None

    return df


def add_db_element(excel_filename, db_path, Structure_data, added_masses_data, Meta_values):
    """
    Adds a new structure entry to the database using the provided data.

    Parameters
    ----------
    db_path : str
        Path to the SQLite database.
    Structure_data : pd.DataFrame
        Main structure data to be saved in a new table named after the Identifier.
    added_masses_data : pd.DataFrame
        Additional mass data to be saved in a new table named '{Identifier}__ADDED_MASSES'.
    Meta_values : list
        List of metadata values in the correct order. This must match the order of columns
        in the existing 'META' table in the database:
        ["Identifier", "Project ID", "Phase", "Structure ID", "Water Depth", "Height Reference", "comments"]

    Notes
    -----
    - The column names from the 'META' table are treated as the master schema.
    - The value for the 'Identifier' column is used to determine the new table name.
    - The function shows a warning and aborts if:
        - The number of metadata values is incorrect
        - The identifier already exists in the database
    """
    # Load existing META table and schema
    META = load_db_table(excel_filename, db_path, "META")
    if META is None:
        return

    meta_columns = META.columns.tolist()
    Meta_values = list(Meta_values)

    if len(Meta_values) != len(meta_columns):
        ex.show_message_box(
            excel_filename,
            f"The number of provided meta data values ({len(Meta_values)}) does not match to the expected number of columns ({len(meta_columns)})."
        )
        return False

    Identifier = Meta_values[0]
    if Identifier in META["Identifier"].values:
        ex.show_message_box(
            excel_filename,
            "The provided Identifier already exists in the database. Please provide a unique name."
        )
        return False

    # Convert list to DataFrame with correct columns
    Meta_infos = pd.DataFrame([Meta_values], columns=meta_columns)

    # Save structure data
    if not create_db_table(excel_filename, db_path, Identifier, Structure_data, if_exists='fail'):
        return False

    # Save added masses data
    if not create_db_table(excel_filename, db_path, f"{Identifier}__ADDED_MASSES", added_masses_data, if_exists='fail'):
        return False

    # Append new metadata row and save updated META table
    META = pd.concat([META, Meta_infos], ignore_index=True)
    if not create_db_table(excel_filename, db_path, "META", META, if_exists='replace'):
        return False

    ex.show_message_box(excel_filename, f"Data saved in new database entry '{Identifier}'.")
    return True


def delete_db_element(excel_filename, db_path, Identifier):
    """
    Deletes an existing structure and its associated tables from the database.

    Parameters
    ----------
    db_path : str
        Path to the SQLite database.
    Identifier : str
        Identifier of the structure to delete. This determines the main and mass table names.

    Returns
    -------
    bool
        True if deletion was successful, False otherwise.

    Notes
    -----
    - This function removes the entry from the 'META' table and deletes the corresponding tables:
        - '{Identifier}'
        - '{Identifier}__ADDED_MASSES'
    - If any step fails, an error message is shown and the operation is aborted.
    """
    META = load_db_table(excel_filename, db_path, "META")
    if META is None:
        return False

    if Identifier not in META["Identifier"].values:
        ex.show_message_box(
            excel_filename,
            f"No entry found for Identifier '{Identifier}' in the database."
        )
        return False

    # Remove the identifier row from META
    META = META[META["Identifier"] != Identifier]

    if not drop_db_table(excel_filename, db_path, Identifier):
        return False
    if not drop_db_table(excel_filename, db_path, f"{Identifier}__ADDED_MASSES"):
        return False
    if not create_db_table(excel_filename, db_path, "META", META, if_exists='replace'):
        return False

    ex.show_message_box(excel_filename, f"Deleted structure '{Identifier}' from the database.")
    return True


def replace_db_element(excel_filename, db_path, Structure_data, added_masses_data, Meta_infos, old_id):
    """
    Replaces an existing structure entry in the database with updated data and metadata.

    Parameters
    ----------
    db_path : str
        Path to the SQLite database.
    Structure_data : pd.DataFrame
        Updated structure data to be saved under the new identifier.
    added_masses_data : pd.DataFrame
        Updated additional mass data to be saved.
    Meta_infos : list
        Updated metadata values in the correct order, corresponding to columns in the 'META' table.
    old_id : str
        Identifier of the structure entry to be replaced.

    Returns
    -------
    bool
        True if replacement was successful, False otherwise.

    Notes
    -----
    - If the identifier is changed (`old_id` → `new_id`), the old entry and its tables will be dropped.
    - The function checks for duplicate identifiers before proceeding.
    """

    META = load_db_table(excel_filename, db_path, "META", dtype=str)
    if META is None:
        return

    Meta_infos = list(Meta_infos)
    new_id = Meta_infos[0]

    # Check for existing entry and valid update
    if old_id not in META["Identifier"].values:
        ex.show_message_box(excel_filename, f"No existing entry found for '{old_id}' in the database.")
        return False

    # Replace row in META
    META.loc[META["Identifier"] == old_id, :] = Meta_infos

    # Ensure no duplicate Identifiers after update
    if META["Identifier"].value_counts().get(new_id, 0) > 1:
        ex.show_message_box(excel_filename, f"The identifier '{new_id}' is already used more than once. Please choose a unique name.")
        return False

    # Save new data
    if not hardwrite_db_element_data(excel_filename, db_path, new_id, Structure_data, added_masses_data):
        return False

    # Drop old tables if identifier was changed
    if new_id != old_id:
        if not drop_db_table(excel_filename, db_path, old_id):
            return False
        if not drop_db_table(excel_filename, db_path, f"{old_id}__ADDED_MASSES"):
            return False

    # Save updated META table
    if not create_db_table(excel_filename, db_path, "META", META, if_exists="replace"):
        return False

    ex.show_message_box(excel_filename, f"Data for '{old_id}' successfully replaced with new entry '{new_id}'.")
    return True


def hardwrite_db_element_data(excel_filename, db_path, change_id, Structure_data, added_masses_data):
    """
    Replaces the structure and added masses tables in the database with new data.

    Parameters
    ----------
    db_path : str
        Path to the SQLite database.
    change_id : str
        Identifier for the structure; determines the table names.
    Structure_data : pd.DataFrame
        New structure data to overwrite the existing table '{change_id}'.
    added_masses_data : pd.DataFrame
        New added mass data to overwrite the existing table '{change_id}__ADDED_MASSES'.

    Returns
    -------
    bool
        True if replacement was successful, False otherwise.

    Notes
    -----
    - Both tables will be replaced (`if_exists='replace'`).
    - A warning is shown if invalid (NaN) values are detected in `Structure_data`.
    """
    if pd.isna(Structure_data.values).any():
        ex.show_message_box(excel_filename, "Structure data contains invalid (NaN) values. Please correct them before proceeding.")
        return False

    if not create_db_table(excel_filename, db_path, change_id, Structure_data, if_exists="replace"):
        return False

    if not create_db_table(excel_filename, db_path, f"{change_id}__ADDED_MASSES", added_masses_data, if_exists="replace"):
        return False

    return True


def check_db_integrity():
    return


def load_META(excel_filename, Structure, db_path):
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

    META = load_db_table(excel_filename, db_path, "META")
    if META is None:
        return
    ex.set_dropdown_values(
        excel_filename,
        sheet_name_structure_loading,
        f"Dropdown_{Structure}_Structures2",
        list(META.loc[:, "Identifier"].values)
    )
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_META_FULL", META)
    return


def load_DATA(excel_filename, Structure, Structure_name, db_path):
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

    if db_path is None or db_path == "":
        ex.show_message_box(excel_filename, "No database path provided.")
        return

    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(excel_filename, db_path, "META")
    if META is None:
        return

    DATA = load_db_table(excel_filename, db_path, Structure_name)

    if DATA is None:
        return
    MASSES = load_db_table(excel_filename, db_path, Structure_name + "__ADDED_MASSES")

    if MASSES is None:
        return

    META_relevant = META.loc[META["Identifier"] == Structure_name]

    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_META_TRUE", META_relevant)
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_META", META_relevant)
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_DATA_TRUE", DATA)
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_DATA", DATA)
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_MASSES_TRUE", MASSES)
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"{Structure}_MASSES", MASSES)

    ex.clear_excel_table_contents(excel_filename, sheet_name_structure_loading, f"{Structure}_META_NEW")
    ex.call_vba_dropdown_macro(excel_filename, sheet_name_structure_loading, f"Dropdown_{Structure}_Structures2", Structure_name)


def save_data(excel_filename, Structure, db_path, selected_structure):
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

    def saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR, MASSES_DB, MASSES_CURR):

        def valid_data(data):
            if pd.isna(data.values).any():
                return False, data
            try:
                return True, data.astype(float)
            except (ValueError, TypeError):
                return False, data

        succes, DATA_CURR = valid_data(DATA_CURR)

        if not succes:
            _ = ex.show_message_box(excel_filename,
                                    "Invalid data found in Structure data! Aborting.")
            return False, _

        # check, if data has changed
        data_changed = not (DATA_DB.equals(DATA_CURR)) or not (MASSES_DB.equals(MASSES_CURR))

        # check, if values are in META
        if META_CURR.empty:
            meta_loaded_changed = False
        else:
            meta_loaded_changed = (META_CURR.values[0][:-1] != '').any()

        # check, if values are in META_NEW
        if META_CURR_NEW.empty:
            meta_new_populated = False
        else:
            meta_new_populated = (META_CURR_NEW.values[0][:-1] != '').any()

        if meta_new_populated:
            if ((META_CURR_NEW.values[0][0:-1] == '').any()):
                _ = ex.show_message_box(excel_filename,
                                        "Please fully populate the NEW Meta table to create a new DB entry or clear it of all data to overwrite the loaded structure.")
                return False, _

            sucess = add_db_element(excel_filename, db_path, DATA_CURR, MASSES_CURR, META_CURR_NEW.values[0])

            if sucess:
                return True, META_CURR_NEW["Identifier"].values[0]
            else:
                return False, None

        if meta_loaded_changed:
            if not ((META_CURR.values[0][0:-1] != None).all()):
                _ = ex.show_message_box(excel_filename, "Please fully populate the current Meta table to modify the DB entry.")
                return False, _
            if selected_structure == "":
                _ = ex.show_message_box(excel_filename, "Now database loaded. Please enter data into \"NEW\" data field to create new database.")

            sucess = replace_db_element(excel_filename, db_path, DATA_CURR, MASSES_CURR, META_CURR.values[0], selected_structure)

            if sucess:
                return True, META_CURR["Identifier"].values[0]
            else:
                return False, None

        if data_changed:
            sucess = hardwrite_db_element_data(excel_filename, db_path, selected_structure, DATA_CURR, MASSES_CURR)

            if sucess:
                return True, META_CURR["Identifier"].values[0]
            else:
                return False, None

        _ = ex.show_message_box(excel_filename, f"No changes detected.")
        return False, _

    META_FULL = load_db_table(excel_filename, db_path, "META")
    if META_FULL is None:
        return
    META_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_META", dtype=str)
    META_CURR_NEW = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_META_NEW", dtype=str)
    DATA_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_DATA")
    MASSES_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_MASSES")

    # drop "" and nan rows
    META_CURR = META_CURR[~(META_CURR == "").all(axis=1)]
    META_CURR_NEW = META_CURR_NEW[~(META_CURR_NEW == "").all(axis=1)]

    DATA_CURR = DATA_CURR.dropna(how='all')
    MASSES_CURR = MASSES_CURR.dropna(how='all')

    META_DB = META_FULL.loc[META_FULL["Identifier"] == selected_structure]

    if selected_structure != "":
        DATA_DB = load_db_table(excel_filename, db_path, selected_structure)
        if DATA_DB is None:
            return
        MASSES_DB = load_db_table(excel_filename, db_path, selected_structure + "__ADDED_MASSES")
        if MASSES_DB is None:
            return
    else:
        DATA_DB = pd.DataFrame()
        MASSES_DB = pd.DataFrame()

    saved, structure_load_after = saving_logic(META_FULL, META_DB, META_CURR, META_CURR_NEW, DATA_DB, DATA_CURR, MASSES_DB, MASSES_CURR)

    if saved:
        load_META(excel_filename, Structure, db_path)

        load_DATA(excel_filename, Structure, structure_load_after, db_path)


def delete_data(excel_filename, Structure, db_path, selected_structure):
    answer = ex.show_message_box(excel_filename, f"Are you sure you want to delete the structure {selected_structure} from the database?", icon="vbYesNo",
                                 buttons="vbYesNo")
    if answer == "Yes":
        delete_db_element(excel_filename, db_path, selected_structure)

        load_META(excel_filename, Structure, db_path)

        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_META_NEW")
        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_META_TRUE")
        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_META")
        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_DATA_TRUE")
        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_DATA")
        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_MASSES_TRUE")
        ex.clear_excel_table_contents(excel_filename, "BuildYourStructure", f"{Structure}_MASSES")

    return


# %% MP
def load_MP_META(excel_caller, db_path):
    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the MP SQLite database.

    Returns:
        None
    """
    excel_filename = os.path.basename(excel_caller)
    load_META(excel_filename,"MP", db_path)


def load_MP_DATA(excel_caller, Structure_name, db_path):
    """
    Load metadata and structure-specific data from the MP database
    and write them to the Excel workbook.

    Args:
        Structure_name (str): Name of the structure (table) to load.
        db_path (str): Path to the MP SQLite database.

    Returns:
        None
    """
    excel_filename = os.path.basename(excel_caller)

    load_DATA(excel_filename, "MP", Structure_name, db_path)
    ex_plt.plot_MP(excel_caller)

def save_MP_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    save_data(excel_filename, "MP", db_path, selected_structure)

    return


def delete_MP_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    delete_data(excel_filename, "MP", db_path, selected_structure)

    return


def load_MP_from_MPTool(excel_caller, MP_path):
    excel_filename = os.path.basename(excel_caller)

    try:
        Section_col = ex.read_excel_range(MP_path, "Geometry", "C1:C1000")
        Section_col = Section_col.iloc[:, 0].dropna()
        row_MP = Section_col[Section_col == "Section"].index.values[1]

        Data = ex.read_excel_range(MP_path, "Geometry", f"C{row_MP + 3}:H1000", dtype=float)
        Data = Data.dropna(how="all")
        ex.write_df_to_table(excel_filename, "BuildYourStructure", "MP_DATA", Data)
    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading {MP_path} TP and MP. Please make sure, the path leads to a valid MP_tool xlsm file and has the TP data on the 'Geometry' sheet 3 rows under the second 'Section' header, empty rows allowed. Error thrown by Python: {e}.")
        return

    try:
        Parameter_col = ex.read_excel_range(MP_path, "Control", "E1:E1000", dtype=str, use_header=False)
        Parameter_col = Parameter_col.iloc[:, 0].dropna()

        row_RL = Parameter_col[Parameter_col.str.strip() == "Reference level"].index[0]
        row_ML = Parameter_col[Parameter_col.str.strip() == "Mudline"].index.values[0]

        Refercene_Level = ex.read_excel_range(MP_path, "Control", f"F{row_RL + 1}", dtype=str, use_header=False)
        Mudline = ex.read_excel_range(MP_path, "Control", f"F{row_ML + 1}", dtype=float, use_header=False)

        META_NEW = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_META_NEW", dtype=str)

        if len(META_NEW) == 0:
            META_NEW.iloc[0, :] = ""

        META_NEW.loc[0, "Height Reference"] = Refercene_Level
        META_NEW.loc[0, "Water Depth [m]"] = -Mudline

        META_NEW.loc[0, "Structure ID"] = os.path.basename(MP_path).replace(".xlsm", "")
        ex.write_df_to_table(excel_filename, "BuildYourStructure", "MP_META_NEW", META_NEW)

        # update porject name geneatation
        wb = xw.Book(excel_filename)
        wb.macro('UpdateIdentifierColumn')("BuildYourStructure", "MP_META")

    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading {MP_path} water level. Please make sure, the path leads to a valid MP_tool xlsm file and has the data on the 'Control' sheet in the range E:F, empty rows allowed. The Values are identified by the Keywords 'Reference level' and 'Mudline'. Error thrown by Python: {e}.")
        return

    return


def load_MPMasses_from_GeomConv(excel_caller, GeomConv_path):
    excel_filename = os.path.basename(excel_caller)

    try:
        Data_temp = ex.read_excel_range(GeomConv_path, "Additional entries", "K11:M1000", use_header=False)
        Data_temp = Data_temp.dropna(how="all")
        Data = pd.DataFrame({
            "Name": Data_temp.iloc[:, 2],
            "Top [m]": Data_temp.iloc[:, 0],
            "Bottom [m]": Data_temp.iloc[:, 0],
            "Mass [kg]": Data_temp.iloc[:, 1],
            "Diameter [m]": None,
            "Orientation [°]": None,
            "Distance Axis to Axis": None,
            "Gap between surfaces": None,
            "Surface roughness [m]": None,
        })[[
            "Name", "Top [m]", "Bottom [m]", "Mass [kg]", "Diameter [m]",
            "Orientation [°]", "Distance Axis to Axis", "Gap between surfaces", "Surface roughness [m]"
        ]]

        Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_DATA", dtype=float)
        Structure = Structure.dropna(how="all")

        if len(Structure) == 0:
            ex.show_message_box(excel_filename,
                                f"Please provide MP Data.")
            return

        heigt_range = (Structure.loc[0, "Top [m]"], Structure.loc[len(Structure) - 1, "Bottom [m]"])

        Data = Data.loc[(Data["Top [m]"] <= heigt_range[0]) & (Data["Top [m]"] >= heigt_range[1]), :]

        EXCEL_MASSES = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_MASSES", dtype=str)
        EXCEL_MASSES = EXCEL_MASSES[~(EXCEL_MASSES == "").all(axis=1)]

        if EXCEL_MASSES.shape[0] != 0:
            ALL_MASSES = pd.concat([EXCEL_MASSES, Data], axis=0)
        else:
            ALL_MASSES = Data

        ex.write_df_to_table(excel_filename, "BuildYourStructure", "MP_MASSES", ALL_MASSES)

    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading MP masses from {GeomConv_path}. Please make sure, the path leads to a valid GeometrieConverter xlsm file and has the masses data on the 'Zusätzliche Eingaben' sheet in columns K11:M1000. Error thrown by Python: {e}.")
        return
    return


# %% TP
def load_TP_META(excel_caller, db_path):

    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the TP SQLite database.

    Returns:
        None
    """
    excel_filename = os.path.basename(excel_caller)

    load_META(excel_filename, "TP", db_path)


def load_TP_DATA(excel_caller, Structure_name, db_path):

    """
    Load metadata and structure-specific data from the TP database
    and write them to the Excel workbook.

    Args:
        Structure_name (str): Name of the structure (table) to load.
        db_path (str): Path to the TP SQLite database.

    Returns:
        None
    """
    excel_filename = os.path.basename(excel_caller)

    load_DATA(excel_filename, "TP", Structure_name, db_path)
    ex_plt.plot_TP(excel_caller)


def save_TP_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    save_data(excel_filename, "TP", db_path, selected_structure)

    return


def delete_TP_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    delete_data(excel_filename,"TP", db_path, selected_structure)

    return


def load_TPMasses_from_GeomConv(excel_caller, GeomConv_path):
    excel_filename = os.path.basename(excel_caller)

    try:
        Data_temp = ex.read_excel_range(GeomConv_path, "Additional Entries", "K11:M1000", use_header=False)
        Data_temp = Data_temp.dropna(how="all")
        Data = pd.DataFrame({
            "Name": Data_temp.iloc[:, 2],
            "Top [m]": Data_temp.iloc[:, 0],
            "Bottom [m]": Data_temp.iloc[:, 0],
            "Mass [kg]": Data_temp.iloc[:, 1],
            "Diameter [m]": None,
            "Orientation [°]": None,
            "Distance Axis to Axis": None,
            "Gap between surfaces": None,
            "Surface roughness [m]": None,
        })[[
            "Name", "Top [m]", "Bottom [m]", "Mass [kg]", "Diameter [m]",
            "Orientation [°]", "Distance Axis to Axis", "Gap between surfaces", "Surface roughness [m]"
        ]]

        Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_DATA", dtype=float)
        Structure = Structure.dropna(how="all")

        if len(Structure) == 0:
            ex.show_message_box(excel_filename,
                                f"Please provide TP data.")
            return

        heigt_range = (Structure.loc[0, "Top [m]"], Structure.loc[len(Structure) - 1, "Bottom [m]"])

        Data = Data.loc[(Data["Top [m]"] <= heigt_range[0]) & (Data["Top [m]"] >= heigt_range[1]), :]

        EXCEL_MASSES = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_MASSES", dtype=str)
        EXCEL_MASSES = EXCEL_MASSES[~(EXCEL_MASSES == "").all(axis=1)]

        if EXCEL_MASSES.shape[0] != 0:
            ALL_MASSES = pd.concat([EXCEL_MASSES, Data], axis=0)
        else:
            ALL_MASSES = Data

        ex.write_df_to_table(excel_filename, "BuildYourStructure", "TP_MASSES", ALL_MASSES)

    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading TP masses from {GeomConv_path}. Please make sure, the path leads to a valid GeometrieConverter xlsm file and has the masses data on the 'Zusätzliche Eingaben' sheet in columns K11:M1000. Error thrown by Python: {e}.")
        return
    return


# %% TOWER
def load_TOWER_META(excel_caller, db_path):

    """
    Load the META table from the MP database and update the MP structures dropdown
    in the Excel workbook.

    Args:
        db_path (str): Path to the TOWER SQLite database.

    Returns:
        None
    """
    excel_filename = os.path.basename(excel_caller)

    load_META(excel_filename, "TOWER", db_path)


def load_TOWER_DATA(excel_caller, Structure_name, db_path):
    """
    Load metadata and structure-specific data from the TOWER database
    and write them to the Excel workbook.

    Args:
        Structure_name (str): Name of the structure (table) to load.
        db_path (str): Path to the TOWER SQLite database.

    Returns:
        None
    """
    excel_filename = os.path.basename(excel_caller)

    load_DATA(excel_filename, "TOWER", Structure_name, db_path)
    ex_plt.plot_TOWER(excel_caller)


def save_TOWER_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    save_data(excel_filename, "TOWER", db_path, selected_structure)

    return


def delete_TOWER_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    delete_data(excel_filename, "TOWER", db_path, selected_structure)

    return


def load_TP_from_MPTool(excel_caller, MP_path):
    excel_filename = os.path.basename(excel_caller)

    try:
        Section_col = ex.read_excel_range(MP_path, "Geometry", "C1:C1000")
        Section_col = Section_col.iloc[:, 0].dropna()
        row_TP = Section_col[Section_col == "Section"].index.values[0]
        row_MP = Section_col[Section_col == "Section"].index.values[1]

        Data = ex.read_excel_range(MP_path, "Geometry", f"C{row_TP + 3}:H{row_MP - 2}", dtype=float)
        Data = Data.dropna(how="all")
        ex.write_df_to_table(excel_filename, "BuildYourStructure", "TP_DATA", Data)
    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading {MP_path} TP and MP. Please make sure, the path leads to a valid MP_tool xlsm file and has the TP data on the 'Geometry' sheet 3 rows under the first 'Section' header, empty rows allowed. Error thrown by Python: {e}.")

    try:
        Parameter_col = ex.read_excel_range(MP_path, "Control", "E1:E1000", dtype=str, use_header=False)
        Parameter_col = Parameter_col.iloc[:, 0].dropna()

        row_RL = Parameter_col[Parameter_col.str.strip() == "Reference level"].index[0]
        row_ML = Parameter_col[Parameter_col.str.strip() == "Mudline"].index.values[0]

        Refercene_Level = ex.read_excel_range(MP_path, "Control", f"F{row_RL + 1}", dtype=str, use_header=False)
        Mudline = ex.read_excel_range(MP_path, "Control", f"F{row_ML + 1}", dtype=float, use_header=False)

        META_NEW = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_META_NEW", dtype=str)

        if len(META_NEW) == 0:
            META_NEW.iloc[0, :] = ""

        META_NEW.loc[0, "Height Reference"] = Refercene_Level
        META_NEW.loc[0, "Water Depth [m]"] = -Mudline

        META_NEW.loc[0, "Structure ID"] = os.path.basename(MP_path).replace(".xlsm", "")
        ex.write_df_to_table(excel_filename, "BuildYourStructure", "TP_META_NEW", META_NEW)

        # update project name generation
        wb = xw.Book(excel_filename)
        wb.macro('UpdateIdentifierColumn')("BuildYourStructure", "TP_META")

    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading {MP_path} water level data. Please make sure, the path leads to a valid MP_tool xlsm file and has the data on the 'Control' sheet in the range E:F, empty rows allowed. The Values are identified by the Keywords 'Reference level' and 'Mudline'. Error thrown by Python: {e}.")

    return


def load_TOWERMasses_from_GeomConv(excel_caller, GeomConv_path):
    excel_filename = os.path.basename(excel_caller)

    try:
        Data_temp = ex.read_excel_range(GeomConv_path, "Additional Entries", "K11:M1000", use_header=False)
        Data_temp = Data_temp.dropna(how="all")
        Data = pd.DataFrame({
            "Name": Data_temp.iloc[:, 2],
            "Top [m]": Data_temp.iloc[:, 0],
            "Bottom [m]": Data_temp.iloc[:, 0],
            "Mass [kg]": Data_temp.iloc[:, 1],
            "Diameter [m]": None,
            "Orientation [°]": None,
            "Distance Axis to Axis": None,
            "Gap between surfaces": None,
            "Surface roughness [m]": None,
        })[[
            "Name", "Top [m]", "Bottom [m]", "Mass [kg]", "Diameter [m]",
            "Orientation [°]", "Distance Axis to Axis", "Gap between surfaces", "Surface roughness [m]"
        ]]
        Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_DATA", dtype=float)

        Structure = Structure.dropna(how="all")

        if len(Structure) == 0:
            ex.show_message_box(excel_filename,
                                f"Please provide TOWER Data")
            return

        heigt_range = (Structure.loc[0, "Top [m]"], None)

        Data = Data.loc[Data["Top [m]"] >= heigt_range[0], :]

        EXCEL_MASSES = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_MASSES", dtype=str)
        EXCEL_MASSES = EXCEL_MASSES[~(EXCEL_MASSES == "").all(axis=1)]

        if EXCEL_MASSES.shape[0] != 0:
            ALL_MASSES = pd.concat([EXCEL_MASSES, Data], axis=0)
        else:
            ALL_MASSES = Data

        ex.write_df_to_table(excel_filename, "BuildYourStructure", "TOWER_MASSES", ALL_MASSES)

    except Exception as e:
        ex.show_message_box(excel_filename,
                            f"Error reading TOWER masses from {GeomConv_path}. Please make sure, the path leads to a valid GeometrieConverter xlsm file and has the masses data on the 'Zusätzliche Eingaben' sheet in columns K11:M1000. Error thrown by Python: {e}.")
        return
    return


# %% RNA
def load_RNA_DATA(excel_caller, db_path):
    excel_filename = os.path.basename(excel_caller)

    sheet_name_structure_loading = "BuildYourStructure"

    data = load_db_table(excel_filename, db_path, "data")
    if data is None:
        return

    ex.set_dropdown_values(
        excel_filename,
        sheet_name_structure_loading,
        f"Dropdown_RNA_Structures",
        list(data.loc[:, "Identifier"].values)
    )
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"RNA_DATA", data)
    ex.write_df_to_table(excel_filename, sheet_name_structure_loading, f"RNA_DATA_TRUE", data)
    ex.set_dropdown_values(excel_filename, sheet_name_structure_loading, "Dropdown_RNA_Structures", list(data.loc[:, "Identifier"].values))


def save_RNA_data(excel_caller, db_path, selected_structure):
    excel_filename = os.path.basename(excel_caller)

    DATA_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"RNA_DATA", dtype=str)

    create_db_table(excel_filename, db_path, "data", DATA_CURR, if_exists='replace')

    return

