import os.path
import pandas as pd
import excel as ex
import misc as mc
import numpy as np
from pandas.api.types import CategoricalDtype
import os

from ALaPy import periphery as pe

import plot as plt
# %% helpers
os.environ["MPLBACKEND"] = "Agg"  # kein GUI nötig

# optional zusätzlich:
try:
    import matplotlib

    matplotlib.use("Agg", force=True)
except Exception:
    pass

def check_values(df: pd.DataFrame, columns=None, mode='missing') -> list[str]:
    """
    Check for missing or present values in specified columns of a DataFrame.

    Parameters:
    - df (pd.DataFrame): The DataFrame to check.
    - columns (list, optional): List of column names to check. Defaults to all columns.
    - mode (str): 'missing' to check for NaN/None values, 'present' to check for any present values.

    Returns:
    - list[str]: A list of error message strings describing the found issues.
    """
    if columns is None:
        columns = df.columns
    else:
        missing_cols = [col for col in columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"The following columns are not in the DataFrame: {missing_cols}")

    error_messages = []

    for col in columns:
        if mode == 'missing':
            mask = df[col].isna()
            if mask.any():
                for idx in df.index[mask]:
                    error_messages.append(f"Missing value in column '{col}' at index {idx}")
        elif mode == 'present':
            mask = ~df[col].isna()
            if mask.any():
                for idx in df.index[mask]:
                    error_messages.append(f"Unexpected value in column '{col}' at index {idx}: {df.at[idx, col]!r}")
        else:
            raise ValueError("Invalid mode. Use 'missing' or 'present'.")

    return error_messages


def fill_dataframe_with_defaults(df: pd.DataFrame, default: pd.Series) -> pd.DataFrame:
    """
    Fill missing or empty values in all configuration columns with defaults.

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with configuration columns (one column per config).
    default : pandas.Series
        The default values to fall back to.

    Returns
    -------
    pandas.DataFrame
        The DataFrame with missing/empty values filled by defaults.
    """
    # Replace empty strings with NaN for consistent filling
    df = df.replace("", np.nan)

    # Fill NaNs column by column using defaults
    return df.apply(lambda col: col.fillna(default))


def str_to_bool(s):
    """Convert string ('True'/'False', 'WAHR'/'FALSCH') into bool (safe for bools too)."""
    if isinstance(s, bool):
        return s

    s = str(s).strip().lower()
    if s in ("true", "wahr"):
        return True
    elif s in ("false", "falsch"):
        return False
    else:
        raise ValueError(f"Invalid boolean value: {s}")


# %% checks:
def check_marine_growth(mg: pd.DataFrame, name: str = "MARINE_GROWTH") -> tuple[bool, str]:
    required_cols = [
        "Bottom [m]", "Top [m]", "Marine Growth [mm]",
        "Density  [kg/m^3]", "Surface Roughness [m]"
    ]

    missing_cols = [col for col in required_cols if col not in mg.columns]
    if missing_cols:
        return False, f"Missing required columns in {name}:\n" + "\n".join(missing_cols)

    missing_vals = [col for col in required_cols if mg[col].isnull().any()]
    if missing_vals:
        return False, f"Missing values in {name}:\n" + "\n".join(missing_vals)

    if (mg["Marine Growth [mm]"] < 0).any():
        return False, f"{name} contains negative marine growth thickness values."

    if (mg["Density  [kg/m^3]"] <= 0).any():
        return False, f"{name} contains non-positive density values."

    return True, ""


def check_appurtenances(apps: pd.DataFrame) -> tuple[bool, str]:
    """
    Validate an appurtenances DataFrame, including mutual exclusivity rules.

    Parameters
    ----------
    apps : pd.DataFrame
        Appurtenance data. Required columns:
        'Top [m]', 'Bottom [m]', 'Mass [kg]', 'Diameter [m]',
        'Orientation [°]', 'Surface roughness [m]', 'Name',
        'Distance Axis to Axis [m]', 'Gap between surfaces [m]'.

    Returns
    -------
    tuple[bool, str]
        (True, "") if valid, else (False, error_message).
    """
    required_cols = [
        "Top [m]", "Bottom [m]", "Mass [kg]",
        "Diameter [m]", "Orientation [°]", "Surface roughness [m]", "Name",
        "Distance Axis to Axis [m]", "Gap between surfaces [m]"
    ]

    # Check required columns
    missing_cols = [col for col in required_cols if col not in apps.columns]
    if missing_cols:
        return False, "Missing required columns:\n" + "\n".join(missing_cols)

    # Check for missing values (excluding mutually exclusive pair)
    check_cols = [c for c in required_cols if c not in ["Distance Axis to Axis [m]", "Gap between surfaces [m]"]]
    missing_info = []
    for col in check_cols:
        mask = apps[col].isnull()
        if mask.any():
            names = apps.loc[mask, "Name"].astype(str).tolist()
            missing_info.append(f"Column '{col}' has missing values for: {names}")

    if missing_info:
        return False, "Missing values found:\n" + "\n".join(missing_info)

    # Mutual exclusivity check
    err_list = []
    for idx, row in apps.iterrows():
        axis_to_axis = row["Distance Axis to Axis [m]"]
        gap_between = row["Gap between surfaces [m]"]
        name = row["Name"]

        if pd.isna(axis_to_axis) and pd.isna(gap_between):
            err_list.append(
                f"'{name}': Define either 'Distance Axis to Axis [m]' "
                f"or 'Gap between surfaces [m]'."
            )
        elif pd.notna(axis_to_axis) and pd.notna(gap_between):
            err_list.append(
                f"'{name}': Define only one of 'Distance Axis to Axis [m]' "
                f"or 'Gap between surfaces [m]', not both."
            )

    if err_list:
        return False, "Geometry specification issues:\n" + "\n".join(err_list)

    return True, ""


def check_added_masses(masses: pd.DataFrame, name: str = "ADDITIONAL_MASSES") -> tuple[bool, str]:
    """
    Validate an additional masses DataFrame.

    Parameters
    ----------
    masses : pd.DataFrame
        Table of additional point masses. Required columns:
        'Top [m]', 'Bottom [m]', 'Mass [kg]', 'Name'.
    name : str, optional
        Name of the dataset (used in error messages).

    Returns
    -------
    tuple[bool, str]
        (True, "") if valid, else (False, error_message).
    """
    required_cols = ["Top [m]", "Bottom [m]", "Mass [kg]", "Name"]

    # Check required columns
    missing_cols = [col for col in required_cols if col not in masses.columns]
    if missing_cols:
        return False, f"Missing required columns in {name}:\n" + "\n".join(missing_cols)

    # Check for missing values
    missing_vals = [col for col in required_cols if masses[col].isnull().any()]
    if missing_vals:
        return False, f"Missing values in {name}:\n" + "\n".join(missing_vals)

    # Optional sanity checks
    if (masses["Mass [kg]"] <= 0).any():
        return False, f"{name} contains non-positive mass values."

    if (masses["Top [m]"] < masses["Bottom [m]"]).any():
        return False, f"{name} contains rows where Top [m] is below Bottom [m]."

    return True, ""


# %% JBOOST
def export_JBOOST(excel_caller, jboost_path):
    success = fill_JBOOST_auto_excel(excel_caller)

    if not success:
        return

    excel_filename = os.path.basename(excel_caller)
    jboost_path = os.path.abspath(jboost_path)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    jboost_exe_path = os.path.join(os.path.dirname(script_dir), "JBOOST\\JBOOST.exe")

    GEOMETRY = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
    GEOMETRY = GEOMETRY.drop(columns=["Section", "local Section"])
    MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES", dropnan=True)
    MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
    PARAMETERS = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PARAMETER", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    PROJECT = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", dtype=str)
    RNA = ex.read_excel_table(excel_filename, "StructureOverview", "RNA", dropnan=True)

    if len(RNA) == 0:
        ex.show_message_box(excel_filename, "Please define RNA parameters. Aborting")
        return

    if PARAMETERS.loc[PARAMETERS["Parameter"] == "RNA Inertia", "Value"].values[0] == "fore-aft":
        RNA["Inertia"] = RNA["Inertia of RNA fore-aft @COG [kg m^2]"]
    elif PARAMETERS.loc[PARAMETERS["Parameter"] == "RNA Inertia", "Value"].values[0] == "side-side":
        RNA["Inertia"] = RNA["Inertia of RNA side-side @COG [kg m^2]"]
    else:
        ex.show_message_box(excel_filename, "Please define 'side-side' or 'fore-aft' for RNA Inertia. Aborting")
        return

    # check Geometry
    sucess_GEOMETRY = mc.sanity_check_structure(excel_filename, GEOMETRY)
    if not sucess_GEOMETRY:
        ex.show_message_box(excel_filename, "Geometry is messed up. Aborting.")
        return

    Model_name = PARAMETERS.loc[PARAMETERS["Parameter"] == "ModelName", "Value"].values[0]

    # proj file
    PROJECT = PROJECT.set_index("Project Settings")
    default = PROJECT.loc[:, "default"]
    proj_configs = PROJECT.iloc[:, 3:]

    # --- fill defaults ---
    proj_configs = fill_dataframe_with_defaults(proj_configs, default)
    # -------------------------------------

    # iterate through configs
    for config_name, config_data in proj_configs.items():
        config_struct = {row: data for row, data in config_data.items()}
        config_struct.pop("runFEModul", None)
        config_struct.pop("runFrequencyModul", None)

        runFEModul = str_to_bool(config_data["runFEModul"])
        runFrequencyModul = str_to_bool(config_data["runFrequencyModul"])

        proj_text = pe.create_JBOOST_proj(
            config_struct,
            MARINE_GROWTH,
            modelname=Model_name,
            runFEModul=runFEModul,
            runFrequencyModul=runFrequencyModul,
            runHindcastValidation=False,
            wavefile="wave.lua",
            windfile="wind.lua",
            write_JBOOST_graph=True
        )

        struct_text = pe.create_JBOOST_struct(
            GEOMETRY,
            RNA,
            (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Value"].values[0],
             PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Unit"].values[0]),
            (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Value"].values[0],
             PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Unit"].values[0]),
            MASSES=MASSES,
            MARINE_GROWTH=MARINE_GROWTH,
            defl_TP=(PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Value"].values[0],
                     PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Unit"].values[0]),
            ModelName=Model_name,
            EModul=PARAMETERS.loc[PARAMETERS["Parameter"] == "EModul", "Value"].values[0],
            fyk="355",
            poisson="0.3",
            dens=PARAMETERS.loc[PARAMETERS["Parameter"] == "Steel Density", "Value"].values[0],
            addMass=0,
            member_id=1,
            create_node_tolerance=PARAMETERS.loc[
                PARAMETERS["Parameter"] == "Dimensional tolerance for node generating [m]", "Value"].values[0],
            seabed_level=STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"].values[0],
            waterlevel=config_struct["water_level"]
        )

        path_config = os.path.join(jboost_path, config_name)
        os.makedirs(path_config, exist_ok=True)

        path_proj = os.path.join(path_config, "proj.lua")
        path_struct = os.path.join(path_config, Model_name + ".lua")

        with open(path_proj, 'w') as file:
            file.write(proj_text)
        with open(path_struct, 'w') as file:
            file.write(struct_text)

    ex.show_message_box(
        excel_filename,
        f"JBOOST Structure {PARAMETERS.loc[PARAMETERS['Parameter'] == 'ModelName', 'Value'].values[0]} "
        f"saved successfully at {jboost_path}"
    )

    return


def fill_JBOOST_auto_excel(excel_caller):
    """
    Fills missing or automatically assigned values in the `JBOOST_PROJECT` Excel table
    based on defaults and metadata from the `StructureOverview` sheet.

    Reads the `JBOOST_PROJECT` table from the "ExportStructure" sheet
    and the `STRUCTURE_META` table from the "StructureOverview" sheet. Replaces
    missing or 'auto' values with either:
      - defaults from the `default` column, or
      - values from `STRUCTURE_META` (e.g., water level, seabed level, hub height).

    If a required 'auto' value cannot be resolved from `STRUCTURE_META`, a message box
    is shown and the function aborts.

    Parameters
    ----------
    excel_caller : str
        Path or filename of the Excel file containing the required tables.
        Only the basename is used internally.

    Returns
    -------
    bool
        True if all values were successfully filled; False if the process was aborted.
    """

    excel_filename = os.path.basename(excel_caller)

    # Load tables
    PROJECT = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", dtype=str)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    hubheight = ex.read_named_range(excel_caller, "HubHeight", sheet_name="StructureOverview", dtype=float, use_header=False)

    # Use 'Project Settings' as the index
    PROJECT.index = PROJECT["Project Settings"]

    default_values = PROJECT["default"]
    proj_configs = PROJECT.iloc[:, 4:]

    def resolve_auto_value(parameter, config_key, description, config_name):
        """Helper to resolve 'auto' values from STRUCTURE_META."""
        value = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == parameter, "Value"].values

        if len(value) > 0:
            if isinstance(value[0], (int, float)):
                PROJECT.at[config_key, config_name] = value[0]
                return True
            else:
                # Wrong type in Excel
                ex.show_message_box(
                    excel_filename,
                    f"Invalid type for {description} in StructureOverview.\n"
                    f"You set {config_key} in {config_name} to 'auto', but the corresponding value "
                    f"is not a number (found: {value[0]!r}). Aborting."
                )
                return False

        # Missing entry in Excel
        ex.show_message_box(
            excel_filename,
            f"Please set {description} in StructureOverview, as you set {config_key} "
            f"in {config_name} to 'auto'. Aborting."
        )
        return False

    # Iterate over project configurations
    for config_name, config_data in proj_configs.items():
        # Fill empty cells with defaults
        config_data = config_data.replace("", np.nan).fillna(default_values)

        # Fill specific 'auto' values
        if config_data.at["water_level"] == "auto":
            if not resolve_auto_value("Water level", "water_level", "a water level", config_name):
                return False
        if config_data.at["seabed_level"] == "auto":
            if not resolve_auto_value("Seabed level", "seabed_level", "a seabed level", config_name):
                return False
        if config_data.at["h_hub"] == "auto":
            PROJECT.at["h_hub", config_name] = hubheight
        if config_data.at["h_refwindspeed"] == "auto":
            PROJECT.at["h_refwindspeed", config_name] = PROJECT.at["h_hub", config_name]

    # Restore 'Project Settings' column
    PROJECT["Project Settings"] = PROJECT.index

    # Write back updated table
    ex.write_df_to_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", PROJECT)

    return True


def run_JBOOST_excel(excel_caller, export_path=""):
    success = fill_JBOOST_auto_excel(excel_caller)

    if success:
        excel_filename = os.path.basename(excel_caller)

        if len(export_path) > 0:
            export_path = os.path.abspath(export_path)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        jboost_path = os.path.join(os.path.dirname(script_dir), "JBOOST\\JBOOST.exe")

        GEOMETRY = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
        GEOMETRY = GEOMETRY.drop(columns=["Section", "local Section"])
        MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES", dropnan=True)
        MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
        PARAMETERS = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PARAMETER", dropnan=True)
        STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
        PROJECT = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", dtype=str)
        RNA = ex.read_excel_table(excel_filename, "StructureOverview", "RNA", dropnan=True)

        if len(RNA) == 0:
            ex.show_message_box(excel_filename, "Please define RNA parameters. Aborting")
            return

        if PARAMETERS.loc[PARAMETERS["Parameter"] == "RNA Inertia", "Value"].values[0] == "fore-aft":
            RNA["Inertia"] = RNA["Inertia of RNA fore-aft @COG [kg m^2]"]
        elif PARAMETERS.loc[PARAMETERS["Parameter"] == "RNA Inertia", "Value"].values[0] == "side-side":
            RNA["Inertia"] = RNA["Inertia of RNA side-side @COG [kg m^2]"]
        else:
            ex.show_message_box(excel_filename, "Please define 'side-side' or 'fore-aft' for RNA Inertia. Aborting")
            return

        # check Geometry
        sucess_GEOMETRY = mc.sanity_check_structure(excel_filename, GEOMETRY)
        if not sucess_GEOMETRY:
            ex.show_message_box(excel_filename, "Geometry is messed up. Aborting.")
            return

        Model_name = PARAMETERS.loc[PARAMETERS["Parameter"] == "ModelName", "Value"].values[0]

        # proj file
        PROJECT = PROJECT.set_index("Project Settings")
        default = PROJECT.loc[:, "default"]
        proj_configs = PROJECT.iloc[:, 3:]

        # --- fill defaults ---
        proj_configs = fill_dataframe_with_defaults(proj_configs, default)
        # -------------------------------------

        # iterate through configs
        Modeshapes = {}
        waterlevels = {}
        for config_name, config_data in proj_configs.items():
            config_struct = {row: data for row, data in config_data.items()}
            config_struct.pop("runFEModul", None)
            config_struct.pop("runFrequencyModul", None)

            proj_text = pe.create_JBOOST_proj(
                config_struct,
                MARINE_GROWTH,
                modelname=Model_name,
                write_JBOOST_graph=True
            )

            struct_text = pe.create_JBOOST_struct(
                GEOMETRY,
                RNA,
                (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Value"].values[0],
                 PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Unit"].values[0]),
                (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Value"].values[0],
                 PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Unit"].values[0]),
                MASSES=MASSES,
                MARINE_GROWTH=MARINE_GROWTH,
                defl_TP=(PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Value"].values[0],
                         PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Unit"].values[0]),
                ModelName=Model_name,
                EModul=PARAMETERS.loc[PARAMETERS["Parameter"] == "EModul", "Value"].values[0],
                fyk="355",
                poisson="0.3",
                dens=PARAMETERS.loc[PARAMETERS["Parameter"] == "Steel Density", "Value"].values[0],
                addMass=0,
                member_id=1,
                create_node_tolerance=PARAMETERS.loc[
                    PARAMETERS["Parameter"] == "Dimensional tolerance for node generating [m]", "Value"].values[0],
                seabed_level=STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"].values[0],
                waterlevel=config_struct["water_level"]
            )

            JBOOST_OUT = pe.run_JBOOST(jboost_path, proj_text, struct_text, set_calculation={"FEModul": True, "FreqDomain": True, "HindValid": False})

            sheet_names = []
            sheets = []

            if len(export_path) > 0:

                path_config = os.path.join(export_path, config_name)
                path_struct = os.path.join(path_config, Model_name + ".lua")
                path_proj = os.path.join(path_config, "proj.lua")
                path_out = os.path.join(path_config, "JBOOST_OUT.xlsx")

                os.makedirs(path_config, exist_ok=True)

                with open(path_proj, 'w') as file:
                    file.write(proj_text)
                with open(path_struct, 'w') as file:
                    file.write(struct_text)

                for key, value in JBOOST_OUT.items():
                    if value is not None:
                        sheet_names.append(key)
                        sheets.append(value)
                        pe.save_df_list_to_excel(path_out, sheets, sheet_names=sheet_names)

            Modeshapes[config_name] = JBOOST_OUT["Mode_shapes"]
            waterlevels[config_name] = config_struct["water_level"]

        reversed_dfs = {k: df.iloc[::-1].reset_index(drop=True) for k, df in Modeshapes.items()}

        FIG = plt.plot_modeshapes(reversed_dfs, order=(1, 2, 3), waterlevels=waterlevels)

        ex.insert_plot(FIG, excel_filename, "ExportStructure", f"FIG_JBOOST_MODESHAPES")

    return


def load_JBOOST_soil_file(excel_caller, path):
    excel_filename = os.path.basename(excel_caller)

    try:
        _, sparse,_ = pe.read_soil_stiffness_matrix_csv(path)
        sparse = sparse.T

        # Set default value for all columns
        sparse.loc["Use for JBOOST config? (Y/N)", :] = "N"
        sparse.loc["Short name", :] = sparse.columns

        # Boolean masks for columns
        reloading_init_col = sparse.columns.str.contains("reloading") & sparse.columns.str.contains("init")
        reloading_loaded_col = sparse.columns.str.contains("reloading") & sparse.columns.str.contains("loaded")
        static_red_init_col = (
                sparse.columns.str.contains("static")
                & sparse.columns.str.contains("init")
                & sparse.columns.str.contains("red")
        )
        static_red_loaded_col = (
                sparse.columns.str.contains("static")
                & sparse.columns.str.contains("loaded")
                & sparse.columns.str.contains("red")
        )

        # Apply changes if there are matches
        if reloading_init_col.any():
            sparse.loc["Use for JBOOST config? (Y/N)", reloading_init_col] = "Y"
            sparse.loc["Short name", reloading_init_col] = "reloading_init"
        if reloading_loaded_col.any():
            sparse.loc["Use for JBOOST config? (Y/N)", reloading_loaded_col] = "Y"
            sparse.loc["Short name", reloading_loaded_col] = "reloading_loaded"
        if static_red_init_col.any():
            sparse.loc["Use for JBOOST config? (Y/N)", static_red_init_col] = "Y"
            sparse.loc["Short name", static_red_init_col] = "static_red_init"
        if static_red_loaded_col.any():
            sparse.loc["Use for JBOOST config? (Y/N)", static_red_loaded_col] = "Y"
            sparse.loc["Short name", static_red_loaded_col] = "static_red_loaded"

        sparse.insert(0, "Stiffness", sparse.index)

        ex.clear_excel_table_contents(excel_filename, "ExportStructure", "JBOOST_soil_stiffness")
        ex.write_df_to_table_flexible(excel_filename, "ExportStructure", "JBOOST_soil_stiffness", sparse)

    except:
        ex.show_message_box(excel_filename, f"PY data file could not be read, make sure it is the right format and it is reachable.")
        ex.clear_excel_table_contents(excel_filename, "ExportStructure", "JBOOST_soil_stiffness")
    return


def create_JBOOST_soil_configs(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    PROJECT = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", dtype=str)
    PROJECT = PROJECT.iloc[:, 0:4]

    JBOOST_soil_stiffness = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_soil_stiffness", dtype=str, dropnan=True)
    JBOOST_soil_stiffness.set_index("Stiffness", inplace=True)

    mask = JBOOST_soil_stiffness.loc["Use for JBOOST config? (Y/N)"] == "Y"
    JBOOST_soil_stiffness = JBOOST_soil_stiffness.loc[:, mask]

    if len(JBOOST_soil_stiffness) == 0:
        ex.show_message_box(excel_filename, "Please fill Soil Stiffness table or toggle some configs, aborting.")
        return

    for col_name, values in JBOOST_soil_stiffness.items():
        Short_name = values["Short name"]
        PROJECT.loc[:, Short_name] = ""
        PROJECT.loc[PROJECT["Project Settings"] == "found_stiff_trans", Short_name] = values["found_stiff_trans [N/m]"]
        PROJECT.loc[PROJECT["Project Settings"] == "found_stiff_rotat", Short_name] = values["found_stiff_rotat [Nm/rad]"]
        PROJECT.loc[PROJECT["Project Settings"] == "found_stiff_coupl", Short_name] = values["found_stiff_coupl [Nm/m]"]

    ex.clear_excel_table_contents(excel_filename, "ExportStructure", "JBOOST_PROJECT")
    ex.write_df_to_table_flexible(excel_filename, "ExportStructure", "JBOOST_PROJECT", PROJECT)


# %% WLGEN

def export_WLGen(excel_caller, WLGen_path):
    excel_filename = os.path.basename(excel_caller)

    APPURTANCES_MASSES = ex.read_excel_table(excel_filename, "ExportStructure", "APPURTANCES", dropnan=True)
    STRUCTURE = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
    STRUCTURE = STRUCTURE.drop(columns=["Section", "local Section"])
    MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    SKIRT = ex.read_excel_table(excel_filename, "StructureOverview", "SKIRT", dropnan=True)

    APPURTANCES = APPURTANCES_MASSES.loc[APPURTANCES_MASSES.iloc[:, 0] == "WL", :]
    ADDITIONAL_MASSES = APPURTANCES_MASSES.loc[APPURTANCES_MASSES.iloc[:, 0] == "AM", :]

    # check APP

    # cut of below waterline
    z_wl = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"].values[0]
    STRUCTURE = mc.add_element(STRUCTURE, z_new=z_wl)
    STRUCTURE = STRUCTURE[STRUCTURE["Bottom [m]"] >= z_wl]

    # only take MP and TP
    MP = STRUCTURE.loc[STRUCTURE.loc[:, "Affiliation"] == "MP", :]
    TP = STRUCTURE.loc[STRUCTURE.loc[:, "Affiliation"] == "TP", :]

    if len(SKIRT) == 0:
        SKIRT = None

    # check
    ok, err = check_added_masses(ADDITIONAL_MASSES, "ADDITIONAL_MASSES")
    if not ok:
        ex.show_message_box(excel_filename, f"WLGen Structure creation failed, problem with Added masses: {err}")
        return

    ok, err = check_appurtenances(APPURTANCES)
    if not ok:
        ex.show_message_box(excel_filename, f"WLGen Structure creation failed, problem with Appurtances definition: {err}")
        return

    ok, err = check_marine_growth(MARINE_GROWTH, "MARINE_GROWTH")
    if not ok:
        ex.show_message_box(excel_filename, f"WLGen Structure creation failed, problem with Marine Growth: {err}")
        return

    # run WGEN creation
    try:
        text = pe.create_WLGen_file(APPURTANCES, ADDITIONAL_MASSES, MP, TP, MARINE_GROWTH, skirt=SKIRT)
    except Exception as e:
        ex.show_message_box(excel_filename, f"WLGen Structure could not be created: {e}")
        return

    model_name = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Model Name", "Value"].values[0]
    if model_name is None:
        model_name = "input_WLGen.lua"
        ex.show_message_box(excel_filename, f"No model name defined in Structure Overview. File named {model_name}.")
    else:
        model_name = model_name + ".lua"

    WLGen_path = os.path.abspath(os.path.join(WLGen_path, model_name))

    with open(WLGen_path, 'w') as file:
        file.write(text)
    ex.show_message_box(excel_filename, f"WLGen Structure created successfully and saved at {WLGen_path}.")
    return


def fill_WLGenMasses(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES")
    MASSES = MASSES.loc[(MASSES["Affiliation"] == "TP") | (MASSES["Affiliation"] == "MP")]

    def categorize_row(row):
        # Check mandatory fields
        if pd.isna(row['Top [m]']) or pd.isna(row['Mass [kg]']):
            return 'INVALID'

        # Check if it qualifies as an APPURTENANCE
        has_bottom = not pd.isna(row['Bottom [m]'])
        has_diameter = not pd.isna(row['Diameter [m]'])
        has_orientation = not pd.isna(row['Orientation [°]'])
        has_roughness = not pd.isna(row['Surface roughness [m]'])

        has_axis_to_axis = not pd.isna(row['Distance Axis to Axis [m]'])
        has_gap = not pd.isna(row['Gap between surfaces [m]'])
        # xor_axis_gap = has_axis_to_axis != has_gap  # exclusive OR

        if all([has_bottom, has_diameter, has_orientation, has_roughness]) and (has_axis_to_axis or has_gap):
            return 'WL'
        else:
            return 'AM'

    cols = ["Use For"] + list(MASSES.columns)
    MASSES_WL = pd.DataFrame(columns=cols)

    for idx, row in MASSES.iterrows():
        kind = categorize_row(row)

        row["Use For"] = kind

        row_df = row.to_frame().T  # Convert Series to 1-row DataFrame
        row_aligned = row_df[MASSES_WL.columns.intersection(row_df.columns)]

        MASSES_WL = pd.concat([MASSES_WL, row_aligned], ignore_index=True)

    # sort df
    # Define custom order
    cat_order = CategoricalDtype(categories=["WL", "AM", "INVALID"], ordered=True)

    # Convert the column to categorical
    MASSES_WL['Use For'] = MASSES_WL['Use For'].astype(cat_order)

    # Sort the DataFrame
    MASSES_WL = MASSES_WL.sort_values('Use For')

    ex.write_df_to_table(excel_filename, "ExportStructure", "APPURTANCES", MASSES_WL)


# %% BLADED

def fill_Bladed_table(excel_caller, incluce_py_nodes=False, selected_loadcase=None, py_path=None):
    excel_filename = os.path.basename(excel_caller)
    incluce_py_nodes = str_to_bool(incluce_py_nodes)
    # Read inputs
    Bladed_Settings = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Settings", dropnan=True)
    Bladed_Material = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Material", dropnan=True)
    GEOMETRY = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
    MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
    MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")

    APPURTANCES = MASSES[MASSES["Top [m]"] != MASSES["Bottom [m]"]]
    ADDITIONAL_MASSES = MASSES[MASSES["Top [m]"] == MASSES["Bottom [m]"]]
    soil_density = Bladed_Settings.loc[
        Bladed_Settings["Parameter"] == "Soil density", "Value"
    ].values[0]


    # check
    ok, err = check_added_masses(ADDITIONAL_MASSES, "ADDITIONAL_MASSES")
    if not ok:
        ex.show_message_box(excel_filename, f"Bladed Structure creation failed, problem with Added masses: {err}")
        return

    ok, err = check_appurtenances(APPURTANCES)
    if not ok:
        ex.show_message_box(excel_filename, f"Bladed Structure creation failed, problem with Appurtances definition: {err}")
        return

    ok, err = check_marine_growth(MARINE_GROWTH, "MARINE_GROWTH")
    if not ok:
        ex.show_message_box(excel_filename, f"Bladed Structure creation failed, problem with Marine Growth: {err}")
        return

    if incluce_py_nodes:
        PY_data = pe.read_geo_py_curves(py_path)
        if selected_loadcase not in PY_data:
            raise ValueError(f"Selected loadcase '{selected_loadcase}' not found in {py_path}.")
        PY_loadcase = PY_data[selected_loadcase]

        PY_loadcase_spring_heights = pd.Series(
            np.unique(PY_loadcase["z [m]"]),
            index=np.unique(PY_loadcase["Spring [-]"])
        )
        cut_embedded = False

    else:
        cut_embedded = True
        PY_loadcase_spring_heights = None


    Bladed_Elements, Bladed_Nodes = pe.build_Bladed_dataframes(
        Bladed_Settings, Bladed_Material, GEOMETRY, MARINE_GROWTH, MASSES, STRUCTURE_META,
        cut_embedded=cut_embedded, PY_springs=PY_loadcase_spring_heights, soil_density=soil_density
    )

    # Write outputs
    ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Elements", Bladed_Elements)
    ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Nodes", Bladed_Nodes)


def fill_bladed_py_dropdown(excel_caller, py_path):
    py_path = os.path.abspath(py_path)
    excel_filename = os.path.basename(excel_caller)


    try:
        PY_data = pe.read_geo_py_curves(py_path)
    except ValueError as err:
        ex.show_message_box(excel_filename, "PY data file could not be read, make sure it is the right format.")
        ex.set_dropdown_values(excel_filename, "ExportStructure", "Dropdown_Bladed_py_loadcase", [""])
        return

    ex.set_dropdown_values(excel_filename, "ExportStructure", "Dropdown_Bladed_py_loadcase", list(PY_data.keys()))

    return


def plot_bladed_py(excel_caller, py_path, selected_loadcase):
    py_path = os.path.abspath(py_path)
    excel_filename = os.path.basename(excel_caller)

    Bladed_Settings = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Settings", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")

    max_lines = Bladed_Settings.loc[Bladed_Settings["Parameter"] == "py lines per axis", "Value"].values[0]

    basename = STRUCTURE_META.loc[ STRUCTURE_META["Parameter"] == "Model Name", "Value"].values[0]
    if basename is None:
        basename = "Bladed_PJ_file"

    Bladed_Settings.loc[Bladed_Settings["Parameter"] == "PJ file name", "Value"] = basename + f"_{selected_loadcase}"

    try:
        PY_data = pe.read_geo_py_curves(py_path)
        PY_loadcase = PY_data[selected_loadcase]

        FIG = plt.plot_py_curves(PY_loadcase, loadcase=selected_loadcase, max_lines=int(max_lines))

        ex.insert_plot(FIG, excel_filename, "ExportStructure", f"FIG_PY_CURVES", replace=True)

    except ValueError as err:
        ex.show_message_box(excel_filename, f"PY data file could not be read or {selected_loadcase} not part of the file, make sure it is the right format and it is reachable.")
        ex.set_dropdown_values(excel_filename, "ExportStructure", "Dropdown_Bladed_py_loadcase", [""])
        return

    ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Settings", Bladed_Settings)


    return
def update_bladed_name(excel_caller, selected_loadcase):
    excel_filename = os.path.basename(excel_caller)
    Bladed_Settings = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Settings", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    basename = STRUCTURE_META.loc[ STRUCTURE_META["Parameter"] == "Model Name", "Value"].values[0]
    if basename is None:
        basename = "Bladed_PJ_file"
    Bladed_Settings.loc[Bladed_Settings["Parameter"] == "PJ file name", "Value"] = basename + f"_{selected_loadcase}"

    ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Settings", Bladed_Settings)

    return

def apply_bladed_py_curves(excel_caller, py_path, Bladed_pj_path, selected_loadcase, insert_mode=False, fig_path=None, update_tables=True):
    """
    Apply p–y curves to a Bladed project by generating and inserting corresponding PJ files,
    figures, and updated structural data.

    This function reads configuration and structural data from Excel, loads the p–y curve data
    from a Python-formatted file, generates Bladed-compatible PJ input text, and optionally
    inserts the results into an existing PJ file. It also creates control figures for
    interpolation verification and updates relevant Excel tables.

    Parameters
    ----------
    excel_caller : str
        Path to the Excel file that calls this function.
    py_path : str
        Path to the `.py` file containing the p–y curve data.
    Bladed_pj_path : str
        Directory path to the Bladed project where the PJ file will be created or updated.
    selected_loadcase : str
        The load case name to extract from the p–y curve data.
    insert_mode : bool, optional
        If True, inserts new data blocks into an existing PJ file instead of overwriting it.
        Default is False.
    fig_path : str, optional
        Directory path where interpolation control figures will be saved. If None, defaults
        to the Bladed project folder.

    Returns
    -------
    None
        Writes PJ file(s), figures, and Excel tables as side effects.

    Raises
    ------
    FileNotFoundError
        If any of the required files or directories cannot be found.
    ValueError
        If the p–y curve data or required Excel fields are missing or malformed.
        :param update_tables:
    """

    insert_mode = str_to_bool(insert_mode)
    update_tables = str_to_bool(update_tables)

    # --- Validate and normalize paths ---
    py_path = os.path.abspath(py_path)
    Bladed_pj_path = os.path.abspath(Bladed_pj_path)
    if fig_path is None:
        fig_path = os.path.abspath(os.path.dirname(Bladed_pj_path))
    else:
        fig_path = os.path.abspath(fig_path)

    excel_filename = os.path.basename(excel_caller)

    if not os.path.exists(py_path):
        raise FileNotFoundError(f"p–y curve file not found: {py_path}")
    if not os.path.exists(Bladed_pj_path):
        raise FileNotFoundError(f"Bladed project path not found: {Bladed_pj_path}")

    # --- Read Excel configuration and data ---
    try:
        Bladed_Settings = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Settings", dropnan=True)
        Bladed_Material = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Material", dropnan=True)
        GEOMETRY = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
        MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
        MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES", dropnan=True)
        STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    except Exception as err:
        raise ValueError(f"Failed to read one or more required Excel tables: {err}")

    # --- Extract basic configuration values ---
    try:
        PJ_file_name = Bladed_Settings.loc[
            Bladed_Settings["Parameter"] == "PJ file name", "Value"
        ].values[0] + ".$PJ"
        soil_density = Bladed_Settings.loc[
            Bladed_Settings["Parameter"] == "Soil density", "Value"
        ].values[0]
        print_figs = Bladed_Settings.loc[
            Bladed_Settings["Parameter"] == "export PJ validation figures", "Value"
        ].values[0]
        print_figs = str_to_bool(print_figs)
        seabed_level = STRUCTURE_META.loc[
            STRUCTURE_META["Parameter"] == "Seabed level", "Value"
        ].values[0]
    except Exception:
        ex.show_message_box(excel_filename, "Missing or invalid parameters in Bladed_Settings or STRUCTURE_META.")
        return

    # --- Path assignmet ---
    if insert_mode:
        if not os.path.isfile(Bladed_pj_path):
            raise ValueError("Bladed_pj_path has to point to a file")
    else:
        Bladed_pj_path = os.path.join(os.path.dirname(Bladed_pj_path), PJ_file_name)

    # --- Split masses into point and extended types ---
    APPURTANCES = MASSES[MASSES["Top [m]"] != MASSES["Bottom [m]"]]
    ADDITIONAL_MASSES = MASSES[MASSES["Top [m]"] == MASSES["Bottom [m]"]]
    pile_end = GEOMETRY.iloc[-1]["Bottom [m]"]

    # --- Load and process p–y curve data ---
    try:
        PY_data = pe.read_geo_py_curves(py_path)
        if selected_loadcase not in PY_data:
            raise ValueError(f"Selected loadcase '{selected_loadcase}' not found in {py_path}.")
        PY_loadcase = PY_data[selected_loadcase]

        PJ_txt, Interpol_control_FIGs = pe.create_bladed_PJ_py_file(PY_loadcase, pile_end=pile_end, make_plots=print_figs)
    except Exception as err:
        ex.show_message_box(excel_filename,
            f"PY data file could not be read or loadcase '{selected_loadcase}' not found.\nError: {err}"
        )
        ex.set_dropdown_values(excel_filename, "ExportStructure", "Dropdown_Bladed_py_loadcase", [""])
        return

    # --- Save interpolation control figures ---
    if print_figs:
        figs_folder = os.path.join(fig_path, "interpol_control_figs")
        os.makedirs(figs_folder, exist_ok=True)
        for i, fig in enumerate(Interpol_control_FIGs, start=1):
            try:
                fig_out_path = os.path.join(figs_folder, f"interpol_control_fig_{i:03d}.png")
                fig.savefig(fig_out_path, dpi=200, bbox_inches="tight")
            except Exception as err:
                print(f"Warning: Failed to save figure {i}: {err}")

    # --- Run input checks (explicit, one by one for clarity) ---
    ok, err = check_added_masses(ADDITIONAL_MASSES, "ADDITIONAL_MASSES")
    if not ok:
        ex.show_message_box(excel_filename, f"Bladed Structure creation failed, problem with Added masses: {err}")
        return

    ok, err = check_appurtenances(APPURTANCES)
    if not ok:
        ex.show_message_box(excel_filename, f"Bladed Structure creation failed, problem with Appurtenances definition: {err}")
        return

    ok, err = check_marine_growth(MARINE_GROWTH, "MARINE_GROWTH")
    if not ok:
        ex.show_message_box(excel_filename, f"Bladed Structure creation failed, problem with Marine Growth: {err}")
        return

    if seabed_level is None:
        ex.show_message_box(excel_filename, "Seabed level must be provided in STRUCTURE_META when p–y curves are applied.")
        return

    # --- Build Bladed structure data ---
    try:
        PY_loadcase_spring_heights = pd.Series(
            np.unique(PY_loadcase["z [m]"]),
            index=np.unique(PY_loadcase["Spring [-]"])
        )

        Bladed_Elements, Bladed_Nodes = pe.build_Bladed_dataframes(
            Bladed_Settings, Bladed_Material, GEOMETRY, MARINE_GROWTH, MASSES, STRUCTURE_META,
            cut_embedded=False, PY_springs=PY_loadcase_spring_heights, soil_density=soil_density
        )
    except Exception as err:
        ex.show_message_box(excel_filename, f"Error building Bladed dataframes: {err}")
        return

    # --- Generate node definitions and save PJ file ---
    try:
        Nodes_with_spring = Bladed_Nodes.loc[
            Bladed_Nodes["Foundation"].apply(lambda x: isinstance(x, str) and x != ""),
            "Node [-]"
        ].values
        node_def_lines = pe.create_bladed_PJ_node_defnition(Nodes_with_spring)

        TP_top_node_value = Bladed_Elements.loc[(Bladed_Elements["Affiliation [-]"]=="TP") | (Bladed_Elements["Affiliation [-]"]=="MP"), "Node [-]"].values[0]
        TP_top_NODE_idx = Bladed_Nodes.loc[Bladed_Nodes["Node [-]"] == TP_top_node_value, "Node [-]"].index[0]
        seabed_NODE_idx = Bladed_Nodes.loc[Bladed_Nodes["Elevation [m]"] == seabed_level, "Node [-]"].index[0]
        Nodes_interface_mudline = Bladed_Nodes.loc[TP_top_NODE_idx:seabed_NODE_idx, "Node [-]"].values
        Nodes_interface_mudline = list(Nodes_interface_mudline)

        output_node_lines = pe.create_bladed_PJ_output_definition(Nodes_interface_mudline)

        if insert_mode:
            pe.replace_pj_blocks(Bladed_pj_path, PJ_txt, node_def_lines[0], node_def_lines[1], output_node_lines)
            ex.show_message_box(excel_filename, f"PY foundation configuration inserted into PJ file '{Bladed_pj_path}'.")
        else:
            with open(Bladed_pj_path, 'w') as file:
                file.write("\n".join(node_def_lines) + "\n\n\n\n\n" + PJ_txt)
            ex.show_message_box(excel_filename, f"PY foundation configuration saved to PJ file '{Bladed_pj_path}'.")

    except Exception as err:
        ex.show_message_box(excel_filename, f"Failed to write PJ file: {err}")
        return

    # --- Export resulting DataFrames back to Excel ---
    if update_tables:
        try:
            ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Elements", Bladed_Elements)
            ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Nodes", Bladed_Nodes)
        except Exception as err:
            ex.show_message_box(excel_filename, f"Failed to write output tables to Excel: {err}")
            return
    else:
        return


def load_Bladed_soil_file_mat(excel_caller, path):
    excel_filename = os.path.basename(excel_caller)

    try:
        _, sparse, _ = pe.read_soil_stiffness_matrix_csv(path)
        sparse = sparse.T

        sparse.insert(0, "Stiffness", sparse.index)

        ex.clear_excel_table_contents(excel_filename, "ExportStructure", "Bladed_soil_stiffness_mat")
        ex.write_df_to_table_flexible(excel_filename, "ExportStructure", "Bladed_soil_stiffness_mat", sparse)

        ex.set_dropdown_values(excel_filename, "ExportStructure", "Dropdown_Bladed_stiff_mat", list(sparse.columns[1:]))

    except:
        ex.show_message_box(excel_filename, f"PY data file could not be read, make sure it is the right format and it is reachable.")
        ex.clear_excel_table_contents(excel_filename, "ExportStructure", "Bladed_soil_stiffness_mat")
    return


def apply_bladed_stiff_mat(excel_caller, Bladed_stiff_path, Bladed_pj_export_path, config_name):
    excel_filename = os.path.basename(excel_caller)
    Bladed_Nodes = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Nodes", dropnan=True)
    Bladed_Elements = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Elements", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    seabed_level = STRUCTURE_META.loc[
        STRUCTURE_META["Parameter"] == "Seabed level", "Value"
    ].values[0]

    try:
        _, _, stiff_mat  = pe.read_soil_stiffness_matrix_csv(Bladed_stiff_path)

        config_data = stiff_mat[config_name]

        MFONDS_str = pe.create_bladed_pj_stiff_mat_file(config_data, config_name)

        Nodes_with_spring = Bladed_Nodes.loc[Bladed_Nodes["Elevation [m]"] == seabed_level, "Node [-]"].values
        node_def_lines = pe.create_bladed_PJ_node_defnition(Nodes_with_spring)

        TP_top_node_value = Bladed_Elements.loc[(Bladed_Elements["Affiliation"]=="TP") | (Bladed_Elements["Affiliation"]=="MP"), "Node [-]"].values[0]
        TP_top_NODE_idx = Bladed_Nodes.loc[Bladed_Nodes["Node [-]"] == TP_top_node_value, "Node [-]"].index[0]
        seabed_NODE_idx = Bladed_Nodes.loc[Bladed_Nodes["Elevation [m]"] == seabed_level, "Node [-]"].index[0]
        Nodes_interface_mudline = Bladed_Nodes.loc[TP_top_NODE_idx:seabed_NODE_idx, "Node [-]"].values
        Nodes_interface_mudline = list(Nodes_interface_mudline)

        output_node_lines = pe.create_bladed_PJ_output_definition(Nodes_interface_mudline)

        pe.replace_pj_blocks(Bladed_pj_export_path, MFONDS_str, node_def_lines[0], node_def_lines[1],  output_node_lines)

    except Exception as e:
        ex.show_message_box(excel_filename, f"Failed to write PJ file: {e}")

        return

    return

# excel_caller  = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometryConverter/GeometryConverter.xlsm"
# #
# # #fill_Bladed_table(excel_caller)
# # py_path  = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/PY-curves_Bladed/24A525-JBO-TNMPCD-EN-1003-03 - Preliminary MP-TP Concept Design - Annex A1 - Springs_(L).csv"
# # Bladed_pj_path  = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/PY-curves_Bladed/insert_pj_mode/DKT_12_v04_Wdir270_Wavedir300_yen8_s01_____.$PJ"
# # selected_loadcase  = "FLS_(Reloading_BE)"
# # insert_mode = True
# # #
# stiff_Mat_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/PY-curves_Bladed/24A525-JBO-TNMPCD-EN-1003-03 - Preliminary MP-TP Concept Design - Annex A1 - Lateral_Stiffness.csv"
# # #pply_bladed_py_curves(excel_caller, py_path, Bladed_pj_path, selected_loadcase, insert_mode=insert_mode, fig_path=None)
# # # #excel_caller = "I:/2025/A/518_RWE_WBO_FOU_Design/100_Engr/110_Loads/01_LILA/02_preLILA_Vestas/2025-10-14_GeometryConverter_v1.5_MP_DP-C_013Hz_L0_G0_S1-BCe.xlsm"
# # #load_Bladed_soil_file_mat(excel_caller, stiff_Mat_path)
# #
# #
# apply_bladed_stiff_mat(excel_caller, stiff_Mat_path, "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/PY-curves_Bladed/24A525_DP-B4_SG276_21p5_FLS_relo_load_Prod_EOG.prj", "LC1_FLS_reloading_initial")
