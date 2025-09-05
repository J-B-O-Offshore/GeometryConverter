import os.path

import pandas as pd
import excel as ex

import misc as mc
import numpy as np
from pandas.api.types import CategoricalDtype
from ALaPy import periphery as pe
from ALaPy import periphery as pe

import plot as plt
# %% helpers
import os
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

# %% macros

def export_JBOOST(excel_caller, jboost_path):
    excel_filename = os.path.basename(excel_caller)
    jboost_path = os.path.abspath(jboost_path)

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
            windfile="wind.lua"
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

        path_proj = os.path.join(path_config, config_name + ".lua")
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


def export_WLGen(excel_caller, WLGen_path):
    excel_filename = os.path.basename(excel_caller)

    APPURTANCES_MASSES = ex.read_excel_table(excel_filename, "ExportStructure", "APPURTANCES")
    STRUCTURE = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE")
    STRUCTURE = STRUCTURE.drop(columns=["Section", "local Section"])
    MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH")
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
    text, msg = pe.create_WLGen_file(APPURTANCES, ADDITIONAL_MASSES, MP, TP, MARINE_GROWTH, skirt=SKIRT)

    # Feedback to user
    if not text:
        ex.show_message_box(excel_filename, f"WLGen Structure could not be created: {msg}")
    else:
        model_name = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Model Name", "Value"].values[0]
        if model_name is None:
            model_name = "WLGen_input.lua"
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


def fill_Bladed_table(excel_caller):
    excel_filename = os.path.basename(excel_caller)

    # Read inputs
    Bladed_Settings = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Settings", dropnan=True)
    Bladed_Material = ex.read_excel_table(excel_filename, "ExportStructure", "Bladed_Material", dropnan=True)
    GEOMETRY = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
    MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
    MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES")
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")

    # Build dataframes
    Bladed_Elements, Bladed_Nodes = pe.build_Bladed_dataframes(
        Bladed_Settings, Bladed_Material, GEOMETRY, MARINE_GROWTH, MASSES, STRUCTURE_META
    )

    # Write outputs
    ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Elements", Bladed_Elements)
    ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Nodes", Bladed_Nodes)


# def fill_Sesam_table(excel_caller):
#     excel_filename = os.path.basename(excel_caller)
#     Sesam_Settings = ex.read_excel_table(excel_filename, "ExportStructure", "Sesam_Settings", dropnan=True)
#     Sesam_Material = ex.read_excel_table(excel_filename, "ExportStructure", "Sesam_Material", dropnan=True)
#
#     GEOMETRY = ex.read_excel_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", dropnan=True)
#     MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", "MARINE_GROWTH", dropnan=True)
#     MASSES = ex.read_excel_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES")
#     STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
#
#     Sesam_Elements = pd.DataFrame(columns=["Affiliation [-]", "Member [-]", "Node [-]", "Diameter [m]", "Wall thickness [mm]", "cd [-]", "cm [-]", "Marine growth [mm]", "Density [kg*m^-3]", "Material [-]", "Elevation [m]"])
#     Sesam_Nodes = pd.DataFrame(columns=["Node [-]", "Elevation [m]", "Local x [m]", "Local y [m]", "Point mass [kg]"])
#
#     create_node_tolerance = Sesam_Settings.loc[Sesam_Settings["Parameter"] == "Dimensional Tolerance for Node generating [m]", "Value"].values[0]
#     seabed_level = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"].values[0]
#     material = Sesam_Material.loc[0, "Material"]
#     density = Sesam_Material.loc[0, "Density"]
#
#     # Filter geometry below seabed
#     if seabed_level is not None:
#         GEOMETRY = mc.add_element(GEOMETRY, seabed_level)
#     GEOMETRY = GEOMETRY.loc[GEOMETRY["Bottom [m]"] >= seabed_level]
#
#     NODES = mc.extract_nodes_from_elements(GEOMETRY)
#
#     # Add masses
#     NODES["pMass"] = 0.0
#     NODES["pMassNames"] = None
#     NODES["added"] = False
#     NODES["comment"] = None
#
#     if MASSES is not None:
#         for idx in MASSES.index:
#             z_bot = MASSES.loc[idx, "Bottom [m]"]
#             z_Mass = (z_bot + MASSES.loc[idx, "Top [m]"]) / 2 if pd.notna(z_bot) else MASSES.loc[idx, "Top [m]"]
#
#             differences = np.abs(NODES["Elevation [m]"].values - z_Mass)
#             within_tol = np.where(differences <= create_node_tolerance)[0]
#
#             # if node is (nearly) on Node
#             if len(within_tol) > 0:
#                 closest_index = within_tol[np.argmin(differences[within_tol])]
#                 NODES.loc[closest_index, "pMass"] += MASSES.loc[idx, "Mass [kg]"]
#
#                 if NODES.loc[closest_index, "comment"] is None:
#                     NODES.loc[closest_index, "comment"] = MASSES.loc[idx, "Name"] + " "
#                 else:
#                     NODES.loc[closest_index, "comment"] += MASSES.loc[idx, "Name"] + " "
#
#             # if node mass is over bottom
#             elif z_Mass >= GEOMETRY["Bottom [m]"].values[-1]:
#
#                 # add Node
#                 NODES = add_node(NODES, z_Mass, defaults={"float": 0})
#                 GEOMETRY = mc.add_element(GEOMETRY, z_Mass)
#
#                 NODES.loc[NODES["Elevation [m]"] == z_Mass, "added"] = True
#
#                 NODES.loc[NODES["Elevation [m]"] == z_Mass, "pMass"] += MASSES.loc[idx, "Mass [kg]"]
#
#                 NODES.loc[NODES["Elevation [m]"] == z_Mass, "comment"] = MASSES.loc[idx, "Name"] + " "
#
#             else:
#                 print(f"Warning! Mass '{MASSES.loc[idx, 'Name']}' not added, it is below the seabed level!")
#
#     # Nodes
#     Sesam_Nodes.loc[:, "Node [-]"] = np.linspace(1, len(NODES), len(NODES))
#     Sesam_Nodes.loc[:, "Elevation [m]"] = NODES.loc[:, "Elevation [m]"]
#     Sesam_Nodes.loc[:, "Local x [m]"] = 0.0
#     Sesam_Nodes.loc[:, "Local y [m]"] = 0.0
#     Sesam_Nodes.loc[:, "Point mass [kg]"] = NODES.loc[:, "pMass"]
#     Sesam_Nodes.loc[:, "Added"] = NODES.loc[:, "added"]
#     Sesam_Nodes.loc[:, "Comment"] = NODES.loc[:, "comment"]
#
#     # Geometry
#     GEOMETRY.loc[:, "Section"] = np.linspace(1, len(GEOMETRY), len(GEOMETRY))
#     Sesam_Elements.loc[:, "Affiliation [-]"] = np.array([[aff_elem, aff_elem] for aff_elem in GEOMETRY["Affiliation"].values]).flatten()
#     Sesam_Elements.loc[:, "Member [-]"] = np.array([[f"{int(sec_elem)} (End 1)", f"{int(sec_elem)} (End 2)"] for sec_elem in GEOMETRY["Section"].values]).flatten()
#
#     Sesam_Elements.loc[:, "Elevation [m]"] = np.array([[row["Top [m]"], row["Bottom [m]"]] for i, row in GEOMETRY.iterrows()]).flatten()
#
#     for i, row in Sesam_Elements.iterrows():
#
#         # node
#         elevation = row["Elevation [m]"]
#         # Find matching node based on elevation
#         node = Sesam_Nodes.loc[Sesam_Nodes["Elevation [m]"] == elevation, "Node [-]"]
#         if not node.empty:
#             Sesam_Elements.at[i, "Node [-]"] = int(node.values[0])
#
#         marineGrowth = MARINE_GROWTH.loc[(MARINE_GROWTH["Bottom [m]"] < elevation) & (MARINE_GROWTH["Top [m]"] >= elevation), "Marine Growth [mm]"]
#
#         if not marineGrowth.empty:
#             Sesam_Elements.at[i, "Marine growth [mm]"] = marineGrowth.values[0]
#         else:
#             Sesam_Elements.at[i, "Marine growth [mm]"] = 0
#
#     Sesam_Elements.drop(columns=["Elevation [m]"], inplace=True)
#
#     Sesam_Elements.loc[:, "Diameter [m]"] = np.array([[row["D, top [m]"], row["D, bottom [m]"]] for i, row in GEOMETRY.iterrows()]).flatten()
#     Sesam_Elements.loc[:, "Wall thickness [mm]"] = np.array([[row["t [mm]"], row["t [mm]"]] for i, row in GEOMETRY.iterrows()]).flatten()
#
#     Sesam_Elements.loc[:, "cd [-]"] = 0.9
#     Sesam_Elements.loc[:, "cm [-]"] = 2.0
#     Sesam_Elements.loc[:, "Density [kg*m^-3]"] = density
#     Sesam_Elements.loc[:, "Material [-]"] = material
#
#     ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Elements", Sesam_Elements)
#     ex.write_df_to_table(excel_filename, "ExportStructure", "Bladed_Nodes", Sesam_Nodes)
#
#     return
#

def fill_JBOOST_auto_excel(excel_caller):
    """
    Fills missing or automatically assigned values in the `JBOOST_PROJECT` Excel table
    based on defaults and metadata from the `StructureOverview` sheet.

    This function reads the `JBOOST_PROJECT` table from the "ExportStructure" sheet
    and the `STRUCTURE_META` table from the "StructureOverview" sheet of the given
    Excel file. It then replaces missing or "auto" values in project configurations
    with either:
    - default values from the `default` column, or
    - resolved values from the `STRUCTURE_META` table (e.g., water level, seabed level,
      hub height).

    If a required "auto" value cannot be resolved from `STRUCTURE_META`, a message box
    is shown to the user and the process is aborted.

    Parameters
    ----------
    excel_caller : str
        Path or filename of the Excel file containing the required tables.
        Only the basename is used internally.
    """

    excel_filename = os.path.basename(excel_caller)

    PROJECT = ex.read_excel_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", dtype=str)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")

    # Keep original for reference
    PROJECT.index = PROJECT["Project Settings"]

    default = PROJECT.loc[:, "default"]
    proj_configs = PROJECT.iloc[:, 4:]

    for config_name, config_data in proj_configs.items():

        # Fill missing values in config_data with defaults
        config_data = config_data.replace("", np.nan).fillna(default)

        # helper: resolve auto values
        def resolve_auto_value(parameter, config_key, description):
            var = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == parameter, "Value"].values
            if len(var) > 0 and isinstance(var[0], (int, float)):
                PROJECT.at[config_key, config_name] = var[0]
                return True
            else:
                ex.show_message_box(
                    excel_filename,
                    f"Please set {description} in the StructureOverview, as you set {config_key} in {config_name} to 'auto'. Aborting."
                )
                return False

        # only overwrite if original was "auto"
        if config_data.at["water_level"] == "auto":
            if not resolve_auto_value("Water level", "water_level", "a water level"):
                return False
        if config_data.at["seabed_level"] == "auto":
            if not resolve_auto_value("Seabed level", "seabed_level", "a seabed level"):
                return False
        if config_data.at["h_hub"] == "auto":
            if not resolve_auto_value("Hubheight", "h_hub", "Hubheight"):
                return False
        if config_data.at["h_refwindspeed"] == "auto":
            PROJECT.at["h_refwindspeed", config_name] = config_data["h_hub"]

    PROJECT["Project Settings"] = PROJECT.index

    # write back only the modified table
    ex.write_df_to_table(excel_filename, "ExportStructure", "JBOOST_PROJECT", PROJECT)

    return True


def run_JBOOST_excel(excel_caller):

    if fill_JBOOST_auto_excel(excel_caller):

        excel_filename = os.path.basename(excel_caller)

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

            Modeshapes[config_name] = JBOOST_OUT["Mode_shapes"]
            waterlevels[config_name] = config_struct["water_level"]

        reversed_dfs = {k: df.iloc[::-1].reset_index(drop=True) for k, df in Modeshapes.items()}

        FIG = plt.plot_modeshapes(reversed_dfs, order=(1,2,3), waterlevels=waterlevels)

        ex.insert_plot(FIG, excel_filename, "ExportStructure", f"FIG_JBOOST_MODESHAPES")

    return


#run_JBOOST_excel("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometryConverter/GeometryConverter.xlsm")
