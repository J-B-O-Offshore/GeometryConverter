import os
import numpy as np
import pandas as pd
import excel as ex
import plot as GCplt


def valid_data(data):
    if pd.isna(data.values).any():
        return False, data
    try:
        return True, data.astype(float)
    except (ValueError, TypeError):
        return False, data


def sanity_check_structure(excel_filename, df):
    # check, if sections are on top of each other
    height_diff = (df["Top [m]"].values[1:] - df["Bottom [m]"].values[:-1]) == 0
    if not all(height_diff):
        missaligned_sections = [int(df.iloc[i, 0]) for i, value in enumerate(height_diff) if not value]
        ex.show_message_box(excel_filename, f"The sections overlap or have gaps between them - at section(s): {missaligned_sections}.")
        return False
    else:
        return True


def check_convert_structure(excel_filename, df: pd.DataFrame, Table):
    success, df = valid_data(df)
    if not success:
        ex.show_message_box(excel_filename, f"The {Table} table contains invalid data (nan or non numerical).")
        return success, df

    success = sanity_check_structure
    return success, df


def center_of_mass_hollow_frustum(d1, d2, z_bot, z_top, t):
    """
    Calculates the center of mass (z-coordinate) of a hollow conical frustum
    (truncated cone) with constant wall thickness, based on absolute z-positions.

    Supports scalar, list, or numpy array input (all inputs must be same shape).

    Parameters:
    d1     : float, list, or np.ndarray - Inner diameter at the bottom
    d2     : float, list, or np.ndarray - Inner diameter at the top
    z_bot  : float, list, or np.ndarray - z-position of the bottom surface
    z_top  : float, list, or np.ndarray - z-position of the top surface
    t      : float, list, or np.ndarray - Constant wall thickness

    Returns:
    z_cm   : float or np.ndarray - z-position of the center of mass
    """
    # Convert to numpy arrays
    d1 = np.asarray(d1, dtype=np.float64)
    d2 = np.asarray(d2, dtype=np.float64)
    z_bot = np.asarray(z_bot, dtype=np.float64)
    z_top = np.asarray(z_top, dtype=np.float64)
    t = np.asarray(t, dtype=np.float64)

    # Compute height and radii
    h = z_top - z_bot
    r1 = d1 / 2
    r2 = d2 / 2
    R1 = r1 + t
    R2 = r2 + t

    # Volume and center of mass for solid frustum
    def volume(r1, r2, h):
        return (np.pi * h / 3) * (r1 ** 2 + r1 * r2 + r2 ** 2)

    def com_z_rel(r1, r2, h):
        num = r1 ** 2 + 2 * r1 * r2 + 3 * r2 ** 2
        den = r1 ** 2 + r1 * r2 + r2 ** 2
        return h * num / (4 * den)

    # Compute relative center of mass (from bottom), then convert to absolute z
    V_outer = volume(R1, R2, h)
    V_inner = volume(r1, r2, h)
    z_outer_rel = com_z_rel(R1, R2, h)
    z_inner_rel = com_z_rel(r1, r2, h)

    z_cm_rel = (z_outer_rel * V_outer - z_inner_rel * V_inner) / (V_outer - V_inner)
    z_cm_abs = z_bot + z_cm_rel

    return z_cm_abs


def calc_weight(rho, t, z_top, z_bot, d_top, d_bot):
    rho = np.asarray(rho)
    t = np.asarray(t)
    z_top = np.asarray(z_top)
    z_bot = np.asarray(z_bot)
    d_top = np.asarray(d_top)
    d_bot = np.asarray(d_bot)

    h = np.abs(z_top - z_bot)
    d1 = d_top
    d2 = d_bot

    volume = (1 / 3) * np.pi * h / 4 * (
            d1 ** 2 + d1 * d2 + d2 ** 2
            - (d1 - 2 * t) ** 2
            - (d1 - 2 * t) * (d2 - 2 * t)
            - (d2 - 2 * t) ** 2
    )

    return rho * volume


def add_element(df, z_new, defaults=None, add_outside_bound=False):
    """
    Inserts an interpolated structural segment into a DataFrame at a specified height.

    The function adds a new row at height `z_new` by splitting an existing segment in the
    DataFrame. The existing segment is divided into two parts: the lower part is updated
    to end at `z_new`, and a new upper part is created starting at `z_new`, using linear
    interpolation to compute the diameter at that height.

    If `z_new` matches an existing "Top [m]" or "Bottom [m]" value, or lies outside the
    structure bounds and `add_outside_bound` is False, the original DataFrame is returned unchanged.

    Parameters
    ----------
    df : pandas.DataFrame
        Structural DataFrame with the following required columns:
            - "Top [m]": Top height of each segment (in meters)
            - "Bottom [m]": Bottom height of each segment (in meters)
            - "D, top [m]": Diameter at the top of the segment (in meters)
            - "D, bottom [m]": Diameter at the bottom of the segment (in meters)
            - "t [mm]": Wall thickness (in millimeters)
        Optionally, the column:
            - "Affiliation": Metadata copied into the new segment if present.
        Other columns are allowed and will be preserved. For these, the inserted row will
        contain a default value based on the column's data type:
            - float or int → np.nan
            - bool → False
            - object (e.g. string) → None
            - datetime64 → pd.NaT
            - other/unknown types → None (fallback)

    z_new : float
        Height (in meters) at which to insert the new node.

    defaults : dict, optional
        Dictionary mapping column names to custom default values for the inserted row.

    add_outside_bound : bool, default False
        If False, prevents insertion of nodes outside the vertical extent of the structure.

    Returns
    -------
    pandas.DataFrame
        Updated DataFrame with the new interpolated row inserted.
        Returns original DataFrame if:
        - z_new already exists
        - z_new is outside the structure bounds and add_outside_bound=False
        - z_new lies within overlapping segments (non-unique match)
    """
    df = df.reset_index(drop=True)
    defaults = defaults or {}

    # Skip if height already exists
    if len(df.loc[(df["Top [m]"] == z_new) | (df["Bottom [m]"] == z_new)]) > 0:
        return df

    id_inter = df.loc[(df["Top [m]"] > z_new) & (df["Bottom [m]"] < z_new)].index

    if len(id_inter) == 0:
        if not add_outside_bound:
            print("Interpolation not possible, outside bounds.")
            return df
        else:
            print("No segment contains z_new, but add_outside_bound=True. No interpolation performed.")
            return df

    if len(id_inter) > 1:
        print("Interpolation not possible, structure not consecutive.")
        return df

    id_inter = id_inter[0]
    df = df.copy()
    row_base = df.loc[id_inter]
    new_row = {}

    # Required fields
    new_row["Top [m]"] = z_new
    new_row["Bottom [m]"] = row_base["Bottom [m]"]
    new_row["t [mm]"] = row_base["t [mm]"]

    if "Affiliation" in df.columns:
        new_row["Affiliation"] = row_base["Affiliation"]

    # Diameter interpolation
    inter_x_rel = (z_new - row_base["Bottom [m]"]) / (row_base["Top [m]"] - row_base["Bottom [m]"])
    d_inter = (row_base["D, top [m]"] - row_base["D, bottom [m]"]) * inter_x_rel + row_base["D, bottom [m]"]

    new_row["D, top [m]"] = d_inter
    new_row["D, bottom [m]"] = row_base["D, bottom [m]"]

    # Fill all other columns using defaults or type-based rules
    for col in df.columns:
        if col not in new_row:
            if col in defaults:
                new_row[col] = defaults[col]
            else:
                dtype = df[col].dtype
                if pd.api.types.is_bool_dtype(dtype):
                    new_row[col] = False
                elif pd.api.types.is_numeric_dtype(dtype):
                    new_row[col] = np.nan
                elif pd.api.types.is_datetime64_any_dtype(dtype):
                    new_row[col] = pd.NaT
                elif pd.api.types.is_object_dtype(dtype):
                    new_row[col] = None
                else:
                    new_row[col] = None  # Fallback

    # Update existing row and insert new one
    df.loc[id_inter, "Bottom [m]"] = z_new
    df = pd.concat([df.iloc[:id_inter + 1], pd.DataFrame([new_row]), df.iloc[id_inter + 1:]]).reset_index(drop=True)

    return df


def assemble_structure(MP_DATA, TP_DATA, TOWER_DATA=None, MP_MASSES=None, TP_MASSES=None, TOWER_MASSES=None, excel_caller=None, interactive=True, rho=7900, ignore_hovering=False, overlapp_mode="Skirt"):
    """
    Assemble the full offshore wind turbine structure from Monopile (MP), Transition Piece (TP),
    and Tower data. Handles geometric overlaps, continuity, and integrates additional mass data.

    Parameters
    ----------
    MP_DATA : pd.DataFrame
        Structural data for the Monopile section. Must include columns:
        - "Top [m]", "Bottom [m]", "t [mm]", "D, top [m]", "D, bottom [m]"

    TP_DATA : pd.DataFrame
        Structural data for the Transition Piece. Same required columns as `MP_DATA`.

    TOWER_DATA : pd.DataFrame
        Structural data for the Tower. Same required columns as `MP_DATA`.

    MP_MASSES : pd.DataFrame, optional
        Optional additional point or distributed masses associated with the Monopile.

    TP_MASSES : pd.DataFrame, optional
        Optional additional point or distributed masses associated with the Transition Piece.

    TOWER_MASSES : pd.DataFrame, optional
        Optional additional point or distributed masses associated with the Tower.

    excel_caller : object, optional
        Excel interface object used for displaying interactive message boxes. Required if `interactive=True`.

    interactive : bool, default=True
        If True, prompts the user to resolve overlaps and connection issues via message boxes.
        Requires `excel_caller` to be defined.

    rho : float, default=7900
        Density in kg/m³ used for calculating skirt weight. Default assumes steel.

    ignore_hovering : bool, default=False
        If True, allows structures to remain disconnected (e.g., TP hovering above MP) without raising an error.

    overlapp_mode : {"Skirt", "Grout"}, default="Skirt"
        Mode for resolving overlaps between MP and TP if not interactive:
        - "Skirt": add overlapping TP section to MP as a skirt and compute mass.
        - "Grout": placeholder mode; currently not implemented.

    Returns
    -------
    WHOLE_STRUCTURE : pd.DataFrame
        Combined structural DataFrame for the MP, TP, and Tower sections.
        Includes a "Section" column with continuous indexing and an "Affiliation" column indicating origin.

    ALL_MASSES : pd.DataFrame or None
        Combined and elevation-adjusted DataFrame of all additional masses.
        Returns None if no mass data was provided.

    SKIRT : pd.DataFrame or None
        If a skirt is added to resolve MP-TP overlap, returns the structural data for the skirt.
        Otherwise, returns None.

    SKIRT_POINTMASS : pd.DataFrame or None
        If a skirt is added, returns a single-row DataFrame representing the skirt's equivalent point mass.
        Otherwise, returns None.

    Notes
    -----
    - The function modifies the input DataFrames by adding an "Affiliation" column and may insert or alter elevation values.
    - If `interactive=True` and `excel_caller=None`, a ValueError is raised.
    - The top of the MP must not be lower than the bottom of the TP unless `ignore_hovering=True`.
    - If overlap is detected, it is resolved by either:
        - User decision (interactive mode), or
        - Automatically using the `overlapp_mode` setting (non-interactive).
    - The function currently only implements the "Skirt" resolution mode.
    - Elevations are adjusted to ensure structural continuity from MP → TP → Tower.
    """
    if "Affiliation" not in MP_DATA.columns:
        MP_DATA.insert(0, "Affiliation", "MP")
    else:
        MP_DATA["Affiliation"] = "MP"

    if "Affiliation" not in TP_DATA.columns:
        TP_DATA.insert(0, "Affiliation", "TP")
    else:
        TP_DATA["Affiliation"] = "TP"
    if TOWER_DATA is not None:
        if "Affiliation" not in TOWER_DATA.columns:
            TOWER_DATA.insert(0, "Affiliation", "TOWER")
        else:
            TOWER_DATA["Affiliation"] = "TOWER"

    SKIRT = None
    SKIRT_POINTMASS = None
    # Extract ranges
    range_MP = MP_DATA["Top [m]"].to_list() + list([MP_DATA["Bottom [m]"].values[-1]])
    range_TP = TP_DATA["Top [m]"].to_list() + list([TP_DATA["Bottom [m]"].values[-1]])

    if interactive:
        if excel_caller is None:
            raise ValueError("excel_caller is None, has to be defined when interactive is True")

    WHOLE_STRUCTURE = MP_DATA

    # Assemble MP TP
    MP_top = range_MP[0]
    TP_bot = range_TP[-1]

    if MP_top > TP_bot:

        if interactive:
            result = ex.show_message_box(excel_caller,
                                         f"The MP and the TP are overlapping by {-range_TP[-1] + range_MP[0]}m. Combine stiffness etc as grouted connection (yes) or add as skirt (no)?",
                                         buttons="vbYesNo", icon="vbYesNo", )
        else:
            if overlapp_mode == "Grout":
                result = "Yes"
            elif overlapp_mode == "Skirt":
                result = "No"
            else:
                raise ValueError("overlapp_mode has to be Skirt or Grout.")

        if result == "Yes":

            ex.show_message_box(excel_caller,
                                f"under construction...")
            return

        elif result == "No":

            TP_DATA = add_element(TP_DATA, MP_top)
            SKIRT = TP_DATA.loc[TP_DATA["Top [m]"] <= MP_top]
            SKIRT.loc[:, "Affiliation"] = "SKIRT"
            SKIRT = SKIRT.drop("Section", axis=1)
            skirt_weights = calc_weight(rho, SKIRT["t [mm]"].values / 1000, SKIRT["Top [m]"].values, SKIRT["Bottom [m]"].values, SKIRT["D, top [m]"].values,
                                        SKIRT["D, bottom [m]"].values) / 1000
            skirt_heihgts = center_of_mass_hollow_frustum(SKIRT["D, bottom [m]"].values, SKIRT["D, top [m]"].values, SKIRT["Bottom [m]"], SKIRT["Top [m]"].values,
                                                          SKIRT["t [mm]"].values / 1000)
            skirt_weight = sum(skirt_weights)

            skirt_center_of_mass = sum([m * h for m, h in zip(list(skirt_weights), list(skirt_heihgts))]) / skirt_weight

            # cut TP
            TP_DATA = TP_DATA.loc[TP_DATA["Bottom [m]"] >= MP_top]
            WHOLE_STRUCTURE = pd.concat([TP_DATA, WHOLE_STRUCTURE], axis=0)

            SKIRT_POINTMASS = pd.DataFrame(columns=["Affiliation", "Elevation [m]", "Mass [t]", "comment"], index=[0])
            SKIRT_POINTMASS.loc[:, "Affiliation"] = "SKIRT"
            SKIRT_POINTMASS.loc[:, "Elevation [m]"] = skirt_center_of_mass
            SKIRT_POINTMASS.loc[:, "Mass [t]"] = skirt_weight
            SKIRT_POINTMASS.loc[:, "comment"] = "Skirt"

    elif MP_top < TP_bot:
        if not ignore_hovering:
            if interactive:
                ex.show_message_box(excel_caller,
                                    f"The top of the MP at {range_MP[0]} is lower than the bottom of the TP at {range_TP[-1]}, so the TP is hovering midair at {range_TP[-1] - range_MP[0]}m over the MP. This can't work, aborting.")

            raise ValueError
        else:
            if interactive:
                ex.show_message_box(excel_caller,
                                        f"The top of the MP at {range_MP[0]} is lower than the bottom of the TP at {range_TP[-1]}, so the TP is hovering midair at {range_TP[-1] - range_MP[0]}m over the MP. Not aborting because of function setting ignore_hovering=True.")

            WHOLE_STRUCTURE = pd.concat([TP_DATA, WHOLE_STRUCTURE], axis=0)
    else:
        if interactive:
            ex.show_message_box(excel_caller, f"The MP and the TP are fitting together perfectly.")

        WHOLE_STRUCTURE = pd.concat([TP_DATA, WHOLE_STRUCTURE], axis=0)

    if TOWER_DATA is not None:
        # Add Tower
        WHOLE_STRUCTURE, tower_offset = add_tower_on_top(WHOLE_STRUCTURE, TOWER_DATA)

        WHOLE_STRUCTURE.rename(columns={"Section": "local Section"}, inplace=True)
        WHOLE_STRUCTURE = WHOLE_STRUCTURE.reset_index(drop=True)
        WHOLE_STRUCTURE.insert(0, "Section", WHOLE_STRUCTURE.index.values + 1)

    all_masses = []
    if MP_MASSES is not None:
        MP_MASSES.insert(0, "Affiliation", "MP")
        all_masses.append(MP_MASSES)

    if TP_MASSES is not None:
        TP_MASSES.insert(0, "Affiliation", "TP")
        all_masses.append(TP_MASSES)

    if TOWER_MASSES is not None:
        TOWER_MASSES["Top [m]"] = TOWER_MASSES["Top [m]"] + tower_offset
        mask = pd.to_numeric(TOWER_MASSES["Bottom [m]"], errors='coerce').notna()
        TOWER_MASSES.loc[mask, "Bottom [m]"] += tower_offset
        TOWER_MASSES.insert(0, "Affiliation", "TOWER")
        all_masses.append(TOWER_MASSES)

    if len(all_masses) != 0:
        ALL_MASSES = pd.concat(all_masses, axis=0)
        ALL_MASSES.sort_values(inplace=True, ascending=False, axis=0, by=["Top [m]"])
    else:
        ALL_MASSES = None

    return WHOLE_STRUCTURE, ALL_MASSES, SKIRT, SKIRT_POINTMASS


def assemble_structure_excel(excel_caller, rho, RNA_config):
    def all_same_ignoring_none(*values):
        non_none = [v for v in values if v is not None]
        return len(non_none) <= 1 or all(v == non_none[0] for v in non_none)

    excel_filename = os.path.basename(excel_caller)
    # load structure Data
    MP_DATA = ex.read_excel_table(excel_filename, "BuildYourStructure", "MP_DATA")
    TP_DATA = ex.read_excel_table(excel_filename, "BuildYourStructure", "TP_DATA")
    TOWER_DATA = ex.read_excel_table(excel_filename, "BuildYourStructure", "TOWER_DATA")
    RNA_DATA = ex.read_excel_table(excel_filename, "BuildYourStructure", "RNA_DATA")

    MP_META = ex.read_excel_table(excel_filename, "BuildYourStructure", "MP_META")
    TP_META = ex.read_excel_table(excel_filename, "BuildYourStructure", "TP_META")
    TOWER_META = ex.read_excel_table(excel_filename, "BuildYourStructure", "TOWER_META")
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", "STRUCTURE_META")
    STRUCTURE_META.loc[:, "Value"] = ""

    MP_MASSES = ex.read_excel_table(excel_filename, "BuildYourStructure", "MP_MASSES", dropnan=True)
    TP_MASSES = ex.read_excel_table(excel_filename, "BuildYourStructure", "TP_MASSES", dropnan=True)
    TOWER_MASSES = ex.read_excel_table(excel_filename, "BuildYourStructure", "TOWER_MASSES", dropnan=True)

    if len(MP_DATA) == 0:
        ex.show_message_box(excel_filename, f"Please provide a MP structure to assamble.")
        return
    if len(TP_DATA) == 0:
        ex.show_message_box(excel_filename, f"Please provide a TP structure to assamble.")
        return

    # Quality Checks/Warings of single datasets, if any fail fataly, abort
    sucess_MP, MP_DATA = check_convert_structure(excel_filename, MP_DATA, "MP")
    sucess_TP, TP_DATA = check_convert_structure(excel_filename, TP_DATA, "TP")
    sucess_TOWER, TOWER_DATA = check_convert_structure(excel_filename, TOWER_DATA, "TOWER")

    if len(TOWER_DATA) == 0:
        TOWER_DATA = None
    if len(TOWER_MASSES) == 0:
        TOWER_MASSES = None

    if not all([sucess_MP, sucess_TP, sucess_TOWER]):
        return

    # RNA choosing
    if RNA_config == "":
        ex.show_message_box(excel_filename,
                            f"Caution, no RNA selected.")
    else:
        if not RNA_config in RNA_DATA["Identifier"].values:
            ex.show_message_box(excel_filename,
                                f"Chosen RNA not in RNA dropdown menu. Aborting.")
            return None
        else:
            RNA = RNA_DATA.loc[RNA_DATA["Identifier"] == RNA_config, :]
            ex.write_df_to_table(excel_filename, "StructureOverview", "RNA", RNA)

    # Height Reference handling
    WL_ref_MP = MP_META.loc[0, "Height Reference"]
    WL_ref_MT = TP_META.loc[0, "Height Reference"]
    WL_ref_TOWER = TOWER_META.loc[0, "Height Reference"]

    if not all_same_ignoring_none(WL_ref_MP, WL_ref_MT, WL_ref_TOWER):
        answer = ex.show_message_box(excel_filename,
                                     f"Warning, not all height references are the same (MP: {WL_ref_MP}, TP: {WL_ref_MT}, TOWER: {WL_ref_TOWER}). Assemble anyway?",
                                     buttons="vbYesNo", icon="warning")
        if answer == "No":
            return
    else:
        STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Height Reference", "Value"] = [v for v in [WL_ref_MP, WL_ref_MT, WL_ref_TOWER] if v is not None][0]
        ex.show_message_box(excel_filename,
                            f"Height references are the same or not defined. (MP: {WL_ref_MP}, TP: {WL_ref_MT}, TOWER: {WL_ref_TOWER}).")

    # waterdepth handling
    if MP_META.loc[0, "Water Depth [m]"] is not None:
        STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"] = - float(MP_META.loc[0, "Water Depth [m]"])

    try:
        WHOLE_STRUCTURE, ALL_MASSES, SKIRT, SKIRT_POINTMASS = assemble_structure(MP_DATA, TP_DATA, TOWER_DATA, MP_MASSES=MP_MASSES, TP_MASSES=TP_MASSES, TOWER_MASSES=TOWER_MASSES,
                                                                                 excel_caller=excel_filename, rho=rho)
    except ValueError:
        GCplt.plot_Assambly_Overview(excel_caller)
        return

    ex.write_df_to_table(excel_filename, "StructureOverview", "WHOLE_STRUCTURE", WHOLE_STRUCTURE)
    ex.write_df_to_table(excel_filename, "StructureOverview", "ALL_ADDED_MASSES", ALL_MASSES)
    ex.write_df_to_table(excel_filename, "StructureOverview", "STRUCTURE_META", STRUCTURE_META)

    if SKIRT is not None:
        ex.write_df_to_table(excel_filename, "StructureOverview", "SKIRT", SKIRT)
    if SKIRT_POINTMASS is not None:
        ex.write_df_to_table(excel_filename, "StructureOverview", "SKIRT_POINTMASS", SKIRT_POINTMASS)

    # plot assambly
    GCplt.plot_Assambly_Overview(excel_caller)

    return


def move_structure(excel_filename, displ, Structure):
    try:
        displ = float(displ)
    except ValueError:
        ex.show_message_box(excel_filename, f"Please enter a valid float value for the displacement.")
        return
    META_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_META", dtype=str)
    DATA_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_DATA", dtype=float)
    MASSES_CURR = ex.read_excel_table(excel_filename, "BuildYourStructure", f"{Structure}_MASSES")

    META_CURR.loc[:, "Height Reference"] = None
    DATA_CURR.loc[:, "Top [m]"] = DATA_CURR.loc[:, "Top [m]"] + displ
    DATA_CURR.loc[:, "Bottom [m]"] = DATA_CURR.loc[:, "Bottom [m]"] + displ
    MASSES_CURR.loc[:, "Top [m]"] = MASSES_CURR.loc[:, "Top [m]"] + displ
    MASSES_CURR.loc[:, "Bottom [m]"] = MASSES_CURR.loc[:, "Bottom [m]"] + displ

    ex.write_df_to_table(excel_filename, "BuildYourStructure", f"{Structure}_META", META_CURR)
    ex.write_df_to_table(excel_filename, "BuildYourStructure", f"{Structure}_DATA", DATA_CURR)
    ex.write_df_to_table(excel_filename, "BuildYourStructure", f"{Structure}_MASSES", MASSES_CURR)


def move_structure_MP(excel_caller, displ):
    excel_filename = os.path.basename(excel_caller)

    move_structure(excel_filename, displ, "MP")

    return


def move_structure_TP(excel_caller, displ):
    excel_filename = os.path.basename(excel_caller)

    move_structure(excel_filename, displ, "TP")

    return


def extract_nodes_from_elements(df_elements: pd.DataFrame) -> pd.DataFrame:
    """
        Generate a DataFrame of nodes from a given element-based DataFrame.

        Each node represents a unique depth where elements meet (top or bottom).
        If two adjacent elements have different 'Affiliation' values, the node is marked as 'BORDER',
        otherwise it inherits the common affiliation.

        Parameters:
        -----------
        df_elements : pd.DataFrame
            A DataFrame containing element definitions with the following required columns:
            - 'Section' (optional, for indexing)
            - 'Affiliation' (str): the type or group of the element (e.g., TOWER, TP, MP)
            - 'Top [m]' (float): the top elevation of the element
            - 'Bottom [m]' (float): the bottom elevation of the element

        Returns:
        --------
        pd.DataFrame
            A DataFrame of nodes with the following columns:
            - 'Node' (int): sequential node ID
            - 'Elevation [m]' (float): the depth of the node
            - 'Affiliation' (str): 'BORDER' if adjacent affiliations differ, else the shared affiliation
        """

    # Ensure sorting by Top depth (descending), in case not already sorted
    #df_sorted = df_elements.sort_values(by='Top [m]', ascending=False).reset_index(drop=True)

    nodes = []

    for i, row in df_elements.iterrows():
        # Top node of the first element
        if i == 0:
            nodes.append({
                'node': len(nodes) + 1,
                'Elevation [m]': row['Top [m]'],
                'Affiliation': row['Affiliation']
            })

        # Bottom node of current element
        bottom_elev = row['Bottom [m]']

        if i + 1 < len(df_elements):
            next_affiliation = df_elements.loc[i + 1, 'Affiliation']
        else:
            next_affiliation = row['Affiliation']  # Last element — use same

        node_affiliation = (
            'BORDER'
            if row['Affiliation'] != next_affiliation
            else row['Affiliation']
        )

        nodes.append({
            'node': len(nodes) + 1,
            'Elevation [m]': bottom_elev,
            'Affiliation': node_affiliation
        })

    return pd.DataFrame(nodes)


def add_tower_on_top(STRUCTURE, TOWER):
    """
       Attach a tower section on top of an existing structure by vertically aligning
       their boundaries and merging them into one continuous structure.

       The function computes the required vertical offset so that the bottom of the
       tower aligns exactly with the top of the existing structure. It then shifts
       the tower section accordingly and concatenates both parts.

       Parameters
       ----------
       STRUCTURE : pandas.DataFrame
           DataFrame describing the existing structure with at least the columns
           "Top [m]" and "Bottom [m]". The topmost elevation is taken from the first row.
       TOWER : pandas.DataFrame
           DataFrame describing the tower section with at least the columns
           "Top [m]" and "Bottom [m]". The bottommost elevation is taken from the last row.

       Returns
       -------
       WHOLE_STRUCTURE : pandas.DataFrame
           Combined DataFrame representing the full structure with the tower placed
           on top of the existing structure.

       Notes
       -----
       - The order of rows in the resulting DataFrame is `[TOWER, STRUCTURE]`.
       - Assumes both inputs follow a consistent convention for elevations in meters.
       """
    tower_offset = STRUCTURE["Top [m]"].values[0] - TOWER["Bottom [m]"].values[-1]
    TOWER["Top [m]"] = TOWER["Top [m]"] + tower_offset
    TOWER["Bottom [m]"] = TOWER["Bottom [m]"] + tower_offset

    WHOLE_STRUCTURE = pd.concat([TOWER, STRUCTURE], axis=0)

    return WHOLE_STRUCTURE, tower_offset
