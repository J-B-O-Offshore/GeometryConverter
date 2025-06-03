import xlwings
import pandas as pd
import sqlite3
import excel as ex
import numpy as np


def valid_data(data):
    if pd.isna(data.values).any():
        return False, data
    try:
        return True, data.astype(float)
    except (ValueError, TypeError):
        return False, data

def sanity_check_structure(df):
    # check, if sections are on top of each other
    height_diff = (df["Top [m]"].values[1:] - df["Bottom [m]"].values[:-1]) == 0
    if not all(height_diff):
        missaligned_sections = [int(df.iloc[i, 0]) for i, value in enumerate(height_diff) if not value]
        ex.show_message_box("GeometrieConverter.xlsm", f"The {Table} Table Sections are overlapping or have space in between at Section(s): {missaligned_sections} ")
        return False
    else:
        return True

def check_convert_structure(df: pd.DataFrame, Table):
    success, df = valid_data(df)
    if not success:
        ex.show_message_box("GeometrieConverter.xlsm", f"The {Table} Table containes invalid data (nan or non numerical)")
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
        return (np.pi * h / 3) * (r1**2 + r1*r2 + r2**2)

    def com_z_rel(r1, r2, h):
        num = r1**2 + 2*r1*r2 + 3*r2**2
        den = r1**2 + r1*r2 + r2**2
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
        d1**2 + d1 * d2 + d2**2
        - (d1 - 2 * t) ** 2
        - (d1 - 2 * t) * (d2 - 2 * t)
        - (d2 - 2 * t) ** 2
    )

    return rho * volume


def interpolate_node(df, height):
    if len(df.loc[(df["Top [m]"] == height) | (df["Bottom [m]"] == height)].index) > 0:
        return df

    id_inter = df.loc[(df["Top [m]"] > height) & (df["Bottom [m]"] < height)].index
    if len(id_inter) == 0:
        print("interpolation not possible, outside bounds")
        return None
    if len(id_inter) > 1:
        print("interpolation not possible, structure not consecutive")
        return None
    id_inter = id_inter[0]

    new_row = pd.DataFrame(columns=df.columns)
    new_row.loc[0, "Affiliation"] = "TP"
    new_row.loc[0, "t [mm]"] = df.loc[id_inter, "t [mm]"]

    # height
    new_row.loc[0, "Top [m]"] = height
    new_row.loc[0, "Bottom [m]"] = df.loc[id_inter, "Bottom [m]"]

    # diameter
    inter_x_rel = (height - df.loc[id_inter, "Bottom [m]"]) / (df.loc[id_inter, "Top [m]"] - df.loc[id_inter, "Bottom [m]"])
    d_inter = (df.loc[id_inter, "D, top [m]"] - df.loc[id_inter, "D, bottom [m]"]) * inter_x_rel + df.loc[id_inter, "D, bottom [m]"]
    new_row.loc[0, "D, top [m]"] = d_inter
    new_row.loc[0, "D, bottom [m]"] = df.loc[id_inter, "D, bottom [m]"]

    df.loc[id_inter, "Bottom [m]"] = height

    df = pd.concat([df.iloc[:id_inter + 1], new_row, df.iloc[id_inter + 1:]]).reset_index(drop=True)

    return df


def assemble_structure(rho):
    # load structure Data
    MP_DATA = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "MP_DATA")
    TP_DATA = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "TP_DATA")
    TOWER_DATA = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "TOWER_DATA")

    # Quality Checks/Warings of single datasets, if any fail fataly, abort
    sucess_MP, MP_DATA = check_convert_structure(MP_DATA, "MP")
    sucess_TP, TP_DATA = check_convert_structure(TP_DATA, "TP")
    sucess_TOWER, TOWER_DATA = check_convert_structure(TOWER_DATA, "TOWER")

    if not all([sucess_MP, sucess_TP, sucess_TOWER]):
        return

    MP_DATA.insert(0, "Affiliation", "MP")
    TP_DATA.insert(0, "Affiliation", "TP")
    TOWER_DATA.insert(0, "Affiliation", "TOWER")

    # Extract ranges
    range_MP = MP_DATA["Top [m]"].to_list() + list([MP_DATA["Bottom [m]"].values[-1]])
    range_TP = TP_DATA["Top [m]"].to_list() + list([TP_DATA["Bottom [m]"].values[-1]])

    # check MP TP connection
    if range_MP[0] < range_TP[-1]:
        ex.show_message_box("GeometrieConverter.xlsm",
                            f"The Top of the MP at {range_MP[0]} is lower than the Bottom of the TP at {range_TP[-1]}, so the TP is hovering midair at {range_TP[-1] - range_MP[0]}m over the MP. This cant work, aborting.")
        return
    WHOLE_STRUCTURE = MP_DATA

    # Add Weight column:
    # MP_DATA["Weight [t]"] = calc_weight(rho, MP_DATA["t [mm]"].values/1000, MP_DATA["Top [m]"].values, MP_DATA["Bottom [m]"].values, MP_DATA["D, top [m]"].values, MP_DATA["D, bottom [m]"].values)/1000
    # TP_DATA["Weight [t]"] = calc_weight(rho, TP_DATA["t [mm]"].values/1000, TP_DATA["Top [m]"].values, TP_DATA["Bottom [m]"].values, TP_DATA["D, top [m]"].values, TP_DATA["D, bottom [m]"].values)/1000
    # TOWER_DATA["Weight [t]"] = calc_weight(rho, TOWER_DATA["t [mm]"].values/1000, TOWER_DATA["Top [m]"].values, TOWER_DATA["Bottom [m]"].values, TOWER_DATA["D, top [m]"].values, TOWER_DATA["D, bottom [m]"].values)/1000

    # Assemble MP TP
    MP_top = range_MP[0]
    TP_bot = range_TP[-1]
    if MP_top > TP_bot:
        result = ex.show_message_box("GeometrieConverter.xlsm",
                                     f"The MP and the TP are overlapping by {-range_TP[-1] + range_MP[0]}m. Combine stiffness etc as grouted connection (yes) or add as skirt (no)?",
                                     buttons="vbYesNo", icon="vbYesNo", )

        if result == "Yes":

            ex.show_message_box("GeometrieConverter.xlsm",
                                f"under construction...")
        else:

            TP_DATA = interpolate_node(TP_DATA, MP_top)
            SKIRT = TP_DATA.loc[TP_DATA["Top [m]"] <= MP_top]
            SKIRT.loc[:, "Affiliation"] = "SKIRT"
            SKIRT = SKIRT.drop("Section", axis=1)
            skirt_weights = calc_weight(rho, SKIRT["t [mm]"].values / 1000, SKIRT["Top [m]"].values, SKIRT["Bottom [m]"].values, SKIRT["D, top [m]"].values,
                                       SKIRT["D, bottom [m]"].values) / 1000
            skirt_heihgts = center_of_mass_hollow_frustum(SKIRT["D, bottom [m]"].values, SKIRT["D, top [m]"].values, SKIRT["Bottom [m]"], SKIRT["Top [m]"].values, SKIRT["t [mm]"].values / 1000)
            skirt_weight = sum(skirt_weights)

            skirt_center_of_mass = sum([m * h for m, h in zip(list(skirt_weights), list(skirt_heihgts))])/skirt_weight

            # cut TP
            TP_DATA = TP_DATA.loc[TP_DATA["Bottom [m]"] >= MP_top]
            WHOLE_STRUCTURE = pd.concat([TP_DATA, WHOLE_STRUCTURE], axis=0)

            SKIRT_POINTMASS = pd.DataFrame(columns=["Affiliation", "Elevation [m]", "Mass [kg]", "comment"], index=[0])
            SKIRT_POINTMASS.loc[:, "Affiliation"] = "SKIRT"
            SKIRT_POINTMASS.loc[:, "Elevation [m]"] = skirt_center_of_mass
            SKIRT_POINTMASS.loc[:, "Mass [kg]"] = skirt_weight
            SKIRT_POINTMASS.loc[:, "comment"] = "Skirt"

            ex.write_df_to_table("GeometrieConverter.xlsm", "StructureOverview", "SKIRT", SKIRT)
            ex.write_df_to_table("GeometrieConverter.xlsm", "StructureOverview", "SKIRT_POINTMASS", SKIRT_POINTMASS)

    else:
        ex.show_message_box("GeometrieConverter.xlsm", f"The MP and the TP are fitting together perfectly")

        WHOLE_STRUCTURE = pd.concat([TP_DATA, WHOLE_STRUCTURE], axis=0)

    # Add Tower
    tower_offset = WHOLE_STRUCTURE["Top [m]"].values[0] - TOWER_DATA["Bottom [m]"].values[-1]
    TOWER_DATA["Top [m]"] = TOWER_DATA["Top [m]"] + tower_offset
    TOWER_DATA["Bottom [m]"] = TOWER_DATA["Bottom [m]"] + tower_offset

    WHOLE_STRUCTURE = pd.concat([TOWER_DATA, WHOLE_STRUCTURE], axis=0)

    WHOLE_STRUCTURE.rename(columns={"Section": "local Section"}, inplace=True)
    WHOLE_STRUCTURE = WHOLE_STRUCTURE.reset_index(drop=True)
    WHOLE_STRUCTURE.insert(0, "Section", WHOLE_STRUCTURE.index.values + 1)
    ex.write_df_to_table("GeometrieConverter.xlsm", "StructureOverview", "WHOLE_STRUCTURE", WHOLE_STRUCTURE)

    # ADDED MASSES

    MP_MASSES = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "MP_MASSES")
    TP_MASSES = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "TP_MASSES")
    TOWER_MASSES = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "TOWER_MASSES")

    TOWER_MASSES["Elevation [m]"] = TOWER_MASSES["Elevation [m]"] + tower_offset

    MP_MASSES.insert(0, "Affiliation", "MP")
    TP_MASSES.insert(0, "Affiliation", "TP")
    TOWER_MASSES.insert(0, "Affiliation", "TOWER")

    ALL_MASSES = pd.concat([MP_MASSES, TP_MASSES, TOWER_MASSES], axis=0)
    ALL_MASSES.sort_values(inplace=True, ascending=False, axis=0, by=["Elevation [m]"])

    ex.write_df_to_table("GeometrieConverter.xlsm", "StructureOverview", "ALL_ADDED_MASSES", ALL_MASSES)

    return


def move_structure(displ, Structure):
    displ = float(displ)
    META_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META", dtype=str)
    DATA_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA", dtype=float)
    MASSES_CURR = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_MASSES")

    META_CURR.loc[:, "height reference"] = None
    DATA_CURR.loc[:, "Top [m]"] = DATA_CURR.loc[:, "Top [m]"] + displ
    DATA_CURR.loc[:, "Bottom [m]"] = DATA_CURR.loc[:, "Bottom [m]"] + displ
    MASSES_CURR.loc[:, "Elevation [m]"] = MASSES_CURR.loc[:, "Elevation [m]"] + displ

    ex.write_df_to_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_META", META_CURR)
    ex.write_df_to_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_DATA", DATA_CURR)
    ex.write_df_to_table("GeometrieConverter.xlsm", "BuildYourStructure", f"{Structure}_MASSES", MASSES_CURR)


def move_structure_MP(displ):
    move_structure(displ, "MP")

    return


def move_structure_TP(displ):
    move_structure(displ, "TP")

    return
