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

def check_convert_structure(df: pd.DataFrame, Structure):

    success, df = valid_data(df)
    if not success:
        ex.show_message_box("GeometrieConverter.xlsm", "The MP Table containes invalid data (nan or non numerical)")

    # check, if sections are on top of each other
    height_diff = (df["Top [mLAT]"].values[1:] - df["Bottom [mLAT]"].values[:-1]) == 0

    if not all(height_diff):
        missaligned_sections = [int(df.iloc[i, 0]) for i, value in enumerate(height_diff) if not value]
        ex.show_message_box("GeometrieConverter.xlsm", f"The MP Table Sections are overlapping or have space in between at Section(s): {missaligned_sections} ")
        success = False

    return success, df

def assemble_structure(rho):
    def calc_weight(rho, t, z_top, z_bot, d_top, d_bot):
        h = abs(z_top - z_bot)
        d1, d2 = d_top, d_bot
        volume = (1 / 3) * np.pi * h / 4 * (
                d1 ** 2 + d1 * d2 + d2 ** 2
                - (d1 - 2 * t) ** 2
                - (d1 - 2 * t) * (d2 - 2 * t)
                - (d2 - 2 * t) ** 2
        )
        return rho * volume

    # load structure Data
    MP_DATA = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "MP_DATA")
    TP_DATA = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "TP_DATA")
    TOWER_DATA = ex.read_excel_table("GeometrieConverter.xlsm", "BuildYourStructure", "TOWER_DATA")

    # Quality Checks/Warings of single datasets, if any fail fataly, abort
    sucess_MP, MP_DATA = check_convert_structure(MP_DATA, "MP")
    sucess_TP, TP_DATA = check_convert_structure(TP_DATA, "TP")
    sucess_TOWER, TOWER_DATA = check_convert_structure(TOWER_DATA, "TOWER")

    MP_DATA.insert(0, "type", "MP")
    TP_DATA.insert(0, "type", "TP")
    TOWER_DATA.insert(0, "type", "TOWER")

    if not all([sucess_MP, sucess_TP, sucess_TOWER]):
        return

    # Extract ranges
    range_MP = MP_DATA["Top [mLAT]"].to_list() + list([MP_DATA["Bottom [mLAT]"].values[-1]])
    range_TP = TP_DATA["Top [mLAT]"].to_list() + list([TP_DATA["Bottom [mLAT]"].values[-1]])

    # check MP TP connection
    if range_MP[0] < range_TP[-1]:
        ex.show_message_box("GeometrieConverter.xlsm", f"The Top of the MP at {range_MP[0]} is lower than the Bottom of the TP at {range_TP[-1]}, so the TP is hovering midair at {range_TP[-1] - range_MP[0]}m over the MP. This cant work, aborting.")
        return
    WHOLE_STRUCTURE = MP_DATA

    # Add weight column:
    MP_DATA["weight [t]"] = calc_weight(rho, MP_DATA["t [mm]"].values/1000, MP_DATA["Top [mLAT]"].values, MP_DATA["Bottom [mLAT]"].values, MP_DATA["D, top [m]"].values, MP_DATA["D, bottom [m]"].values)/1000
    TP_DATA["weight [t]"] = calc_weight(rho, TP_DATA["t [mm]"].values/1000, TP_DATA["Top [mLAT]"].values, TP_DATA["Bottom [mLAT]"].values, TP_DATA["D, top [m]"].values, TP_DATA["D, bottom [m]"].values)/1000
    TOWER_DATA["weight [t]"] = calc_weight(rho, TOWER_DATA["t [mm]"].values/1000, TOWER_DATA["Top [mLAT]"].values, TOWER_DATA["Bottom [mLAT]"].values, TOWER_DATA["D, top [m]"].values, TOWER_DATA["D, bottom [m]"].values)/1000


    # Assemble MP TP
    if range_MP[0] > range_TP[-1]:
        result = ex.show_message_box("GeometrieConverter.xlsm", f"The MP and the TP are overlapping by {-range_TP[-1] + range_MP[0]}m. Combine stiffness etc as grouted connection (yes) or add as skirt (no)?",  buttons="vbYesNo", icon="vbYesNo",)

        if result == "Yes":

            ex.show_message_box("GeometrieConverter.xlsm",
                                         f"under construction...")
        else:



            ex.show_message_box("GeometrieConverter.xlsm",
                                         f"under construction...")

    else:
        ex.show_message_box("GeometrieConverter.xlsm", f"The MP and the TP are fitting together perfectly")

        WHOLE_STRUCTURE = pd.concat([TP_DATA,WHOLE_STRUCTURE], axis=0)

    # Add Tower
    tower_offset = WHOLE_STRUCTURE["Top [mLAT]"].values[0] - TOWER_DATA["Bottom [mLAT]"].values[-1]
    TOWER_DATA["Top [mLAT]"] = TOWER_DATA["Top [mLAT]"] + tower_offset
    TOWER_DATA["Bottom [mLAT]"] = TOWER_DATA["Bottom [mLAT]"] + tower_offset

    WHOLE_STRUCTURE = pd.concat([TOWER_DATA, WHOLE_STRUCTURE], axis=0)

    WHOLE_STRUCTURE.rename(columns={"Section": "local Section"}, inplace=True)
    WHOLE_STRUCTURE = WHOLE_STRUCTURE.reset_index(drop=True)
    WHOLE_STRUCTURE.insert(0, "Section", WHOLE_STRUCTURE.index.values + 1)
    ex.write_df_to_table("GeometrieConverter.xlsm", "StructureOverview", "WHOLE_STRUCTURE", WHOLE_STRUCTURE)

    return

assemble_structure(7000)