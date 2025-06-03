import xlwings
import pandas as pd
import sqlite3
import excel as ex
import misc as mc
from typing import Tuple

import numpy as np


def export_JBOOST(txt_path):
    def _convert_to_rad(value: float, unit: str) -> float:
        """Convert deflection from mm/m or degrees to radians."""
        if unit == "deg":
            return np.deg2rad(value)
        elif unit == "mm/m":
            return np.arctan(value / 1000)
        else:
            raise ValueError(f"Unsupported unit '{unit}'. Use 'deg' or 'mm/m'.")

    def _compute_line(z: pd.Series, angle_rad: float, base_z: float) -> pd.Series:
        """Compute the deflection line based on angle and base z level."""
        return np.sin(angle_rad) * (z - base_z)

    def calculate_deflection(
            NODES: pd.DataFrame,
            defl_MP: Tuple[float, str],
            defl_Tower: Tuple[float, str],
            defl_TP: Optional[Tuple[float, str]] = None
    ) -> pd.Series:
        """
        Calculate deflection values based on given tilt angles.

        Parameters:
        - NODES: DataFrame with columns ["z", "Affiliation"]
        - defl_MP, defl_Tower, defl_TP: Tuples (value, unit), where unit is "deg" or "mm/m"

        Returns:
        - pd.Series of deflection values
        """
        affiliations = NODES["Affiliation"]
        z = NODES["z"]
        base_z = z.iloc[-1]

        # Convert all given angles to radians
        angle_MP = _convert_to_rad(*defl_MP)
        angle_Tower = _convert_to_rad(*defl_Tower)
        angle_TP = _convert_to_rad(*defl_TP) if defl_TP else None

        if angle_TP is None and "TP" in affiliations.values:
            raise ValueError("defl_TP must be provided if 'TP' is present in Affiliation.")

        # Compute initial deflection lines
        line_MP = _compute_line(z, angle_MP, base_z)
        line_Tower = _compute_line(z, angle_Tower, base_z)
        line_TP = _compute_line(z, angle_TP, base_z) if angle_TP else None

        # Initialize DEFL column
        NODES["DEFL"] = 0.0

        # Assign MP deflection
        mask_MP = affiliations == "MP"
        NODES.loc[mask_MP, "DEFL"] = line_MP[mask_MP]

        if angle_TP is not None:
            mask_TP = affiliations == "TP"
            TP_offset = line_MP[mask_MP].iloc[0] - line_TP[mask_MP].iloc[0]
            NODES.loc[mask_TP, "DEFL"] = line_TP[mask_TP] + TP_offset

            mask_TOWER = affiliations == "TOWER"
            TOWER_offset = NODES.loc[mask_TP, "DEFL"].iloc[0] - line_Tower[mask_TP].iloc[0]
            NODES.loc[mask_TOWER, "DEFL"] = line_Tower[mask_TOWER] + TOWER_offset
        else:
            mask_TOWER = affiliations == "TOWER"
            TOWER_offset = line_MP[mask_MP].iloc[0] - line_Tower[mask_MP].iloc[0]
            NODES.loc[mask_TOWER, "DEFL"] = line_Tower[mask_TOWER] + TOWER_offset

        return NODES["DEFL"]

    text = ""
    NODES = []
    GEOMETRY = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "WHOLE_STRUCTURE")
    MASSES = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "ALL_ADDED_MASSES")
    MARINE_GROWTH = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "MARINE_GROWTH")

    sucess_GEOMETRY = mc.sanity_check_structure(GEOMETRY)

    if not sucess_GEOMETRY:
        return

    N_Elem = len(GEOMETRY)
    N_Nodes = len(GEOMETRY) + 1

    NODES = pd.DataFrame({
        "z": pd.concat([
            GEOMETRY["Top [m]"],
            pd.Series([GEOMETRY["Bottom [m]"].iloc[-1]])
        ], ignore_index=True)
    })
    nodes = [501+i for i in range(N_Nodes)]
    nodes.reverse()
    NODES["node"] = nodes
    NODES["pInertia"] = 0
    NODES["pMass"] = 0.0
    NODES["Affiliation"] = "NOT DEFINDED"
    NODES.loc[NODES.index[:-1], "Affiliation"] = GEOMETRY.iloc[:]["Affiliation"]
    NODES.loc[NODES.index[-1], "Affiliation"] = GEOMETRY.iloc[-1]["Affiliation"]

    # distribute Masses on Nodes
    for idx in MASSES.index:
        z_Mass = MASSES.loc[idx, "Elevation [m]"]

        # if Mass is on node
        if float(z_Mass) in [float(indx) for indx in NODES.loc[:, "z"]]:
            NODES.loc[idx, "pMass"] = MASSES.loc[idx, "Mass [kg]"]

        # if not, distribution on neibhoring nodes
        else:
            below_idx = NODES[NODES['z'] <= z_Mass].index.min()
            above_idx = NODES[NODES['z'] > z_Mass].index.max()

            dist_below_rel = (z_Mass - NODES.loc[below_idx, "z"])/(NODES.loc[above_idx, "z"] - NODES.loc[below_idx, "z"])
            dist_above_rel = (NODES.loc[above_idx, "z"] - z_Mass)/(NODES.loc[above_idx, "z"] - NODES.loc[below_idx, "z"])

            m_below = MASSES.loc[idx, "Mass [kg]"] * dist_below_rel
            m_above = MASSES.loc[idx, "Mass [kg]"] * dist_above_rel

            NODES.loc[below_idx, "pMass"] += m_below
            NODES.loc[above_idx, "pMass"] += m_above

    calculate_deflection(NODES, (0.75, "deg"), (5, "mm/m"), defl_TP=(0.5, "deg"))

    #Intertia
    GEOMETRY["pInertia"] = 0

    pMass = [0 for i in enumerate(range(N_Nodes))]
    pIntertia = [0 for i in enumerate(range(N_Nodes))]
    pOutOfVertically = [0 for i in enumerate(range(N_Nodes))]

    for node, _ in enumerate(range(N_Nodes, -1, -1)):
        temp = ("os_FeNode{model=ModelName"
                     ", node=" + str(501+node) +
                     ",	x=0"
                     ", y=0"
                     ", z=" + str(GEOMETRY.iloc[node]["Top [m]"]) +
                     ", pMass=" + str(pMass[node]) +
                     ", pInertia" + str(pIntertia[node]) +
                     ", pOutOfVertically" + str(pOutOfVertically[node]) + "}")
        NODES.append(temp)




    with open(txt_path,'w') as file:
        file.write(text)



    return


txt_path = "struct.txt"

export_JBOOST(txt_path)
