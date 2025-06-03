import xlwings
import pandas as pd
import sqlite3
import excel as ex
import misc as mc
from typing import Tuple

import numpy as np


def export_JBOOST(txt_path):
    def calculate_deflection(NODES: pd.DataFrame, defl_MP: Tuple[float, str], defl_Tower: Tuple[float, str], defl_TP=None):
        """
        Calculate p from vertical deflections.

        Parameters:
        - NODES: DataFrame containing structure nodes (expected columns: z, Affiliation with TOWER or MP and optional TP or GROUTED).
        - defl_MP: Tuple containing (value, unit) unit: "mm/m" or "deg" for MP deflection.
        - defl_Tower: Tuple containing (value, unit) unit: "mm/m" or "deg" forTower deflection.

        Returns:
        - Series with calculated deflection values.
        """
        if defl_MP[1] == "deg":
            defl_MP_deg = defl_MP[0] * np.pi/180
        elif defl_MP[1] == "mm/m":
            defl_MP_deg = np.arctan(0.001 * defl_MP[0]) * np.pi/180
        else:
            raise ValueError("defl_MP[1] has to be deg or mm/m!")

        if defl_Tower[1] == "deg":
            defl_Tower_deg = defl_Tower[0] * np.pi/180
        elif defl_Tower[1] == "mm/m":
            defl_Tower_deg = np.arctan(0.001 * defl_Tower[0]) * np.pi/180
        else:
            raise ValueError("defl_Tower[1] has to be deg or mm/m!")

        if defl_TP is not None:
            if defl_TP[1] == "deg":
                defl_TP_deg = defl_TP[0] * np.pi/180
            elif defl_TP[1] == "mm/m":
                defl_TP_deg = np.arctan(0.001 * defl_TP[0]) * np.pi/180
            else:
                raise ValueError("defl_Tower[1] has to be deg or mm/m!")

            line_TP = np.sin(defl_TP_deg) * (NODES["z"] - NODES.loc[NODES.index[-1], "z"])

        else:
            if "TP" in NODES["Affiliation"].unique():
                raise ValueError("defl_TP has to be defined, as TP is in NODES Affiliation column")

        line_MP = np.sin(defl_MP_deg) * (NODES["z"] - NODES.loc[NODES.index[-1], "z"])

        line_TOWER = np.sin(defl_Tower_deg) * (NODES["z"] - NODES.loc[NODES.index[-1], "z"])

        # construct line
        NODES["DEFL"] = 0.0
        NODES.loc[NODES["Affiliation"] == "MP", "DEFL"] = line_MP.loc[NODES["Affiliation"] == "MP"]

        if defl_TP is not None:
            line_TP_new = line_TP - line_TP.loc[NODES["Affiliation"] == "MP"].iloc[0] + line_MP.loc[NODES["Affiliation"] == "MP"].iloc[0]
            NODES.loc[NODES["Affiliation"] == "TP", "DEFL"] = line_TP_new.loc[NODES["Affiliation"] == "TP"]

            line_TOWER_new = line_TOWER - line_TOWER.loc[NODES["Affiliation"] == "TP"].iloc[0] + NODES.loc[NODES["Affiliation"] == "TP", "DEFL"].iloc[0]
            NODES.loc[NODES["Affiliation"] == "TOWER", "DEFL"] = line_TOWER_new.loc[NODES["Affiliation"] == "TOWER"]

        else:
            line_TOWER_new = line_TOWER - line_TOWER.loc[NODES["Affiliation"] == "MP"].iloc[0] + line_MP.loc[NODES["Affiliation"] == "MP"].iloc[0]
            NODES.loc[NODES["Affiliation"] == "TOWER", "DEFL"] = line_TOWER_new.loc[NODES["Affiliation"] == "TP"]

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
