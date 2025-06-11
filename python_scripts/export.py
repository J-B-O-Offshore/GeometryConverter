import os.path

import pandas as pd
import excel as ex
from typing import Tuple, Optional

import misc as mc
import numpy as np
import math

def calculate_deflection(
        NODES: pd.DataFrame,
        defl_MP: Tuple[float, str],
        defl_TP: Tuple[float, str],
        defl_Tower: Tuple[float, str],

) -> pd.Series:
    """
    Calculate deflection values based on given tilt angles. defl_Tower is relative to defl TP, defl_MP and defl_TP is absloute

    Parameters:
    - NODES: DataFrame with columns ["z", "Affiliation"]
    - defl_MP, defl_Tower, defl_TP: Tuples (value, unit), where unit is "deg" or "mm/m"

    Returns:
    - pd.Series of deflection values
    """

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

    affiliations = NODES["Affiliation"]
    z = NODES["z"]
    base_z = z.iloc[-1]

    # Convert all given angles to radians
    angle_MP = _convert_to_rad(*defl_MP)
    angle_Tower = _convert_to_rad(*defl_Tower)
    angle_TP = _convert_to_rad(*defl_TP) if defl_TP else None

    angle_Tower = angle_TP + angle_Tower

    # Compute initial deflection lines
    line_MP = _compute_line(z, angle_MP, base_z)
    line_Tower = _compute_line(z, angle_Tower, base_z)
    line_TP = _compute_line(z, angle_TP, base_z) if angle_TP else None

    # Initialize DEFL column
    NODES["DEFL"] = 0.0

    # Assign MP deflection
    mask_MP = affiliations == "MP"
    NODES.loc[mask_MP, "DEFL"] = pd.Series(line_MP[mask_MP], dtype="float64")

    mask_TP = affiliations == "TP"
    TP_offset = line_MP[mask_MP].iloc[0] - line_TP[mask_MP].iloc[0]
    NODES.loc[mask_TP, "DEFL"] = pd.Series(line_MP[mask_MP], dtype="float64") + TP_offset

    mask_TOWER = affiliations == "TOWER"
    TOWER_offset = NODES.loc[mask_TP, "DEFL"].iloc[0] - line_Tower[mask_TP].iloc[0]
    NODES.loc[mask_TOWER, "DEFL"] = pd.Series(line_MP[mask_MP], dtype="float64") + TOWER_offset


    return NODES["DEFL"]


def add_node(df, z_new):
    """
    Add a new node at the specified height `z_new` to the DataFrame `df`.

    Parameters:
    - df (pd.DataFrame): Original node DataFrame.
    - z_new (float): Height at which to add the new node.

    Returns:
    - pd.DataFrame: Updated DataFrame with new node added and reindexed.
    """
    # Create a new node row
    new_node = pd.DataFrame([{
        'z': z_new,
        'node': 0,  # placeholder
        'pInertia': 0,
        'pMass': 0.0,
        'Affiliation': 'ADDED_FOR_MASS'
    }])

    # Append and sort by height (descending)
    df_updated = pd.concat([df, new_node], ignore_index=True)
    df_updated = df_updated.sort_values(by='z', ascending=False).reset_index(drop=True)

    # Reassign node numbers: highest z gets highest node number
    max_node = df_updated['node'].max() + 1  # ensure uniqueness if original node numbers are reused
    df_updated['node'] = list(range(max_node, max_node - len(df_updated), -1))

    return df_updated


def interpolate_with_neighbors(data):
    """
    Linearly interpolates internal None/NaN values with neighbors and extrapolates
    at the beginning and end using the first/last two known values.

    Parameters:
        data (list of float, None, or NaN): The input list.

    Returns:
        list of float or None: List with interpolated and extrapolated values.
    """
    def is_missing(x):
        return x is None or (isinstance(x, float) and math.isnan(x))

    result = data[:]

    # --- First: Interpolate values with two known neighbors ---
    i = 0
    while i < len(result):
        if is_missing(result[i]):
            start = i
            while i < len(result) and is_missing(result[i]):
                i += 1
            end = i
            if 0 < start and end < len(result) and not is_missing(result[start - 1]) and not is_missing(result[end]):
                left = result[start - 1]
                right = result[end]
                n = end - start + 1
                for j in range(start, end):
                    frac = (j - start + 1) / n
                    result[j] = left + frac * (right - left)
        else:
            i += 1

    # --- Then: Extrapolate missing values at the start ---
    first_known_idx = next((i for i, x in enumerate(result) if not is_missing(x)), None)
    if first_known_idx is not None and first_known_idx >= 2:
        val1 = result[first_known_idx]
        val2 = result[first_known_idx + 1]
        for i in range(first_known_idx - 1, -1, -1):
            result[i] = val1 - (first_known_idx - i) * (val2 - val1)

    # --- Extrapolate missing values at the end ---
    last_known_idx = next((i for i in reversed(range(len(result))) if not is_missing(result[i])), None)
    if last_known_idx is not None and last_known_idx <= len(result) - 3:
        val1 = result[last_known_idx]
        val2 = result[last_known_idx - 1]
        for i in range(last_known_idx + 1, len(result)):
            result[i] = val1 + (i - last_known_idx) * (val1 - val2)

    return result

def create_JBOOST_struct(GEOMETRY, RNA, defl_MP, delf_TOWER, MASSES=None, defl_TP=None, ModelName="Struct", EModul="2.10E+11", fyk="355", poisson="0.3", dens="7850", addMass=0,
                         member_id=1, create_node_tolerance=0.1, seabed_level=None):
    NODES_txt = []
    ELEMENTS_txt = []
    elements = []

    if seabed_level is not None:
        GEOMETRY = mc.add_element(GEOMETRY, seabed_level)
    GEOMETRY = GEOMETRY.loc[GEOMETRY["Bottom [m]"] >= seabed_level]
    # Nodes
    N_Nodes = len(GEOMETRY) + 1
    NODES = pd.DataFrame({
        "z": pd.concat([
            GEOMETRY["Top [m]"],
            pd.Series([GEOMETRY["Bottom [m]"].iloc[-1]])
        ], ignore_index=True)
    })
    nodes = [501 + i for i in range(N_Nodes)]
    nodes.reverse()
    NODES["node"] = nodes
    NODES["pInertia"] = 0
    NODES["pMass"] = 0.0
    NODES["Affiliation"] = "NOT DEFINDED"
    NODES["pMassName"] = ""
    NODES.loc[NODES.index[:-1], "Affiliation"] = GEOMETRY.iloc[:]["Affiliation"]
    NODES.loc[NODES.index[-1], "Affiliation"] = GEOMETRY.iloc[-1]["Affiliation"]

    # calculate deflection
    calculate_deflection(NODES, defl_MP, defl_TP, delf_TOWER)

    # insert Node at WL
    if not 0 in NODES["z"].values:
        NODES = add_node(NODES, 0)
        GEOMETRY = mc.add_element(GEOMETRY, 0)

    # distribute Masses, create new Node if in between
    if MASSES is not None:
        # distribute Masses on Nodes
        for idx in MASSES.index:
            z_Mass = MASSES.loc[idx, "Elevation [m]"]

            # if Mass is near node
            in_tolerance = np.where(np.abs(NODES["z"].values - z_Mass) <= create_node_tolerance)[0]

            if len(in_tolerance) > 0:
                NODES.loc[in_tolerance[0], "pMass"] += MASSES.loc[idx, "Mass [kg]"]
                NODES.loc[in_tolerance[0], "pMassName"] += MASSES.loc[idx, "comment"] + " "
            # if not, create new nodes
            else:
                if z_Mass >= GEOMETRY.loc[:, "Bottom [m]"].values[-1]:
                    NODES = add_node(NODES, z_Mass)
                    NODES.loc[NODES["z"] == z_Mass, "pMass"] += MASSES.loc[idx, "Mass [kg]"]
                    NODES.loc[NODES["z"] == z_Mass, "pMassName"] = MASSES.loc[idx, "comment"]
                    GEOMETRY = mc.add_element(GEOMETRY, z_Mass)
                else:
                    print(f"Warning! Mass '{MASSES.loc[idx, 'comment']}' not added, it is below the seabed level!")

    # add RNA
    MP_top = NODES.iloc[0, :].loc["z"]
    D_MP_top = GEOMETRY.iloc[0, :].loc["D, top [m]"]
    t_MP_top = GEOMETRY.iloc[0, :].loc["t [mm]"]

    z_RNA = MP_top + RNA.loc[0, "Offset TT_COG [m]"]
    NODES = add_node(NODES, z_RNA)
    GEOMETRY = pd.concat([pd.DataFrame({'Affiliation': 'TOWER', 'Top [m]': z_RNA, 'Bottom [m]': MP_top, 'D, top [m]': D_MP_top, 'D, bottom [m]': D_MP_top,
                                     't [mm]': t_MP_top}, index=[-1]), GEOMETRY], ignore_index=True, axis=0)

    NODES.loc[NODES["z"] == z_RNA, "pMass"] += RNA.loc[0, "Mass [t]"]*1000
    NODES.loc[NODES["z"] == z_RNA, "pMassName"] = f"RNA {RNA.loc[0, 'Identifier']}"
    NODES.loc[NODES["z"] == z_RNA, "pInertia"] = RNA.loc[0, 'Inertia [kg·m²]']

    # fill deflection values for newly inserted Nodes
    NODES.loc[:, "DEFL"] = interpolate_with_neighbors(NODES.loc[:, "DEFL"].values)

    # Revert Node order
    NODES = NODES.iloc[::-1].reset_index(drop=True)

    # construct Node lines
    for idx, node in NODES.iterrows():
        temp = ("os_FeNode{model=ModelName" "\t" +
                ",node=" + str(int(node["node"])) + "\t" +
                ",x=0" "\t" +
                ",y=0" "\t" +
                ",z=" + f"{round(node['z'], 2):05.2f}" + "\t" +
                ",pMass=" + f"{node['pMass']:06.0f}" + "\t" +
                ",pInertia=" + str(node["pInertia"]) + "\t" +
                ",pOutOfVertically=" + f"{round(node['DEFL'], 3):06.3f}" + "}")

        # write comment about pointmasses
        if node["pMassName"] != "":
            if node["Affiliation"] == "ADDED_FOR_MASS":
                temp += f"--added node for pointmass '{node['pMassName']}'"
            else:
                temp += f"--placed pointmasse(s) '{node['pMassName']}'"
        NODES_txt.append(temp)

    # Elements
    for i in range(len(NODES) - 1):
        startnode = NODES.loc[i, "node"]
        endnode = NODES.loc[i + 1, "node"]
        diameter = (GEOMETRY.loc[i, "D, top [m]"] + GEOMETRY.loc[i, "D, bottom [m]"]) / 2
        t_wall = GEOMETRY.loc[i, "t [mm]"] / 1000

        elements.append({"elem_id": i + 1, "startnode": startnode, "endnode": endnode, "diameter": diameter, "t_wall": t_wall, "dens": dens})

    ELEM = pd.DataFrame(elements)
    ELEM.loc[ELEM.index[-1], "dens"] = 1.0

    for idx, elem in ELEM.iterrows():
        temp = ("os_FeElem{model=ModelName\t"
                ",elem_id=" + f"{elem['elem_id']:03.0f}" + "\t" +
                ",startnode=" + f"{elem['startnode']:03.0f}" + "\t" +
                ",endnode=" + f"{elem['endnode']:03.0f}" + "\t" +
                ",diameter=" + f"{round(elem['diameter'], 2):02.2f}" + "\t" +
                ",t_wall=" + f"{round(elem['t_wall'], 3):.3f}" + "\t" +
                ",EModul=" + str(EModul) + "\t" +
                ",fky=" + str(fyk) + "\t" +
                ",poisson=" + str(poisson) + "\t" +
                ",dens=" + str(elem['dens']) + "\t" +
                ",addMass=" + str(addMass) + "\t" +
                ",member_id=" + str(member_id) + "\t" +
                "}")
        ELEMENTS_txt.append(temp)

    text = ("--Input for JBOOST generated by Excel\n" +
            "--    Definition Modelname\n"
            'local	ModelName="' + ModelName + '"\n' +
            "--    Definition der FE-Knoten\n" +
            "\n".join(NODES_txt) + "\n\n" +
            "--Definition der FE-Elemente- Zusatzmassen in kg / m" + "\n" +
            "\n".join(ELEMENTS_txt) + "\n"
            )
    return text


def export_JBOOST(lua_path):
    lua_path = os.path.abspath(lua_path)
    GEOMETRY = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "WHOLE_STRUCTURE")
    GEOMETRY = GEOMETRY.drop(columns=["Section", "local Section"])
    MASSES = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "ALL_ADDED_MASSES")
    MARINE_GROWTH = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "MARINE_GROWTH")
    PARAMETERS = ex.read_excel_table("GeometrieConverter.xlsm", "ExportStructure", "JBOOST_PARAMETER")
    STRUCTURE_META = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "STRUCTURE_META")
    RNA = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "RNA")

    # check Geometry
    sucess_GEOMETRY = mc.sanity_check_structure(GEOMETRY)
    if not sucess_GEOMETRY:
        return

    text = create_JBOOST_struct(GEOMETRY,
                                RNA,
                                (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Value"].values[0], PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Unit"].values[0]),
                                (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Value"].values[0], PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Unit"].values[0]),
                                MASSES=MASSES,
                                defl_TP=(PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Value"].values[0], PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Unit"].values[0]),
                                ModelName=PARAMETERS.loc[PARAMETERS["Parameter"] == "ModelName", "Value"].values[0],
                                EModul=PARAMETERS.loc[PARAMETERS["Parameter"] == "EModul", "Value"].values[0],
                                fyk="355",
                                poisson="0.3",
                                dens=PARAMETERS.loc[PARAMETERS["Parameter"] == "dens", "Value"].values[0],
                                addMass=0,
                                member_id=1,
                                create_node_tolerance=PARAMETERS.loc[PARAMETERS["Parameter"] == "Dimensional tolerance for node generating [m]", "Value"].values[0],
                                seabed_level=STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"].values[0])

    lua_path = os.path.join(lua_path, PARAMETERS.loc[PARAMETERS["Parameter"] == "ModelName", "Value"].values[0] + ".lua")
    with open(lua_path, 'w') as file:
        file.write(text)

    ex.show_message_box("GeometrieConverter.xlsm", f"JBOOST Structure {PARAMETERS.loc[PARAMETERS['Parameter'] == 'ModelName', 'Value'].values[0]} saved sucessfully at {lua_path}")

    return

export_JBOOST(".")
