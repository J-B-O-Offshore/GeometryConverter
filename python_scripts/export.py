import os.path

import pandas as pd
import excel as ex
from typing import Tuple, Optional

import misc as mc
import numpy as np
import math
import re


def read_lua_values(file_path, keys):
    """
    Extracts specified key-value pairs from a Lua file.

    Parameters:
    ----------
    file_path : str
        The path to the Lua file from which values need to be extracted.

    keys : list of str
        A list of keys (as strings) for which the corresponding values should be extracted from the Lua file.

    Returns:
    -------
    dict
        A dictionary where the keys are the specified keys from the input list, and the values are the corresponding values
        found in the Lua file. The values are converted to their appropriate types (int, float, bool, or str) based on their
        format in the Lua file.
    """

    # Dictionary to store the values
    values_dict = {}

    # Regular expression to match key-value pairs in the Lua file
    pattern = re.compile(r'(\w+)\s*=\s*(.+)')

    with open(file_path, 'r') as file:
        for line in file:
            # Remove comments and strip any leading/trailing whitespace
            line = line.split('--')[0].strip()
            if not line:
                continue

            # Match the key-value pair
            match = pattern.match(line)
            if match:
                key, value = match.groups()

                # Remove trailing comma if present
                value = value.rstrip(',')

                # Check if the key is in the desired keys
                if key in keys:
                    # Attempt to convert to a number or leave as a string
                    try:
                        # Handle numbers and booleans
                        if value.lower() == "true":
                            value = True
                        elif value.lower() == "false":
                            value = False
                        elif "." in value:
                            value = float(value)
                        else:
                            value = int(value)
                    except ValueError:
                        # Keep as string if conversion fails
                        value = value.strip('"').strip("'")

                    values_dict[key] = value

    return values_dict


def write_lua_variables(lua_string, variables):
    """
    Updates Lua variable assignments in a string.

    Args:
        lua_string (str): The Lua source code as a string.
        variables (dict): Dictionary of variables to update, e.g., {'var_name': new_value}.

    Returns:
        str: The modified Lua code as a string.
    """
    lines = lua_string.splitlines()

    for variable_name, new_value in variables.items():
        variable_pattern = re.compile(rf'^(\s*{re.escape(variable_name)}\s*=\s*).*?(\s*,\s*--.*)?$')

        for i, line in enumerate(lines):
            match = variable_pattern.match(line)
            if match:
                indentation = match.group(1)
                rest_of_line = match.group(2) if match.group(2) else ','
                lines[i] = f"{indentation}{new_value}{rest_of_line}"
                break

    return "\n".join(lines)


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
                         member_id=1, create_node_tolerance=0.1, seabed_level=None, waterlevel=0):
    """
    Generates a JBOOST structural input text block based on geometric and mass data for offshore wind turbine structures.

    This function constructs node and element definitions for a structural model used in the JBOOST simulation tool.
    It processes geometry and mass input data, applies deflections, optionally inserts RNA and additional point masses,
    and formats the data into a textual representation for JBOOST input.

    Parameters:
    ----------
    GEOMETRY : pd.DataFrame
        DataFrame containing structural geometry with columns:
        ["Top [m]", "Bottom [m]", "D, top [m]", "D, bottom [m]", "t [mm]", "Affiliation"].

    RNA : pd.DataFrame
        DataFrame with RNA (Rotor-Nacelle Assembly) properties.
        Must include "Offset TT_COG [m]", "Mass [kg]", "Inertia [kg m^2]", and "Identifier".

    defl_MP : float or np.ndarray
        Monopile deflection at the top node (used for calculating out-of-verticality).

    delf_TOWER : float or np.ndarray
        Tower-specific deflection used in deflection calculation.

    MASSES : pd.DataFrame, optional
        Optional DataFrame of additional point masses with columns:
        ["Elevation [m]", "Mass [kg]", "comment"].

    defl_TP : float or np.ndarray, optional
        Transition piece deflection, if different from tower or monopile.

    ModelName : str, default "Struct"
        Name of the model used in JBOOST output.

    EModul : str, default "2.10E+11"
        Young’s modulus (E) of material used in the elements.

    fyk : str, default "355"
        Yield strength of the material in MPa.

    poisson : str, default "0.3"
        Poisson’s ratio of the material.

    dens : str or float, default "7850"
        Density of the structure in kg/m³.

    addMass : float, default 0
        Additional mass per meter applied uniformly to elements.

    member_id : int, default 1
        ID used for grouping or identifying structural members.

    create_node_tolerance : float, default 0.1
        Tolerance used when placing point masses: if no existing node is within this
        tolerance of a mass elevation, a new node is created.

    seabed_level : float, optional
        Elevation of the seabed. All geometry below this level is removed.

    Returns:
    -------
    str
        A formatted string representing the full JBOOST model, including node and element
        definitions with deflections, mass distributions, and structural properties.
    """

    NODES_txt = []
    ELEMENTS_txt = []
    elements = []
    waterlevel = float(waterlevel)
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
    if not waterlevel in NODES["z"].values:
        NODES = add_node(NODES, waterlevel)
        GEOMETRY = mc.add_element(GEOMETRY, waterlevel)

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

    NODES.loc[NODES["z"] == z_RNA, "pMass"] += RNA.loc[0, "Mass [kg]"]
    NODES.loc[NODES["z"] == z_RNA, "pMassName"] = f"RNA {RNA.loc[0, 'Identifier']}"
    NODES.loc[NODES["z"] == z_RNA, "pInertia"] = RNA.loc[0, 'Inertia [kg m^2]']

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


def create_JBOOST_proj(Parameters, marine_growth=None, modelname="struct.lua", runFEModul=True, runFrequencyModul=False, runHindcastValidation=False, wavefile="wave.lua", windfile="wind."):
    """
       Generates a Lua-based project file text for JBOOST simulations based on input parameters and configuration flags.

       The function creates a Lua script for structural simulation using JBOOST. It allows configuration of modules such
       as FEM, frequency domain, and hindcast validation. The function also embeds parameters and optional marine growth
       data into the Lua template.

       Parameters
       ----------
       Parameters : dict
           A dictionary containing configuration values to be injected into the Lua script using placeholder tags.
       marine_growth : pandas.DataFrame, optional
           DataFrame with marine growth layer specifications. Expected columns:
           - 'Top [m]': Top elevation of the layer.
           - 'Bottom [m]': Bottom elevation of the layer.
           - 'Marine Growth [mm]': Thickness of the marine growth layer in millimeters.
       modelname : str, default="struct.lua"
           Filename for the structural model Lua file.
       runFEModul : bool, default=True
           If True, includes a call to run the FEM module (`os_RunFEModul()`).
       runFrequencyModul : bool, default=False
           If True, includes a call to run the frequency domain module (`os_RunFrequencyDomain()`).
       runHindcastValidation : bool, default=False
           If True, includes a call to run hindcast validation (`os_RunHindcastValidation()`).
       wavefile : str, default="wave.lua"
           Filename for the wave input Lua file.
       windfile : str, default="wind."
           Filename for the wind input Lua file.

       Returns
       -------
       str
           A string containing the fully-formed Lua project script for use with JBOOST.

       Raises
       ------
       ValueError
           If both `runFEModul` and `runFrequencyModul` are set to True. Only one simulation mode should be active.

       Notes
       -----
       - Placeholder tokens in the Lua template (e.g., `?modelname`, `?RunFeModul`, etc.) are replaced based on inputs.
       - Marine growth layers, if provided, are inserted via `os_OceanMG{}` calls.
       - Assumes `write_lua_variables()` is defined elsewhere to handle parameter injection.
       """
    if runFEModul and runFrequencyModul:
        raise ValueError("RunFEModul and runFrequencyModul can not both be set to True")
    Proj_text = """\
-- ++++++++++++++++++++++++++++++++++++++++ Model Data +++++++++++++++++++++++++++++++++++++++++++++++++++

model = getdata(?modelname)
model()
wave = (getdata(?wavefile))
wave()
wind = getdata(?wavefile)
wind()

-- ++++++++++++++++++++++++++++++++++++++++ Configurations +++++++++++++++++++++++++++++++++++++++++++++++++++

os_LoadConfigGeneral(
{
    Config_name       = ,
    Result_name       = ,
    Scatter_name      = ,
    FeModel_name      = ,
    Wind_name         = ,
    Ocean_name        = "test",
}
)

-- ++++++++++++++++++++++++++++++++++++++++ FEModul Konfiguration ++++++++++++++++++++++++++++++++++++++++++++++++++++++
os_LoadConfigFEModul(
{
    foundation_superelement  = ,  -- Knotennummer (-1=keine Ersatzmatizen fuer Pfahl, >0 Ersatzsteifigkeit bzw. -massen)

    found_stiff_trans        = ,  -- aus 22A559 EnBW GOA aber 10.4 m pile
    found_stiff_rotat        = ,
    found_stiff_coupl        = ,
    found_mass_trans         = ,
    found_mass_rotat         = ,
    found_mass_coupl         = ,

    shearCorr                = ,  -- Schubkorrekturfaktor (2.0 fuer Kreisquerschnitt und Timoshenko 0.0 fuer Bernoulli)
    res_NumEF                = ,  -- Anzahl der Eigenfrequenzen, die ausgegeben werden
    hydro_add_mass           = ,  -- Added Mass fuer analysen der Strukturdynamik 1 = mit added mass
}
)

?MARINE_GROWTH

os_LoadConfigOcean(
{
    water_density    = ,  -- water density in [kg/m ]
    water_level      = ,  -- wrt LAT in [m] !!! there must be a node @ this height !!!
    seabed_level     = ,  -- wrt LAT in [m]
    growth_density   = ,  -- marine growth density in [kg/m ]
}
)

os_LoadConfigFrequModul(
{
    frequRange            = ,  -- 0 Hz to frequRange in Hz
    frequRangeSolution    = ,  -- no. of frequency components
    maxEFcalc             = ,  -- max number of considered frequencies
    damping_struct        = ,  -- damping ratio structural (material, viscous, soil)
    damping_tower         = ,
    damping_aerodyn       = ,  -- damping ratio aerodynamic (weighted mean value over wind speeds)
    design_life           = ,  -- structural design life in years
    N_ref                 = ,  -- no. of cycles for DEL calculation
    tech_availability     = ,  -- technical availability turbine in percent --> 0.95*30y/31.0y
    -- Note that this coers also the commissioning and decommissioning times
    SN_slope              = ,  -- S-N slope material
    h_refwindspeed        = ,  -- height reference in m of wind speed data in wind wave scatter
    h_hub                 = ,  -- hub height wrt LAT in m
    height_exp            = ,  -- height exponent
    -- height exp was set specifically here so that the average wind speed from the ROUGH Vw Hs scatter is 10.5 m/s at hub height
    TM02_period           = ,  -- [0] if peak periods stated, [1] if zero crossing Tm02 periods stated -> noch nicht eingebaut...FOs
    refineScatter         = ,  -- refinement of scatter data on Tp axis (factor 10 coded)  [0]=off; [1]=on; [2]=on incl. validation plots!
    fullScatter           = ,  -- calc full real scatter FLS for validation purpose
    WindspecDataBase      = ,  -- path on windspec database (*.csv) or "-" in case no wind load is to be considered
    res_Nodes             = 
}
)
-- ----------------------------------------------
-- Exexute program steps
-- ----------------------------------------------

?RunFeModul
?RunFrequencyDomain
?runHindcastValidation

-- Ausgabe der Textdateien
os_WriteResultsText([[Results_JBOOST_Text]])
os_PlotResultsGraph([[Result_JBOOST_Graph]])
"""

    Proj_text = write_lua_variables(Proj_text, Parameters)

    if runFEModul:
        Proj_text = Proj_text.replace("?RunFeModul", "os_RunFEModul()")
    else:
        Proj_text = Proj_text.replace("?RunFeModul", "")

    if runFrequencyModul:
        Proj_text = Proj_text.replace("?RunFrequencyDomain", "os_RunFrequencyDomain()")
    else:
        Proj_text = Proj_text.replace("?RunFrequencyDomain", "")

    if runHindcastValidation:
        Proj_text = Proj_text.replace("?runHindcastValidation", "os_RunHindcastValidation()")
    else:
        Proj_text = Proj_text.replace("?runHindcastValidation", "")

    Proj_text = Proj_text.replace("?modelname", modelname)
    Proj_text = Proj_text.replace("?windfile", windfile)
    Proj_text = Proj_text.replace("?wavefile", wavefile)

    # marine growth
    if marine_growth is not None:
        marine_text = ["os_OceanMG{"+f"id=HD, topMGsection = {row['Top [m]']},   bottomMGsection = {row['Bottom [m]']},    t={row['Marine Growth [mm]']/1000}" + "} -- marine growth layer spec." for _, row in marine_growth.iterrows()]
        marine_text = "\n".join(marine_text)
        Proj_text = Proj_text.replace("?MARINE_GROWTH", marine_text)


    return Proj_text


def export_JBOOST(jboost_path):
    jboost_path = os.path.abspath(jboost_path)
    GEOMETRY = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "WHOLE_STRUCTURE")
    GEOMETRY = GEOMETRY.drop(columns=["Section", "local Section"])
    MASSES = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "ALL_ADDED_MASSES")
    MARINE_GROWTH = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "MARINE_GROWTH")
    PARAMETERS = ex.read_excel_table("GeometrieConverter.xlsm", "ExportStructure", "JBOOST_PARAMETER")
    STRUCTURE_META = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "STRUCTURE_META")
    PROJECT = ex.read_excel_table("GeometrieConverter.xlsm", "ExportStructure", "JBOOST_PROJECT", dtype=str)
    RNA = ex.read_excel_table("GeometrieConverter.xlsm", "StructureOverview", "RNA")
    RNA.dropna(how="all", axis=0, inplace=True)

    if len(RNA) == 0:
        ex.show_message_box("GeometrieConverter.xlsm",
                            f"Please define RNA parameters. Aborting")
        return

    # check Geometry
    sucess_GEOMETRY = mc.sanity_check_structure(GEOMETRY)
    if not sucess_GEOMETRY:
        ex.show_message_box("GeometrieConverter.xlsm",
                            f"Geometry is messed up. Aborting")
        return

    Model_name = PARAMETERS.loc[PARAMETERS["Parameter"] == "ModelName", "Value"].values[0]

    # proj file
    PROJECT = PROJECT.set_index("Project Settings")

    default = PROJECT.loc[:, "default"]
    proj_configs = PROJECT.iloc[:, 2:]

    # itterate trought configs
    for config_name, config_data in proj_configs.items():

        # Fill missing values in config_data with defaults
        config_data = config_data.replace("", np.nan)
        config_data = config_data.fillna(default)

        # filling auto values
        def resolve_auto_value(parameter, config_key, description):
            var = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == parameter, "Value"].values
            if len(var) > 0 and isinstance(var[0], (int, float)):
                config_data[config_key] = var[0]
            else:
                ex.show_message_box(
                    "GeometrieConverter.xlsm",
                    f"Please set {description} in the StructureOverview, as you set {config_key} in {config_name} to 'auto'. Please check exported file or export again with right settings."
                )

        if config_data["water_level"] == 'auto':
            resolve_auto_value("Water level", "water_level", "a water level")

        if config_data["seabed_level"] == 'auto':
            resolve_auto_value("Seabed level", "seabed_level", "a seabed level")

        if config_data["h_hub"] == 'auto':
            resolve_auto_value("Hubheight", "h_hub", "Hubheight")
        if config_data["h_refwindspeed"] == 'auto':
            config_data["h_refwindspeed"] = config_data["h_hub"]

        config_struct = {row: data for row, data in config_data.items()}
        config_struct.pop("runFEModul", None)
        config_struct.pop("runFrequencyModul", None)

        def str_to_bool(s):
            if s == "True":
                return True
            elif s == "False":
                return False
            else:
                raise ValueError(f"Invalid boolean string: {s}")

        runFEModul = str_to_bool(config_data["runFEModul"])
        runFrequencyModul = str_to_bool(config_data["runFrequencyModul"])

        proj_text = create_JBOOST_proj(config_struct,
                                       MARINE_GROWTH,
                                       modelname=Model_name,
                                       runFEModul=runFEModul,
                                       runFrequencyModul=runFrequencyModul,
                                       runHindcastValidation=False,
                                       wavefile="wave.lua",
                                       windfile="wind.")

        struct_text = create_JBOOST_struct(GEOMETRY,
                                           RNA,
                                           (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Value"].values[0],
                                            PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection MP", "Unit"].values[0]),
                                           (PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Value"].values[0],
                                            PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TOWER", "Unit"].values[0]),
                                           MASSES=MASSES,
                                           defl_TP=(PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Value"].values[0],
                                                    PARAMETERS.loc[PARAMETERS["Parameter"] == "deflection TP", "Unit"].values[0]),
                                           ModelName=Model_name,
                                           EModul=PARAMETERS.loc[PARAMETERS["Parameter"] == "EModul", "Value"].values[0],
                                           fyk="355",
                                           poisson="0.3",
                                           dens=PARAMETERS.loc[PARAMETERS["Parameter"] == "dens", "Value"].values[0],
                                           addMass=0,
                                           member_id=1,
                                           create_node_tolerance=PARAMETERS.loc[PARAMETERS["Parameter"] == "Dimensional tolerance for node generating [m]", "Value"].values[0],
                                           seabed_level=STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"].values[0],
                                           waterlevel=config_struct["water_level"]
                                           )

        path_config = os.path.join(jboost_path, config_name)

        if not os.path.exists(path_config):
            os.mkdir(path_config)

        path_proj = os.path.join(path_config, config_name+".lua")
        path_struct = os.path.join(path_config, Model_name+".lua")

        with open(path_proj, 'w') as file:
            file.write(proj_text)
        with open(path_struct, 'w') as file:
            file.write(struct_text)

    ex.show_message_box("GeometrieConverter.xlsm", f"JBOOST Structure {PARAMETERS.loc[PARAMETERS['Parameter'] == 'ModelName', 'Value'].values[0]} saved sucessfully at {jboost_path}")

    return


export_JBOOST(".")