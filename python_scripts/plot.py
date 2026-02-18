import os
from collections import defaultdict

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import Rectangle, Arc
from matplotlib.ticker import FixedLocator, FuncFormatter

import excel as ex
import misc as mc


def get_JBO_colors(n, fixed_colors=None):
    """
    Generate a list of colors for plotting with JBO corporate colors as default.

    Parameters
    ----------
    n : int
        Number of colors needed.
    fixed_colors : list of tuple or None
        List of fixed RGB tuples (0–1 range).
        If None, defaults to JBO corporate colors:

        - Yellow    : (242, 184, 30)  → #F2B81E
        - LightBlue : (130, 182, 221) → #82B6DD
        - DarkBlue  : (34, 76, 130)   → #224C82
        - Teal      : (0, 143, 133)   → #008F85
        - Gray      : (142, 135, 120) → #8E8778

    Returns
    -------
    list of tuple
        List of RGB colors (length = n).
    """
    if fixed_colors is None:
        fixed_colors = [
            (242 / 255, 184 / 255, 30 / 255),  # Yellow
            (130 / 255, 182 / 255, 221 / 255),  # Light Blue
            (34 / 255, 76 / 255, 130 / 255),  # Dark Blue
            (0 / 255, 143 / 255, 133 / 255),  # Teal
            (142 / 255, 135 / 255, 120 / 255),  # Gray
        ]

    # Get Matplotlib’s default color cycle
    default_cycle = plt.rcParams['axes.prop_cycle'].by_key()['color']

    colors = []

    # Add fixed colors first
    for i in range(min(n, len(fixed_colors))):
        colors.append(fixed_colors[i])

    # If more colors are needed, continue with Matplotlib cycle
    if n > len(fixed_colors):
        cycle_len = len(default_cycle)
        for i in range(n - len(fixed_colors)):
            colors.append(default_cycle[i % cycle_len])

    return colors


def plot_Structure(Structure, Added_Masses, waterdepth=None, height_ref="", waterlevel=0, show_section_numbers=True):
    fig, ax = plt.subplots(1, 3, figsize=[22, 6])

    # Check if Structure is empty
    structure_empty = Structure.empty

    if not structure_empty:
        top_structure = Structure["Top [m]"].max()
        bottom_structure = Structure["Bottom [m]"].min()
        left = -max(Structure["D, top [m]"].max(), Structure["D, bottom [m]"].max()) / 2
        right = -left
    else:
        top_structure = -np.inf
        bottom_structure = np.inf
        left, right = -1.0, 1.0  # Default limits

    # Use Added_Masses if present to help set limits
    if not Added_Masses.empty:
        top_masses = Added_Masses["Top [m]"].max()
        bottom_masses = Added_Masses["Bottom [m]"].min()
        top_all = max(top_structure, top_masses)
        bottom_all = min(bottom_structure, bottom_masses)
    else:
        top_all = top_structure if top_structure != -np.inf else 10
        bottom_all = bottom_structure if bottom_structure != np.inf else 0

    for ax_curr in ax:
        ax_curr.set_ylim(bottom_all - 0.05 * (top_all - bottom_all), top_all + 0.05 * (top_all - bottom_all))
        ax_curr.grid(True, linestyle='--')

    # Plot 1: Structure overview
    axis = ax[0]
    if not structure_empty:
        plot_cans(Structure, axis, show_section_numbers=show_section_numbers, set_lims=False, color="k")

    if waterlevel is not None:
        axis.axhline(waterlevel, color="blue", linestyle="--")
    axis.axvline(0, color="grey", linestyle="--", linewidth=1)
    if waterdepth is not None:
        axis.axhline(waterdepth, color="brown", linestyle="-", linewidth=2)

    ticks = axis.get_xticks()
    axis.set_xticks(ticks[ticks >= 0])
    axis.set_xticklabels([f"{2 * t:.0f}" for t in ticks if t >= 0])
    axis.set_xlim(left - 0.4 * abs(right), right + 0.4 * abs(right))
    axis.set_xlabel("Diameter [m]")
    axis.set_ylabel(f"z [m{height_ref}]")
    axis.set_title("Overview and Diameter")

    # Plot 2: Added masses
    axis = ax[1]
    if not structure_empty:
        weight = mc.calc_weight(
            7850,
            Structure["t [mm]"],
            Structure["Top [m]"],
            Structure["Bottom [m]"],
            Structure["D, top [m]"],
            Structure["D, bottom [m]"]
        )
        grey = [0.8, 0.8, 0.8]
        plot_cans(Structure, axis, show_section_numbers=False, color=grey, set_lims=False)

    label_entries = []
    placed_z = defaultdict(int)
    tol = 0.050 * (top_all - bottom_all)
    x_step = 0.04 * (right - left)

    for _, mass in Added_Masses.iterrows():
        name = mass['Name']
        top = mass['Top [m]']
        bottom = mass['Bottom [m]']
        mass_val = mass['Mass [kg]']

        is_point = top == bottom
        z = top if is_point else (top + bottom) / 2

        z_key = round(z / tol) * tol
        index = placed_z[z_key]
        x_offset = index * x_step
        placed_z[z_key] += 1

        color = 'red' if is_point else 'blue'
        x_pos = x_offset

        if is_point:
            axis.plot(x_pos, z, 'o', color=color, markersize=8)
        else:
            axis.plot([x_pos, x_pos], [bottom, top], color=color, linewidth=6, alpha=0.6)

        label_entries.append({
            'y': z,
            'label': f"{name} ({mass_val:.0f} kg)",
            'color': color,
            'x_start': x_pos
        })

    # Labeling logic
    MAX_LABELS = 25
    if len(label_entries) <= MAX_LABELS:
        label_entries.sort(key=lambda x: -x['y'])
        occupied_y = []
        min_dy = 0.03 * (top_all - bottom_all)
        x_text = 0.6 * abs(right)

        for entry in label_entries:
            target_y = entry['y']
            while any(abs(target_y - oy) < min_dy for oy in occupied_y):
                target_y -= min_dy
            occupied_y.append(target_y)

            axis.plot([entry['x_start'], x_text - 0.05], [entry['y'], target_y], '-', color=entry['color'])
            axis.text(x_text, target_y, entry['label'], ha='left', va='center', fontsize=9, color=entry['color'])
    else:
        axis.text(0.8 * abs(right), 0.5 * (top_all + bottom_all), "Too many masses to label", ha='center', va='center', fontsize=13, color='red')

    axis.axvline(0, color="grey", linestyle="--", linewidth=1)
    axis.set_xlim(-0.1 * abs(right), 2 * abs(right))
    axis.set_title("Masses\n(horizontal displacement only for distinction, all lie on center line)")
    axis.set_xticks([])
    axis.set_xticklabels([])

    # Plot 3: Slope, Wall Thickness, D/t Ratio
    axis = ax[2]
    axis2 = axis.twiny()
    if not structure_empty:
        z_nodes = list(Structure["Top [m]"]) + [Structure["Bottom [m]"].values[-1]]
        t = list(Structure["t [mm]"] / 1000)
        slope = -(Structure["D, top [m]"] - Structure["D, bottom [m]"]) / (
                Structure["Top [m]"] - Structure["Bottom [m]"]
        )
        D_nodes = list(Structure["D, top [m]"]) + [Structure["D, bottom [m]"].values[-1]]

        # D/t
        z_d_by_t = []
        d_by_t = []
        for _, row in Structure.iterrows():
            t_curr = row["t [mm]"]
            z_d_by_t += [row["Top [m]"], row["Bottom [m]"]]
            d_by_t += [row["D, top [m]"] / t_curr, row["D, bottom [m]"] / t_curr]

        slope_steps = [np.nan if s == 0 else s for s in slope]

        # upper axis (D/t)
        axis2.plot(d_by_t, z_d_by_t, label="D/t", color="C0")
        axis2.xaxis.set_label_position("top")
        axis2.xaxis.tick_top()
        axis2.set_xlabel("D/t [-]")

        # make upper axis blue (no grid)
        axis2.tick_params(axis="x", colors="C0")
        axis2.xaxis.label.set_color("C0")
        axis2.spines["top"].set_color("C0")

        # lower axis (slope + t)
        line1 = axis.stairs(
            slope_steps,
            z_nodes,
            orientation="horizontal",
            label="slope (where not 0 degree)",
            color="C1",
            baseline=None,
        )
        line2 = axis.stairs(
            t, z_nodes, label="t [mm]", orientation="horizontal", color="C2", baseline=None
        )

        lines = [line1, line2]
        labels = [line.get_label() for line in lines]
        axis.legend(lines, labels, loc="lower left")

    axis.set_title("Slope, Wall Thickness, D/t Ratio")
    return fig


def plot_cans(Structure, axis, show_section_numbers=False, set_lims=True, **plot_kwargs):
    """
    Plots a structural representation of cylindrical segments ("cans") on a given Matplotlib axis.

    Parameters
    ----------
    Structure : pandas.DataFrame
        DataFrame describing the structure, with one row per cylindrical segment.
        Required columns:
        - "Top [m]": top elevation of each segment.
        - "Bottom [m]": bottom elevation of each segment.
        - "D, top [m]": diameter at the top of the segment.
        - "D, bottom [m]": diameter at the bottom of the segment.
        - "Section": (optional) section number used for labeling (required if `show_section_numbers=True`).

    axis : matplotlib.axes.Axes
        Matplotlib axis on which the cans will be drawn.

    show_section_numbers : bool, optional
        If True, displays section numbers at the center of each can.
        Requires the 'Section' column in `Structure`.

    set_lims : bool, optional
        If True, automatically sets the axis limits based on the geometry of the structure.

    **plot_kwargs : dict, optional
        Additional keyword arguments passed to `axis.plot()` for customizing line style, color, etc.

    Returns
    -------
    None
        The function modifies the provided `axis` in place.
    """
    if set_lims:
        top_all = Structure["Top [m]"].max()
        bottom_all = Structure["Bottom [m]"].min()
        left = -max(Structure["D, top [m]"].max(), Structure["D, bottom [m]"].max()) / 2
        right = -left
        axis.set_ylim(bottom_all - 0.05 * (top_all - bottom_all), top_all + 0.05 * (top_all - bottom_all))
        axis.set_xlim(left - 0.05 * (right - left), right + 0.05 * (right - left))

    for _, can in Structure.iterrows():
        up_left = (-can["D, top [m]"] / 2, can["Top [m]"])
        up_right = (can["D, top [m]"] / 2, can["Top [m]"])
        down_left = (-can["D, bottom [m]"] / 2, can["Bottom [m]"])
        down_right = (can["D, bottom [m]"] / 2, can["Bottom [m]"])

        axis.plot([down_left[0], up_left[0]], [down_left[1], up_left[1]], **plot_kwargs)
        axis.plot([down_right[0], up_right[0]], [down_right[1], up_right[1]], **plot_kwargs)
        axis.plot([down_left[0], down_right[0]], [down_left[1], down_right[1]], **plot_kwargs)
        axis.plot([up_left[0], up_right[0]], [up_left[1], up_right[1]], **plot_kwargs)

        if show_section_numbers:
            axis.text(0, (down_left[1] + up_left[1]) / 2, int(can["Section"]),
                      fontsize=8, ha='center', va='center', fontweight='bold')


def plot_Assambly(WHOLE_STRUCTURE, SKIRT=None, seabed=None, waterlevel=0, height_ref=None, RNA=None):
    """
    Plot a vertical assembly of a wind turbine or similar structure, including monopile (MP), transition piece (TP),
    tower sections, optional skirt, seabed, water level, and optionally RNA (Rotor-Nacelle Assembly) components.

    Parameters
    ----------
    WHOLE_STRUCTURE : pandas.DataFrame
        DataFrame containing the structure sections with at least the following columns:
        - 'Affiliation': section type ('MP', 'TP', 'TOWER')
        - 'Top [m]': top elevation of the section
        - 'Bottom [m]': bottom elevation of the section
        - 'D, top [m]': diameter at the top of the section
    SKIRT : pandas.DataFrame, optional
        DataFrame of skirt elements, plotted in red if provided. Default is None.
    seabed : float, optional
        Elevation of the seabed. If provided, a solid brown line is drawn. Default is None.
    waterlevel : float, optional
        Elevation of the water level. Drawn as a dashed blue line. Default is 0.
    height_ref : str, optional
        Text to append to the y-axis label to indicate reference (e.g., "(LAT)"). Default is None.
    RNA : dict, optional
        Dictionary defining the rotor-nacelle assembly (RNA) to be plotted at the top of the structure. Expected keys:
        - 'dz_com': relative height of the center of mass above top of structure
        - 'dz_hub': relative height of the hub above top of structure
        - 'diameter': rotor diameter
        - 'info': text label for RNA
        - 'color': optional color for RNA elements (default: 'darkgreen')

    Returns
    -------
    matplotlib.figure.Figure
        The figure object containing the plotted assembly.

    Notes
    -----
    - Sections are colored: MP (black), TP (blue), TOWER (grey), SKIRT (red, semi-transparent).
    - Boundaries between sections are marked with striped horizontal lines.
    - Major and minor grids are added for clarity, and axes are symmetrically scaled.
    - If RNA is provided, a secondary right y-axis is created to show relative heights to the RNA interface,
      including the center of mass, hub, rotor diameter, and label annotation.
    - X-axis shows diameter (2 × radius) for positive side only; y-axis is elevation in meters.
    """

    fig, axis = plt.subplots(1, 1, figsize=[8, 27])

    # -------------------------------------------------
    # Split structure
    # -------------------------------------------------
    MP = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE["Affiliation"] == "MP", :]
    TP = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE["Affiliation"] == "TP", :]
    TOWER = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE["Affiliation"] == "TOWER", :]

    # -------------------------------------------------
    # Plot structure
    # -------------------------------------------------
    if len(MP) != 0:
        plot_cans(MP, axis, show_section_numbers=False, color="k", set_lims=False)
    if len(TP) != 0:
        plot_cans(TP, axis, show_section_numbers=False, color="blue", set_lims=False)
    if len(TOWER) != 0:
        plot_cans(TOWER, axis, show_section_numbers=False, color="grey", set_lims=False)
    if SKIRT is not None:
        plot_cans(SKIRT, axis, show_section_numbers=False,
                  color="red", set_lims=False, alpha=0.8)

    # -------------------------------------------------
    # Axis labels
    # -------------------------------------------------
    axis.set_xlabel("Diameter in [m]")
    if height_ref is not None:
        axis.set_ylabel(f"z in m{height_ref}")
    else:
        axis.set_ylabel(f"z in m")
    # -------------------------------------------------
    # Reference lines
    # -------------------------------------------------
    axis.axvline(0, color="grey", linestyle="--", linewidth=1)

    if waterlevel is not None:
        axis.axhline(waterlevel, color="blue", linestyle="--", linewidth=1.2)

    if seabed is not None:
        axis.axhline(seabed, color="brown", linewidth=2)

    # -------------------------------------------------
    # Detect MP / TP / TOWER boundaries
    # -------------------------------------------------
    boundaries = []
    for df in (MP, TP, TOWER):
        if len(df) > 0:
            boundaries.append(df["Top [m]"].max())
            boundaries.append(df["Bottom [m]"].min())

    # -------------------------------------------------
    # Grid
    # -------------------------------------------------
    axis.grid(True, which="minor", axis="both", linewidth=0.3, alpha=0.5)
    axis.grid(True, which="major", axis="both", linewidth=0.9, alpha=0.8)
    # -------------------------------------------------
    # Detect structure top
    # -------------------------------------------------
    z_top = WHOLE_STRUCTURE["Top [m]"].max()
    D_top = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE.index[0], "D, top [m]"].max()

    # -------------------------------------------------
    # RNA plotting (optional)
    # -------------------------------------------------
    if RNA is not None:
        # Expected dict keys: dz_com, dz_hub, diameter, color (optional)
        RNA_COLOR = RNA.get("color", "darkgreen")
        com_dz = RNA["dz_com"]
        hub_dz = RNA["dz_hub"]
        rotor_diameter = RNA["diameter"]
        text_str = RNA["info"]
        # Right axis
        rax = axis.twinx()
        rax.tick_params(axis="y", colors=RNA_COLOR)
        rax.spines["right"].set_color(RNA_COLOR)
        rax.set_ylabel("Relative height to RNA interface [m]", color=RNA_COLOR)

        # RNA positions
        z_com = z_top + com_dz
        z_hub = z_top + hub_dz
        boundaries.append(z_com)
        boundaries.append(z_hub)
        boundaries.append(z_top + hub_dz - rotor_diameter / 2)

        # Function to map actual height to top-zero axis

        # Right axis ticks
        rax_ticks = [0, com_dz, hub_dz]
        rax.set_yticks(rax_ticks)
        rax.set_yticklabels([f"{z:.2f}" for z in rax_ticks])

        # RNA box (grey)
        box_w = D_top * 1.4
        box_h = max(com_dz, hub_dz) * 2
        axis.add_patch(
            Rectangle(
                (-box_w / 2, z_top),
                box_w,
                box_h,
                facecolor="lightgrey",
                edgecolor=RNA_COLOR,
                linewidth=1.5,
                zorder=6,
                linestyle="--"
            )
        )

        # Center of mass
        axis.scatter(
            0, z_com,
            s=160,
            color=RNA_COLOR,
            zorder=7,
            edgecolors="black",
        )

        # Rotor center (striped cross)
        cross = rotor_diameter * 0.03
        axis.plot(
            [-cross, cross], [z_hub, z_hub],
            color="red", linestyle=(0, (4, 3)), linewidth=1.8, zorder=7
        )
        axis.plot(
            [0, 0], [z_hub - cross, z_hub + cross],
            color="red", linestyle=(0, (4, 3)), linewidth=1.8, zorder=7
        )

        # Rotor diameter (partial circle)
        axis.add_patch(
            Arc(
                (0, z_hub),
                rotor_diameter,
                rotor_diameter,
                theta1=270 - 20 / rotor_diameter * 180 / np.pi,
                theta2=270 + 20 / rotor_diameter * 180 / np.pi,
                color="red",
                linewidth=2.0,
                zorder=6,
                linestyle="--"
            )
        )
        # Angle in degrees
        angle_deg = 8 / rotor_diameter * 180 / np.pi
        angle_rad = np.deg2rad(angle_deg)

        # Original vector from hub to edge (pointing straight down)
        dx = 0
        dy = -rotor_diameter / 2

        # Rotate vector
        dx_rot = dx * np.cos(angle_rad) - dy * np.sin(angle_rad)
        dy_rot = dx * np.sin(angle_rad) + dy * np.cos(angle_rad)

        # Draw rotated arrow
        rax.annotate(
            "",  # No text
            xy=(dx_rot, hub_dz + dy_rot),  # Arrow tip (rotated)
            xytext=(0, hub_dz),  # Arrow start (hub center)
            arrowprops=dict(arrowstyle="->", color="red", linewidth=1.8),
            zorder=8
        )

        # Label the diameter (horizontal text)
        label_x = dx_rot / 2
        label_y = hub_dz + dy_rot / 2
        rax.text(
            label_x, label_y,
            f"D = {rotor_diameter:.2f} m",
            color="red",
            fontsize=15,
            ha="center",
            va="bottom",
            zorder=9,
            fontweight="bold"
        )
        rax.text(
            0,  # x-position (centered)
            box_h,  # y-position (just above the grey box)
            text_str,
            fontsize=10,
            ha="center",
            va="bottom",
            zorder=10,
            fontweight="bold",
        )
    boundaries = sorted(set(boundaries))

    # -------------------------------------------------
    # Engineering grid + ticks
    # -------------------------------------------------
    axis.set_axisbelow(True)

    ymin, ymax = axis.get_ylim()
    xmin, xmax = axis.get_xlim()

    # ---------------- Y axis ----------------
    y_major = np.arange(
        np.floor(ymin / 50) * 50,
        np.ceil(ymax / 50) * 50 + 1,
        50
    )
    y_minor = np.arange(
        np.floor(ymin / 10) * 10,
        np.ceil(ymax / 10) * 10 + 1,
        10
    )

    # Boundaries must be MAJOR ticks to get labels
    y_major_all = sorted(set(y_major) | set(boundaries))
    y_minor_all = sorted(set(y_minor) | set(boundaries))

    axis.yaxis.set_major_locator(FixedLocator(y_major_all))
    axis.yaxis.set_minor_locator(FixedLocator(y_minor_all))

    # Custom formatter: boundaries .2f, others .0f
    def y_formatter(y, _):
        if np.any(np.isclose(y, boundaries, atol=1e-6)):
            return f"{y:.2f}"
        else:
            return f"{y:.0f}"

    axis.yaxis.set_major_formatter(FuncFormatter(y_formatter))
    # ---------------- X axis ----------------
    # Keep symmetric limits
    max_x = max(abs(xmin), abs(xmax))
    axis.set_xlim(-max_x, max_x)

    x_major = np.arange(
        np.floor(-max_x / 5) * 5,
        np.ceil(max_x / 5) * 5 + 1,
        5
    )
    x_minor = np.arange(
        np.floor(-max_x / 1) * 1,
        np.ceil(max_x / 1) * 1 + 1,
        1
    )

    axis.xaxis.set_major_locator(FixedLocator(x_major))
    axis.xaxis.set_minor_locator(FixedLocator(x_minor))

    # Label only positive side (diameter = 2 * radius)
    axis.xaxis.set_major_formatter(
        FuncFormatter(lambda x, _: f"{2 * x:.0f}" if x >= 0 else "")
    )
    if RNA is not None:
        # Get left axis limits
        left_ymin, left_ymax = axis.get_ylim()

        # Set right axis so top = 0, increase upwards
        rax.set_ylim(left_ymin - z_top, left_ymax - z_top)

    # -------------------------------------------------
    # Boundary lines (striped / dashed, on top)
    # -------------------------------------------------
    for z in boundaries:
        axis.axhline(
            z,
            color="k",
            linestyle=(0, (6, 4)),  # striped
            linewidth=1,
            zorder=10
        )

    return fig


def plot_Assambly_Build(excel_caller):
    """
    Reads structural component data from an Excel workbook, assembles the offshore structure,
    generates an assembly plot, and inserts it back into the Excel file.

    Parameters
    ----------
    excel_caller : str
        Full path to the Excel file from which to read structure data. The filename is extracted
        and used to read various component tables from the "BuildYourStructure" sheet.

    Reads From Excel
    ----------------
    Sheet: "BuildYourStructure"
        - "MP_DATA": Main pile geometry (required).
        - "TP_DATA": Transition piece geometry (required).
        - "TOWER_DATA": Tower geometry (required).
        - "MP_META": Metadata including:
            - "Water Depth [m]": Used to determine seabed elevation.
            - "Height Reference": Optional reference appended to the vertical axis label.

    Excel Output
    ------------
    Sheet: "BuildYourStructure"
        - Plot inserted into cell anchor labeled "Assambly_plot".

    Returns
    -------
    None
        This function performs operations with side effects:
        - Generates a structural assembly plot using `plot_Assambly`.
        - Inserts the figure into the provided Excel workbook.

    Notes
    -----
    - The function calls `mc.assemble_structure` to combine MP, TP, and TOWER into a unified DataFrame.
    - If the "Water Depth [m]" value in metadata is not numeric, seabed is ignored.
    - The assembled structure is visualized using the `plot_Assambly` function with color-coded segments.
    """
    excel_filename = os.path.basename(excel_caller)
    MP = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_DATA", dtype=float, dropnan=True)
    TP = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_DATA", dtype=float, dropnan=True)
    TOWER = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_DATA", dtype=float, dropnan=True)
    META_MP = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_META", dropnan=True)
    RNA_SORCE = ex.read_excel_table(excel_filename, "StructureOverview", f"RNA", dropnan=True)

    MP["Affiliation"] = "MP"
    TP["Affiliation"] = "TP"
    TOWER["Affiliation"] = "TOWER"

    if len(META_MP) != 0:

        value = META_MP.loc[0, "Water Depth [m]"]
        seabed = -value if np.isreal(value) else None

        value = META_MP.loc[0, "Height Reference"]
        height_ref = value if value else None

    else:
        seabed = None
        height_ref = None

    if len(RNA_SORCE) > 0:
        RNA = dict()

        # Helper function to safely get a value or a default
        def safe_value(df, key, default=float('nan'), expected_type=float):
            val = df.get(key)
            if val is None or not isinstance(val.values[0], expected_type):
                return default
            return val.values[0]

        # Assign values to RNA dictionary
        RNA["dz_com"] = safe_value(RNA_SORCE, "Vertical Offset TT_COG [m]", 0.0)
        RNA["dz_hub"] = safe_value(RNA_SORCE, "Vertical Offset TT to HH [m]", 0.0)
        RNA["diameter"] = safe_value(RNA_SORCE, "Rotor Diameter [m]", 0.0)
        RNA["power"] = safe_value(RNA_SORCE, "Power [MW]", float("nan"))
        RNA["mass"] = safe_value(RNA_SORCE, "Mass of RNA [kg]", float("nan"))
        RNA["inertia_fore"] = safe_value(RNA_SORCE, "Inertia of RNA fore-aft @COG [kg m^2]", float("nan"))
        RNA["inertia_side"] = safe_value(RNA_SORCE, "Inertia of RNA side-side @COG [kg m^2]", float("nan"))
        RNA["name"] = safe_value(RNA_SORCE, "Name", "", str)

        # Construct info string
        RNA["info"] = (
            f"{RNA['name']}\n"
            f"Power [MW]: {RNA['power']}\n"
            f"Mass [kg]: {RNA['mass']}\n"
            f"Inertia fore-aft/side-side @COG [kg m^2]:\n"
            f"{RNA['inertia_fore']} / {RNA['inertia_side']}"
        )
    else:
        RNA = None

    WHOLE_STRUCTURE, _, SKIRT, _ = mc.assemble_structure(MP, TP, TOWER, interactive=False, strict_build=False, overlapp_mode="Skirt")

    Fig = plot_Assambly(WHOLE_STRUCTURE, SKIRT=SKIRT, seabed=seabed, waterlevel=0, height_ref=height_ref, RNA=RNA)
    ex.insert_plot(Fig, excel_filename, "BuildYourStructure", f"Assambly_plot")


def plot_Assambly_Overview(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    WHOLE_STRUCTURE = ex.read_excel_table(excel_filename, "StructureOverview", f"WHOLE_STRUCTURE", dropnan=True)
    SKIRT = ex.read_excel_table(excel_filename, "StructureOverview", f"SKIRT", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", f"STRUCTURE_META", dropnan=True)
    RNA_SORCE = ex.read_excel_table(excel_filename, "StructureOverview", f"RNA", dropnan=True)

    water_level = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Water level", "Value"]
    if water_level.empty:
        water_level = None
    elif water_level.values[0] is not None:
        water_level = water_level.values[0]
    else:
        water_level = None

    seabed_level = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Seabed level", "Value"]
    if seabed_level.empty:
        seabed_level = None
    elif seabed_level.values[0] is not None:
        seabed_level = -seabed_level.values[0]
    else:
        seabed_level = None

    height_ref = STRUCTURE_META.loc[STRUCTURE_META["Parameter"] == "Height Reference", "Value"]
    if type(height_ref) != str:
        height_ref = None
    else:
        height_ref = height_ref.values[0]

    if len(RNA_SORCE) > 0:
        RNA = dict()

        # Helper function to safely get a value or a default
        def safe_value(df, key, default=float('nan'), expected_type=float):
            val = df.get(key)
            if val is None or not isinstance(val.values[0], expected_type):
                return default
            return val.values[0]

        # Assign values to RNA dictionary
        RNA["dz_com"] = safe_value(RNA_SORCE, "Vertical Offset TT_COG [m]", 0.0)
        RNA["dz_hub"] = safe_value(RNA_SORCE, "Vertical Offset TT to HH [m]", 0.0)
        RNA["diameter"] = safe_value(RNA_SORCE, "Rotor Diameter [m]", 0.0)
        RNA["power"] = safe_value(RNA_SORCE, "Power [MW]", float("nan"))
        RNA["mass"] = safe_value(RNA_SORCE, "Mass of RNA [kg]", float("nan"))
        RNA["inertia_fore"] = safe_value(RNA_SORCE, "Inertia of RNA fore-aft @COG [kg m^2]", float("nan"))
        RNA["inertia_side"] = safe_value(RNA_SORCE, "Inertia of RNA side-side @COG [kg m^2]", float("nan"))
        RNA["name"] = safe_value(RNA_SORCE, "Name", "", str)

        # Construct info string
        RNA["info"] = (
            f"{RNA['name']}\n"
            f"Power [MW]: {RNA['power']}\n"
            f"Mass [kg]: {RNA['mass']}\n"
            f"Inertia fore-aft/side-side @COG [kg m^2]:\n"
            f"{RNA['inertia_fore']} / {RNA['inertia_side']}"
        )
    else:
        RNA = None

    Fig = plot_Assambly(WHOLE_STRUCTURE, SKIRT=SKIRT, waterlevel=water_level, seabed=-seabed_level, height_ref=height_ref, RNA=RNA)

    ex.insert_plot(Fig, excel_filename, "StructureOverview", f"Assambly_plot_Overview2")


def plot_MP(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_DATA", dtype=float)
    Added_Masses = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_MASSES")
    Added_Masses = Added_Masses.dropna(how="all")
    Structure = Structure.dropna(how="all")

    META = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_META", dropnan=True)

    if len(META) != 0:
        value = META.loc[0, "Water Depth [m]"]
        waterdepth = -value if np.isreal(value) else None

        value = META.loc[0, "Height Reference"]
        height_ref = value if value else None

    else:
        waterdepth = None
        height_ref = None

    FIG = plot_Structure(Structure, Added_Masses, waterdepth=waterdepth, height_ref=height_ref)

    ex.insert_plot(FIG, excel_filename, "BuildYourStructure", f"MP_plot")

    return


def plot_TP(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_DATA", dtype=float)
    Added_Masses = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_MASSES")
    Added_Masses = Added_Masses.dropna(how="all")
    Structure = Structure.dropna(how="all")

    META = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_META", dropnan=True)
    if len(META) != 0:
        value = META.loc[0, "Height Reference"]
        height_ref = value if value else None
    else:
        height_ref = None

    FIG = plot_Structure(Structure, Added_Masses, waterdepth=None, height_ref=height_ref)

    ex.insert_plot(FIG, excel_filename, "BuildYourStructure", f"TP_plot")

    return


def plot_TOWER(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_DATA", dtype=float)
    Added_Masses = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_MASSES")
    Added_Masses = Added_Masses.dropna(how="all")
    Structure = Structure.dropna(how="all")

    META = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_META", dropnan=True)
    if len(META) != 0:
        value = META.loc[0, "Height Reference"]
        height_ref = value if value else None
    else:
        height_ref = None

    FIG = plot_Structure(Structure, Added_Masses, waterdepth=None, height_ref=height_ref, waterlevel=None, show_section_numbers=False)

    ex.insert_plot(FIG, excel_filename, "BuildYourStructure", f"TOWER_plot")

    return


def plot_modeshapes(
    data,
    order=(1, 2),
    waterlevels=None,
    max_lines=10,
    max_label_length=100,
):

    # ---------------------------------------------------------
    # Overview table (first DF)
    # ---------------------------------------------------------
    columns = ["config Nr.", "config"]
    columns += [
        item
        for o in order
        for item in (f"f order {o} [Hz]", f"alpha order {o} [%]")
    ]

    overview_df = pd.DataFrame(columns=columns)

    # ---------------------------------------------------------
    # Prepare figure
    # ---------------------------------------------------------
    fig, axis = plt.subplots(1, len(order), figsize=[6 * len(order), 8])

    if len(order) == 1:
        axis = [axis]

    if waterlevels is not None:
        if len(waterlevels) != len(data):
            raise ValueError("Waterlevels must have same length as data")

    # ---------------------------------------------------------
    # Determine if short labels should be used
    # ---------------------------------------------------------
    test_labels = []
    first_order = order[0]

    for config, values in data.items():
        freq_str = values.columns[first_order + 1]
        for r in ["Mode shape", "(", ")", "f", "=", "Hz"]:
            freq_str = freq_str.replace(r, "")
        frequency = np.round(float(freq_str), 4)

        if waterlevels is not None:
            test_labels.append(f"{config} ({frequency}), alpha(WL)")
        else:
            test_labels.append(f"{config} ({frequency})")

    use_short_labels = any(len(lbl) > max_label_length for lbl in test_labels)

    # ---------------------------------------------------------
    # Prepare modeshape tables (one per order)
    # ---------------------------------------------------------
    modeshape_tables = []

    for n in order:
        first_config_df = next(iter(data.values()))
        z_values = first_config_df["z"].values

        df_mode = pd.DataFrame()
        df_mode["z [m]"] = z_values

        modeshape_tables.append(df_mode)

    # ---------------------------------------------------------
    # Plotting + Table Filling
    # ---------------------------------------------------------
    suppress_plot = len(data) > max_lines

    for mode_index, (n, ax) in enumerate(zip(order, axis)):

        ax.set_title(f"Modeshapes of order {n}")
        ax.set_ylabel("z in m")
        ax.set_xlabel("normalised displacement [-]")
        ax.grid(True)

        colors = get_JBO_colors(len(data))

        for i, (config, values) in enumerate(data.items()):

            # -----------------------------
            # Extract modeshape
            # -----------------------------
            shape = values.iloc[:, n + 1]
            shape = shape * np.sign(shape.iloc[0])

            # -----------------------------
            # Extract frequency
            # -----------------------------
            freq_str = values.columns[n + 1]
            for r in ["Mode shape", "(", ")", "f", "=", "Hz"]:
                freq_str = freq_str.replace(r, "")

            frequency = np.round(float(freq_str), 4)

            # -----------------------------
            # Alpha at waterlevel
            # -----------------------------
            if waterlevels is not None:
                level = float(waterlevels[config])
                alpha_value = abs(
                    np.round(
                        shape.loc[values["z"] == level].iloc[0] * 100,
                        2,
                    )
                )
                overview_df.loc[config, f"alpha order {n} [%]"] = alpha_value
            else:
                alpha_value = None

            # -----------------------------
            # Fill overview table
            # -----------------------------
            overview_df.loc[config, "config Nr."] = i + 1
            overview_df.loc[config, "config"] = config
            overview_df.loc[config, f"f order {n} [Hz]"] = frequency

            # -----------------------------
            # Fill modeshape table
            # -----------------------------
            column_name = f"{i+1}: {config}"
            modeshape_tables[mode_index][column_name] = shape.values

            # -----------------------------
            # Skip plotting if too many lines
            # -----------------------------
            if suppress_plot:
                continue

            # -----------------------------
            # Legend label handling
            # -----------------------------
            if use_short_labels:
                label = f"config {i+1}"
            else:
                if waterlevels is not None:
                    label = (
                        f"{config} ({frequency}), "
                        f"$\\alpha(WL) = {alpha_value}\\%$"
                    )
                else:
                    label = f"{config} ({frequency})"

            # -----------------------------
            # Plot
            # -----------------------------
            ax.plot(shape, values["z"], label=label, color=colors[i])

        # -------------------------------------------------
        # Plot decorations
        # -------------------------------------------------
        if suppress_plot:
            ax.text(
                0.5,
                0.5,
                f"Too many lines to display\n"
                f"(current lines: {len(data)}, max lines: {max_lines})",
                ha="center",
                va="center",
                transform=ax.transAxes,
                fontsize=12,
            )
            ax.set_xticks([])
            ax.set_yticks([])
        else:
            ax.axvline(0, linewidth=1.5, color="k", linestyle="--")
            ax.legend(loc="lower right")

    fig.tight_layout()

    # ---------------------------------------------------------
    # Return list of DataFrames
    # ---------------------------------------------------------
    result_tables = [overview_df] + modeshape_tables

    return fig, result_tables



def export_Modeshapes(excel_caller, jboost_path):
    excel_filename = os.path.basename(excel_caller)
    try:
        MODESHAPES_TABLE = ex.read_excel_table(excel_filename, "ExportStructure", f"MODESHAPE_OVERVIEW")

        ex.save_excel_picture_as_png(excel_filename, "ExportStructure", "Fig_FIG_JBOOST_MODESHAPES", os.path.join(jboost_path, "Modeshape_overview.png"))
        MODESHAPES_TABLE.to_csv(os.path.join(jboost_path, "Modeshape_overview.csv"), index=False)
    except Exception as e:
        ex.show_message_box(excel_filename, f"Error exporting Modeshapes: {e}")

        return
    ex.show_message_box(excel_filename, f"Modeshapes exported successfully at {jboost_path}")

    return


def plot_py_curves(data, max_lines=10, crop_symetric=False, loadcase=None):
    Springs = list(np.unique(data["Spring [-]"].values))
    heights = list(np.unique(data["z [m]"].values))

    NPlot = int(np.ceil(len(Springs) / max_lines))

    fig, axis = plt.subplots(1, NPlot, figsize=[5 * NPlot, 6], dpi=800)

    if len(axis) == 0:
        axis = [axis]

    Springs_group = [Springs[i:i + max_lines] for i in range(0, len(Springs), max_lines)]
    heights_group = [heights[i:i + max_lines] for i in range(0, len(heights), max_lines)]

    for Spring_group, height_group, ax in zip(Springs_group, heights_group, axis):

        ax.set_title(f"py-curves for springs: {Spring_group[0]} ({round(height_group[0], 2)}m) to {Spring_group[-1]} ({round(height_group[-1], 2)}m)")

        ax.set_xlabel("y [m]")
        ax.set_ylabel("p [KN/m]")
        ax.grid(True)
        colors = get_JBO_colors(len(Spring_group))

        for Spring, height, color in zip(Spring_group, height_group, colors):
            p = data.loc[data["Spring [-]"] == Spring, "p [kN/m]"].values
            y = data.loc[data["Spring [-]"] == Spring, "y [m]"].values

            if crop_symetric:
                p = p[p >= 0]
                y = y[y >= 0]

            ax.plot(y, p, color=color, label=f"Spring {str(Spring).zfill(2)} at {round(height, 2):.2f}m b. SB")

        ax.legend(loc="lower right")

    fig.suptitle(
        (f"py-curves for {loadcase}" if loadcase is not None else "") +
        (" cropped to positive axis" if crop_symetric else ""),
        fontsize=13,
        fontweight="bold",
        x=0.1,  # left edge of the figure
        ha="left"  # align text to the left
    )
    fig.tight_layout()
    return fig
