import os
import excel as ex
import numpy as np

from collections import defaultdict
import matplotlib.pyplot as plt
import pandas as pd
import misc as mc


def plot_Structure(Structure, Added_Masses, waterdepth=None, height_ref="", waterlevel=0, show_section_numbers=True):
    fig, ax = plt.subplots(1, 3, figsize=[22, 6])
    top_all = max(Structure["Top [m]"].max(), Added_Masses["Top [m]"].max())
    bottom_all = min(Structure["Bottom [m]"].min(), Added_Masses["Bottom [m]"].min())

    left = -max(Structure["D, top [m]"].max(), Structure["D, bottom [m]"].max()) / 2
    right = -left

    for ax_curr in ax:
        ax_curr.set_ylim(bottom_all - 0.05 * (top_all - bottom_all), top_all + 0.05 * (top_all - bottom_all))
        ax_curr.grid(True, linestyle='--')
    # -------------------------------
    # Plot 1: Structure overview
    # -------------------------------
    axis = ax[0]

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
    axis.set_xlabel("Diameter in [m]")
    axis.set_ylabel(f"z in m{height_ref}")
    axis.set_title("Overview and Diameter")

    # -------------------------------
    # Plot 2: Added masses with smart labels
    # -------------------------------
    axis = ax[1]
    weight = mc.calc_weight(
        7850,
        Structure.loc[:, "t [mm]"],
        Structure.loc[:, "Top [m]"],
        Structure.loc[:, "Bottom [m]"],
        Structure.loc[:, "D, top [m]"],
        Structure.loc[:, "D, bottom [m]"]
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

        if top == bottom:
            z = top
            is_point = True
        else:
            z = (top + bottom) / 2
            is_point = False

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

    # --- Handle too many labels ---
    MAX_LABELS = 25
    too_many_labels = len(label_entries) > MAX_LABELS

    if not too_many_labels:
        # --- Improved label placement ---
        label_entries.sort(key=lambda x: -x['y'])  # top to bottom
        occupied_y = []
        min_dy = 0.03 * (top_all - bottom_all)
        x_text = 0.6 * abs(right)

        for entry in label_entries:
            target_y = entry['y']
            while any(abs(target_y - oy) < min_dy for oy in occupied_y):
                target_y -= min_dy
            occupied_y.append(target_y)

            # Draw connector
            axis.plot([entry['x_start'], x_text - 0.05], [entry['y'], target_y],
                      linestyle='-', color=entry['color'], linewidth=1)

            # Draw label
            axis.text(x_text, target_y, entry['label'], ha='left', va='center',
                      fontsize=9, color=entry['color'])
    else:
        axis.text(
            0.8 * abs(right),
            0.5 * (top_all + bottom_all),
            "Too many masses to label",
            ha='center',
            va='center',
            fontsize=13,
            color='red'
        )

    # Final axis settings
    axis.axvline(0, color="grey", linestyle="--", linewidth=1)
    axis.set_xlim(-0.1 * abs(right), 2 * abs(right))
    axis.set_title("Masses \n (horizontal displacement only for distinction, all lie on center line)")
    axis.set_xticklabels([])
    axis.set_xticks([])

    # -------------------------------
    # Plot 3: Slope, Wall Thickness, D/t Ratio

    z_nodes = list(Structure.loc[:, "Top [m]"]) + [Structure.loc[:, "Bottom [m]"].values[-1]]

    t = list(Structure.loc[:, "t [mm]"] / 1000)
    slope = -(Structure.loc[:, "D, top [m]"] - Structure.loc[:, "D, bottom [m]"]) / (Structure.loc[:, "Top [m]"] - Structure.loc[:, "Bottom [m]"])
    D_nodes = list(Structure.loc[:, "D, top [m]"]) + [Structure.loc[:, "D, bottom [m]"].values[-1]]

    # D/t
    z_d_by_t = []
    d_by_t = []
    for id, row in Structure.iterrows():
        t_curr = row["t [mm]"]

        z_d_by_t.append(row["Top [m]"])
        z_d_by_t.append(row["Bottom [m]"])

        d_by_t.append(row["D, top [m]"] / t_curr)
        d_by_t.append(row["D, bottom [m]"] / t_curr)

    # Step values for t and slope (constant over each segment, len = len(z_nodes)-1)
    slope_steps = list(slope)

    # Plotting
    axis = ax[2]
    axis2 = axis.twiny()  # upper x-axis for D/t

    # Plot D/t (upper x-axis)
    axis2.plot(d_by_t, z_d_by_t, label="D/t", color="C0")
    axis2.set_xlabel("D/t [-]")
    axis2.xaxis.set_label_position('top')
    axis2.xaxis.tick_top()

    slope_steps = [float("nan") if slope == 0 else slope for slope in slope_steps]

    # Plot on both axes
    line1 = axis.stairs(slope_steps, z_nodes, orientation="horizontal", label="slope (where not 0)", color="C1", baseline=None)
    line2 = axis.stairs(t, z_nodes, label="t [mm]", orientation="horizontal", color="C2", baseline=None)

    # Combine legend entries
    lines = [line1, line2]
    labels = [line.get_label() for line in lines]

    # Add combined legend to main axis
    axis.legend(lines, labels, loc="lower left")

    axis.set_title("Slope, Wall Thickness, D/t Ratio")
    return fig


def plot_cans(Structure, axis, show_section_numbers=False, set_lims=True, **plot_kwargs):

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


def plot_Assambly(WHOLE_STRUCTURE, SKIRT=None, seabed=None, waterlevel=0):

    fig, axis = plt.subplots(1, 1, figsize=[8, 27])

    MP_assambled = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE["Affiliation"] == "MP", :]
    TP_assambled = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE["Affiliation"] == "TP", :]
    TOWER_assambled = WHOLE_STRUCTURE.loc[WHOLE_STRUCTURE["Affiliation"] == "TOWER", :]

    if len(MP_assambled) != 0:
        plot_cans(MP_assambled, axis, show_section_numbers=False, color="k", set_lims=False)
    if len(TP_assambled) != 0:
        plot_cans(TP_assambled, axis, show_section_numbers=False, color="blue", set_lims=False)
    if len(TOWER_assambled) != 0:
        plot_cans(TOWER_assambled, axis, show_section_numbers=False, color="grey", set_lims=False)
    if SKIRT is not None:
        plot_cans(SKIRT, axis, show_section_numbers=False, color="red", set_lims=False, alpha=0.8)

    axis.axhline(waterlevel, color="blue", linestyle="--")
    axis.axvline(0, color="grey", linestyle="--", linewidth=1)
    if seabed is not None:
        axis.axhline(seabed, color="brown", linestyle="-", linewidth=2)

    return fig

def plot_Assambly_Build(excel_caller):

    excel_filename = os.path.basename(excel_caller)
    MP = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_DATA", dtype=float, dropnan=True)
    TP = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_DATA", dtype=float, dropnan=True)
    TOWER = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_DATA", dtype=float, dropnan=True)
    META_MP = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_META", dropnan=True)

    if len(META_MP) != 0:
        seabed = -META_MP.loc[0, "Water Depth [m]"]
    else:
        seabed = None

    WHOLE_STRUCTURE, _, SKIRT, _ = mc.assemble_structure(MP, TP, TOWER, interactive=False, ignore_hovering=True, overlapp_mode="Skirt")

    Fig = plot_Assambly(WHOLE_STRUCTURE, SKIRT=SKIRT, seabed=seabed, waterlevel=0)

    ex.insert_plot(Fig, excel_filename, "BuildYourStructure", f"Assambly_plot")


def plot_Assambly_Overview(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    WHOLE_STRUCTURE = ex.read_excel_table(excel_filename, "StructureOverview", f"WHOLE_STRUCTURE", dtype=float, dropnan=True)
    ALL_ADDED_MASSES = ex.read_excel_table(excel_filename, "StructureOverview", f"ALL_ADDED_MASSES", dropnan=True)
    SKIRT_POINTMASS = ex.read_excel_table(excel_filename, "StructureOverview", f"SKIRT_POINTMASS", dropnan=True)
    SKIRT = ex.read_excel_table(excel_filename, "StructureOverview", f"SKIRT", dropnan=True)
    MARINE_GROWTH = ex.read_excel_table(excel_filename, "StructureOverview", f"MARINE_GROWTH", dropnan=True)
    HYDRO_COEFFICIENTS = ex.read_excel_table(excel_filename, "StructureOverview", f"HYDRO_COEFFICIENTS", dropnan=True)
    STRUCTURE_META = ex.read_excel_table(excel_filename, "StructureOverview", f"STRUCTURE_META", dropnan=True)

    Fig = plot_Assambly(WHOLE_STRUCTURE, SKIRT=SKIRT, seabed=seabed, waterlevel=0)

    ex.insert_plot(Fig, excel_filename, "BuildYourStructure", f"Assambly_plot")


def plot_MP(excel_caller):
    excel_filename = os.path.basename(excel_caller)
    Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_DATA", dtype=float)
    Added_Masses = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_MASSES")
    Added_Masses = Added_Masses.dropna(how="all")
    Structure = Structure.dropna(how="all")

    META = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_META")
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

    META = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TP_META")
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

    META = ex.read_excel_table(excel_filename, "BuildYourStructure", f"TOWER_META")
    if len(META) != 0:
        value = META.loc[0, "Height Reference"]
        height_ref = value if value else None
    else:
        height_ref = None

    FIG = plot_Structure(Structure, Added_Masses, waterdepth=None, height_ref=height_ref, waterlevel=None, show_section_numbers=False)

    ex.insert_plot(FIG, excel_filename, "BuildYourStructure", f"TOWER_plot")

    return


#plot_Assambly_Build("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/GeometrieConverter.xlsm")
