

import matplotlib.pyplot as plt
import os
import excel as ex

def plot_Structure(Structure, Added_Masses):

    fig, ax = plt.subplots(1, 4, figsize=[22, 7])
    ax_d = ax[0]

    for id, can in Structure.iterrows():

        # corners
        up_left = (-can.loc["D, top [m]"]/2, can.loc["Top [m]"])
        up_right = (can.loc["D, top [m]"] / 2, can.loc["Top [m]"])
        down_left = (-can.loc["D, bottom [m]"] / 2, can.loc["Bottom [m]"])
        down_right = (can.loc["D, bottom [m]"] / 2, can.loc["Bottom [m]"])

        # left edge
        ax_d.plot([down_left[0], up_left[0]], [down_left[1], up_left[1]])

        # right edge
        ax_d.plot([down_right[0], up_right[0]], [down_right[1], up_right[1]])

        # lower diamter
        ax_d.plot([down_left[0], down_right[0]], [down_left[1], down_right[1]])

        # upper diamter
        ax_d.plot([up_left[0], up_right[0]], [up_left[1], up_right[1]])






    print(1)

    return fig


def plot_MP(excel_caller):

    excel_filename = os.path.basename(excel_caller)
    Structure = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_DATA", dtype=float)
    Added_Masses = ex.read_excel_table(excel_filename, "BuildYourStructure", f"MP_MASSES")

    FIG = plot_Structure(Structure, Added_Masses)

    ex.insert_plot(FIG, excel_filename, "BuildYourStructure", f"MP_plot")

    return

#plot_MP("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/GeometrieConverter.xlsm")