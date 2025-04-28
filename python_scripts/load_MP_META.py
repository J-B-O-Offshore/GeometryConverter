import xlwings as xw
from db_loading import load_db_table
import excel as ex

def main(db_path):
    logger = ex.setup_logger()

    sheet_name_structure_loading = "BuildYourStructure"

    logger.debug(db_path)
    META = load_db_table(db_path, "META")

    ex.set_dropdown_values("GeometrieConverter.xlsm", sheet_name_structure_loading, "Dropdown_MP_Structures", list(META.loc[:, "table_name"].values))

    return


