import xlwings as xw
from db_loading import load_db_table
import excel as ex


def main(Structure_name, db_path):
    logger = ex.setup_logger()

    sheet_name_structure_loading = "BuildYourStructure"

    META = load_db_table(db_path, "META")
    DATA = load_db_table(db_path, Structure_name)

    META_relevant = META.loc[META["table_name"] == Structure_name]
   # ex.write_df_to_existing_table("GeometrieConverter.xlsm", sheet_name_structure_loading, "MP_Data", META_relevant)

    ex.write_df("GeometrieConverter.xlsm", sheet_name_structure_loading, "Table_TP_META", META_relevant)
    ex.write_df("GeometrieConverter.xlsm", sheet_name_structure_loading, "Table_TP_DATA", DATA)



    return
