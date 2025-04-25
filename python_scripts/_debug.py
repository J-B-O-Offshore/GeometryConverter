import excel as ex

sheet_name_structure_loading = "BuildYourStructure"

#
excel_path = "C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/GeometrieConverter.xlsm"
#value = ex.get_textbox_value(excel_path, "BuildYourStructure", "TextBox_MP_db_path")

print("1")

import xlwings as xw
import os


def list_shapes(excel_path, sheet_name):
    abs_path = os.path.abspath(excel_path)

    with xw.App(visible=False) as app:
        wb = app.books.open(abs_path)
        sheet = wb.sheets[sheet_name]
        print("Shapes on the sheet:")
        for shape in sheet.api.Shapes:
            print(f"- Name: {shape.Name}, Type: {shape.Type}")
        wb.close()

list_shapes(excel_path, "BuildYourStructure")

table_name = ex.get_dropdown_value("GeometrieConverter.xlsm", sheet_name_structure_loading, "Dropdown_MP_Structures")

print(table_name)