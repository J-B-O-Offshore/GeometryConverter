import xlwings
import pandas as pd
import sqlite3
import excel as ex

def assemble_structure():

    MP_DATA = ex.read_excel_table("GeometrieConverter.xlsx", "BuildYourStructure", "MP_DATA")
    TP_DATA = ex.read_excel_table("GeometrieConverter.xlsx", "BuildYourStructure", "TP_DATA")
    TOWER_DATA = ex.read_excel_table("GeometrieConverter.xlsx", "BuildYourStructure", "TOWER_DATA")



    return