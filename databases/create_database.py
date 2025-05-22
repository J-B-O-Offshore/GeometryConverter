import sqlite3

import pandas as pd


MP01 = pd.read_csv("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/TP/DP-A1_L0_G0_S0.CSV")
MP02 = pd.read_csv("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/TP/DP-A1_L0_G1_S1.CSV")
MP01_added = pd.read_csv("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/TP/DP-A1_L0_G0_S0__ADDED_MASSES.CSV")
MP02_added = pd.read_csv("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/TP/DP-A1_L0_G1_S1__ADDED_MASSES.CSV")

META = pd.read_csv("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/TP/META.csv")

db = sqlite3.connect('C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/TP.db')

MP01.to_sql('24A535_FEED_DP-A1_L0_G0_S0', db)
MP02.to_sql('24A535_FEED_DP-A1_L0_G1_S1', db)
MP01_added.to_sql('24A535_FEED_DP-A1_L0_G0_S0__ADDED_MASSES', db)
MP02_added.to_sql('24A535_FEED_DP-A1_L0_G1_S1__ADDED_MASSES', db)
META.to_sql('META', db, if_exists='replace')

db.close()

