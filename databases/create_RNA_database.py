import sqlite3

import pandas as pd


data = pd.read_csv("C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/RNAs/data.csv")

db = sqlite3.connect('C:/Users/aaron.lange/Desktop/Projekte/Geometrie_Converter/GeometrieConverter/databases/RNAs.db')

data.to_sql('data', db, if_exists='replace')

db.close()

