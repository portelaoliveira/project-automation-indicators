import pandas as pd
import pathlib

emails = pd.read_excel(r"data/Emails.xlsx")
stores = pd.read_csv(r"data/Lojas.csv", sep=";", encoding="latin-1")
sales = pd.read_excel(r"data/Vendas.xlsx")
