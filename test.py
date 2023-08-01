import pandas as pd
import pathlib

emails = pd.read_excel(r"data/Emails.xlsx")
stores = pd.read_csv(r"data/Lojas.csv", sep=";", encoding="latin-1")
sales = pd.read_excel(r"data/Vendas.xlsx")

sales = sales.merge(stores, on="ID Loja")

dict_stores = dict()
for store in stores["Loja"]:
    dict_stores[store] = sales.loc[sales["Loja"] == store, :]

day_indicators = sales["Data"].max()
day_indicators_formated = (
    f"{day_indicators.day}/{day_indicators.month}/{day_indicators.year}"
)
print(day_indicators_formated)
