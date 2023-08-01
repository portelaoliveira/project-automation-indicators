import pandas as pd
from pathlib import Path

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

path_backup = Path(r"Backup_Lojas")
path_backup.mkdir(parents=True, exist_ok=True)

files_folder_backup = path_backup.iterdir()

list_names_backup = [files.name for files in files_folder_backup]

for store in dict_stores:
    if store not in list_names_backup:
        store_formated = store.replace(" ", "_")
        new_folder_stores = path_backup / store_formated
        new_folder_stores.mkdir(parents=True, exist_ok=True)

    name_file = f"{day_indicators.day}_{day_indicators.month}_{day_indicators.year}_{store_formated}.xlsx"
    local_file = path_backup / store_formated / name_file

    dict_stores[store].to_excel(local_file)
