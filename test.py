import pandas as pd
from pathlib import Path

import mimetypes
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Optional

from config import *


def send_file_email(
    file_path: list[str] | Path,
    email_addresses: Optional[list[str]] = None,
    subject: Optional[str] = None,
    body: Optional[str] = None,
):
    list_file_path = file_path
    message = EmailMessage()
    message["From"] = USER_MAIL
    if email_addresses:
        to = ", ".join(email_addresses)
    else:
        to = USER_MAIL
    message["To"] = to
    if subject:
        message["Subject"] = subject
    if body:
        message.add_alternative(body, subtype="html")
    for file_path_ in list_file_path:
        with open(file_path_, "rb") as f:
            ctype, encoding = mimetypes.guess_type(file_path_)
            if ctype is None or encoding is not None:
                ctype = "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)
            message.add_attachment(
                f.read(),
                maintype=maintype,
                subtype=subtype,
                filename=file_path_.name,
            )
    session = smtplib.SMTP("smtp.gmail.com", 587)
    session.starttls()
    session.login(USER_MAIL, USER_PASS)
    # session.send_message(message)
    session.sendmail(USER_MAIL, to, message.as_string())
    session.quit()


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


goal_billing_day = 1000
goal_billing_year = 1650000
goal_qtproducts_day = 4
goal_qtproducts_year = 120
goal_ticketmean_day = 500
goal_ticketmean_year = 500

for store in dict_stores:
    store_formated = store.replace(" ", "_")
    sales_store = dict_stores[store]
    sales_store_day = sales_store.loc[sales_store["Data"] == day_indicators, :]

    # faturamento
    invoicing_year = sales_store["Valor Final"].sum()
    # print(faturamento_ano)
    invoicing_day = sales_store_day["Valor Final"].sum()
    # print(faturamento_dia)

    # diversidade de produtos
    qt_products_year = len(sales_store["Produto"].unique())
    # print(qtde_produtos_ano)
    qt_products_day = len(sales_store_day["Produto"].unique())
    # print(qtde_produtos_dia)

    # ticket medio
    value_sales = sales_store.groupby("Código Venda")["Valor Final"].sum()
    ticket_mean_year = value_sales.mean()
    # print(ticket_medio_ano)
    # ticket_medio_dia
    value_sales_day = sales_store_day.groupby("Código Venda")[
        "Valor Final"
    ].sum()
    ticket_mean_day = value_sales_day.mean()
    # print(ticket_medio_dia)

    # enviar o e-mail

    name = emails.loc[emails["Loja"] == store, "Gerente"].values[0]
    to = emails.loc[emails["Loja"] == store, "E-mail"].values[0]
    subject = f"OnePage Dia {day_indicators_formated} - Loja {store}"

    if invoicing_day >= goal_billing_day:
        color_invoicing_day = "green"
    else:
        color_invoicing_day = "red"
    if invoicing_year >= goal_billing_year:
        color_invoicing_year = "green"
    else:
        color_invoicing_year = "red"
    if qt_products_day >= goal_qtproducts_day:
        color_qt_day = "green"
    else:
        color_qt_day = "red"
    if qt_products_year >= goal_qtproducts_year:
        color_qt_year = "green"
    else:
        color_qt_year = "red"
    if ticket_mean_day >= goal_ticketmean_day:
        color_ticket_day = "green"
    else:
        color_ticket_day = "red"
    if ticket_mean_year >= goal_ticketmean_year:
        color_ticket_year = "green"
    else:
        color_ticket_year = "red"

    body = f"""\
    <!DOCTYPE html>
    <html>
    <p>Bom dia, {name}</p>

    <p>O resultado de ontem <strong>({day_indicators_formated})</strong> da <strong>Loja {store}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${invoicing_day:.2f}</td>
        <td style="text-align: center">R${goal_billing_day:.2f}</td>
        <td style="text-align: center"><font color="{color_invoicing_day}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qt_products_day}</td>
        <td style="text-align: center">{goal_qtproducts_day}</td>
        <td style="text-align: center"><font color="{color_qt_day}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_mean_day:.2f}</td>
        <td style="text-align: center">R${goal_ticketmean_day:.2f}</td>
        <td style="text-align: center"><font color="{color_ticket_day}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${invoicing_year:.2f}</td>
        <td style="text-align: center">R${goal_billing_year:.2f}</td>
        <td style="text-align: center"><font color="{color_invoicing_year}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qt_products_year}</td>
        <td style="text-align: center">{goal_qtproducts_year}</td>
        <td style="text-align: center"><font color="{color_qt_year}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_mean_year:.2f}</td>
        <td style="text-align: center">R${goal_ticketmean_year:.2f}</td>
        <td style="text-align: center"><font color="{color_ticket_year}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Portela</p>
    </html>
    """

    # Anexos (pode colocar quantos quiser):
    attachment = (
        Path.cwd()
        / path_backup
        / store_formated
        / f"{day_indicators.day}_{day_indicators.month}_{day_indicators.year}_{store_formated}.xlsx"
    )

    send_file_email([attachment], [to], subject, body)
    print("E-mail da Loja {} enviado".format(store))

invoicing_stores = sales.groupby("Loja")[["Valor Final"]].sum()
invoicing_stores_year = invoicing_stores.sort_values(
    by="Valor Final", ascending=False
)

name_file_year = f"{day_indicators.day}_{day_indicators.month}_{day_indicators.year}_Ranking_Anual.xlsx"
name_folder_year = Path(r"Backup_Lojas/Ranking_Anual")
name_folder_year.mkdir(parents=True, exist_ok=True)
invoicing_stores_year.to_excel(name_folder_year / name_file_year)

sales_day = sales.loc[sales["Data"] == day_indicators, :]
invoicing_stores_day = sales_day.groupby("Loja")[["Valor Final"]].sum()
invoicing_stores_day = invoicing_stores_day.sort_values(
    by="Valor Final", ascending=False
)

name_file_day = f"{day_indicators.day}_{day_indicators.month}_{day_indicators.year}_Ranking_Dia.xlsx"
name_folder_day = Path(r"Backup_Lojas/Ranking_Dia")
name_folder_day.mkdir(parents=True, exist_ok=True)
invoicing_stores_day.to_excel(name_folder_day / name_file_day)

# enviar o e-mail

to_ = emails.loc[emails["Loja"] == "Diretoria", "E-mail"].values[0]
subject_ = (
    "Ranking Dia"
    f" {day_indicators.day}/{day_indicators.month}/{day_indicators.year}"
)

list_attachment = [
    Path.cwd() / name_folder_day / name_file_day,
    Path.cwd() / name_folder_year / name_file_year,
]

body_board = f"""
<!DOCTYPE html>
<html>
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {invoicing_stores_day.index[0]} com Faturamento R${invoicing_stores_day.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {invoicing_stores_day.index[-1]} com Faturamento R${invoicing_stores_day.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {invoicing_stores_year.index[0]} com Faturamento R${invoicing_stores_year.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {invoicing_stores_year.index[-1]} com Faturamento R${invoicing_stores_year.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Portela
</html>
"""

send_file_email(list_attachment, [to_], subject_, body_board)
print("E-mail da Diretoria enviado")
