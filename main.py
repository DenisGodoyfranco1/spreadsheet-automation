import os 
from datetime import datetime
import pandas as pd
import win32com.client as win

caminho = 'bases'
arquivos = os.listdir(caminho)


tabelas_consolidada = pd.DataFrame()

for nome_arquivos in arquivos:
    tabelas_vendas = pd.read_csv(os.path.join(caminho, nome_arquivos))
    tabelas_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabelas_vendas["Data de Venda"], unit="d")
    tabelas_consolidada = pd.concat([tabelas_consolidada,tabelas_vendas])
    

tabelas_consolidada = tabelas_consolidada.sort_values(by="Data de Venda")
tabelas_consolidada = tabelas_consolidada.reset_index(drop=True)
tabelas_consolidada.to_excel("Vendas.xlsx", index=False)

outlook = win.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = "denisgoody203@gmail.com"
data_hoje = datetime.today().strftime("%d/%m/%Y")
email.Subject = f"Relatório de Vendas {data_hoje}"
email.Body = f"""
prezados,
Segue a anexo de hoje {data_hoje} totalmente atualizado
Abs,
Denis Godoy Franco
"""
caminho = os.getcwd()
anexo = os.path.join(caminho, "Vendas.xlsx")
email.Attachments.Add(anexo)

email.Send()