# %%
import win32com.client
import time
import pandas as pd
import subprocess
import win32gui
import win32con
import sqlite3
from sqlite_utils import Database
import sys
from openpyxl import load_workbook
import requests
import socket
import json
import base64
import codecs
import pandas as pd
import openpyxl
from datetime import date
import warnings
from datetime import date, timedelta
import socket
import locale
from datetime import date, timedelta
import os

# %%
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8') 

mes = {
    'janeiro': ["2024-01-01", "2024-01-31"],
    'fevereiro': ["2024-02-01", "2024-02-29"],
    'marco': ["2024-03-01", "2024-03-31"],
    'abril': ["2024-04-01", "2024-04-30"],
    'maio': ["2024-05-01", "2024-05-31"],
    'junho': ["2024-06-01", "2024-06-30"],
    'julho': ["2024-07-01", "2024-07-31"],
    'agosto': ["2024-08-01", "2024-08-31"],
    'setembro': ["2024-09-01", "2024-09-30"],
    'outubro': ["2024-10-01", "2024-10-31"],
    'novembro': ["2024-11-01", "2024-11-30"],
    'dezembro': ["2024-12-01", "2024-12-31"]
}



# %%
def expedidas_apollo(mes_relatorio):
        hoje = str(date.today())
        ip_local = socket.gethostbyname(socket.gethostname())
        mes_atual = date.today().strftime("%B").lower()
        mes_periodo = mes[mes_relatorio.lower()]
        mes_caminho = mes_relatorio.lower()

        if mes_relatorio.lower() == mes_atual:
            fim_periodo = hoje  
        else:
            fim_periodo = mes_periodo[1]  

        print(mes_periodo[0], fim_periodo)

        result = {
            "autenticacao": {
                "chave": "Substitua pela sua chave",
                "token": "Substitua pelo seu token"
            },
            "controller": "relatorio",
            "action": "gerar",
            "parametros": {
                "idUsuario": 'Substitua pelo seu ID',
                "ipClient": ip_local,
                "procedure": "excel.statusPorRemessa",
                "parametrosConsulta": {
                    "inicioPeriodo": mes_periodo[0],
                    "fimPeriodo": fim_periodo,
                    "origem": "",
                    "remessa": "",
                    "somenteExpedido": True,
                    "somenteNaoExpedido": False,
                }
            }
        }
        headers = {'Content-type': 'application/json'}
        url = "Substitua pela sua URL"
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        jstr = json.dumps(
            result, default=lambda df: json.loads(df.to_json()), indent=4)
        requisicao = requests.post(url, json=result, verify=False)
        jsonResponse = requisicao.json()["CoreModule"]["conteudo"]["mensagem"]
        message = base64.b64decode(jsonResponse)
        caminho_arquivo = 'substitua pelo seu caminho'
        with open(os.path.join(caminho_arquivo, f"{mes_caminho}.xlsx"), 'wb') as binary_file:
            binary_file.write(message)
            arquivo = fr'os.path.join(caminho_arquivo, f"{mes_caminho}.xlsx'

# %%
meses_ordenados = list(mes.keys())
mes_atual = date.today().strftime("%B").lower()
mes_atual_index = meses_ordenados.index(mes_atual)

for m in meses_ordenados[:mes_atual_index + 1]: 

    expedidas_apollo(m)
    print(f"a Base do mes de {m} foi atualizada")
    


