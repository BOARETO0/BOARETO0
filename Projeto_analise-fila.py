import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options as Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time
import os 
import glob
import win32com.client as win32
from datetime import datetime 
import openpyxl
import locale
import locale
import locale
import pandas as pd
import numpy as np
from datetime import datetime
from workalendar.america import Brazil


links_baldes = {
    'romaneio_maual': 'a', 
    'aguard_agendamento': 'b'
}
usuario_intranet = 'seu usuario' 
senha_intranet = 'sua senha' 
pasta_downloads = 'sua pasta de download' 
login_xpath = '//*[@id="login"]' 
senha_xpath = '//*[@id="senha"]' 
acessar_xpath = '//*[@id="greCAPTCHA"]/div[2]/button' 

arquivos_processados = set() 
chrome_options = Options()
prefs = {
    "download.default_directory": pasta_downloads,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing_for_trusted_sources_enabled": False,
    "safebrowsing.enabled": False
}
chrome_options.add_experimental_option("prefs", prefs) 
drive = webdriver.Chrome(options=chrome_options) 
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8') 
cal = Brazil() 
hoje = datetime.today().date()
ufs_valor_minimo_600 = [
    'ES', 'RJ', 'SP', 'MG'
]
valor_minimo = 400 
data_hoje = datetime.now().strftime("%d-%m-%Y")



def aguardar_download_completo(pasta_downloads, timeout=300):

    caminho_arquivo_especifico = os.path.join(pasta_downloads, "ListaPedidosCompleto.xls")

    while True:

        arquivos_xls = glob.glob(os.path.join(pasta_downloads, "*.xls*"))
        arquivos_completos = [arq for arq in arquivos_xls if 'crdownload' not in arq]

        if caminho_arquivo_especifico in arquivos_completos:
            ultimo_modificado = os.path.getctime(caminho_arquivo_especifico)
            if time.time() - ultimo_modificado > 2:
                return caminho_arquivo_especifico
            
        time.sleep(2)

def converter_ultimo_download_para_xlsx(pasta_downloads, nome_relatorio, arquivos_processados):

    ultimo_arquivo_ = aguardar_download_completo(pasta_downloads, arquivos_processados)
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    time.sleep(2)
    todos_arquivos = glob.glob(os.path.join(pasta_downloads, "*"))

    if todos_arquivos:
        arquivos_xls = [arquivo for arquivo in todos_arquivos if os.path.basename(arquivo).lower().startswith('listapedidoscompleto.xls')]
        if arquivos_xls:          
            for arquivo in arquivos_xls:
                codificacoes = ['utf-8', 'ISO-8859-1', 'cp1252']
                for cod in codificacoes:
                    try:
                        df = pd.read_html(arquivo, encoding=cod)[0]
                        break
                    except UnicodeDecodeError:
                        continue 

            novo_nome_xlsx = os.path.join(pasta_downloads, nome_relatorio + ".xlsx")
            df.to_excel(novo_nome_xlsx, index=False)
            os.remove(ultimo_arquivo_)
    else:
        print("Nenhum arquivo encontrado na pasta de downloads.")
    arquivos_processados.add(ultimo_arquivo_)

def atualizar_relatorios():

    drive.get('seu url')

    login_field = WebDriverWait(drive, 10).until(
        EC.element_to_be_clickable((By.XPATH, login_xpath)))
    login_field.send_keys(usuario_intranet)

    senha_field = WebDriverWait(drive, 10).until(
        EC.element_to_be_clickable((By.XPATH, senha_xpath)))
    senha_field.send_keys(senha_intranet)
    time.sleep(1)

    acessar_field = WebDriverWait(drive, 10).until(
        EC.element_to_be_clickable((By.XPATH, acessar_xpath)))
    acessar_field.click()
    time.sleep(5)

    for balde, link in links_baldes.items():
        drive.get(link)    
        converter_ultimo_download_para_xlsx(pasta_downloads, balde, arquivos_processados)
    
    time.sleep(3)
    drive.quit()

def unificar_relatorios(relatorio_1, relatorio_2):

    todos_arquivos = glob.glob(os.path.join(pasta_downloads, "*.xlsx"))
    ag_agendamento = "caminho da primeira base"
    rom_manual = "caminho da segunda base"
    arquivos_relatorio = [ag_agendamento, rom_manual]
    
    if arquivos_relatorio:
        df_unificado = pd.concat([pd.read_excel(arquivo) for arquivo in arquivos_relatorio], ignore_index=True)
        print(arquivos_relatorio)
        nome_arquivo = "Rom_manual_ag_agendamento.xlsx"
        caminho_arquivo = os.path.join(pasta_downloads, nome_arquivo)
        df_unificado.to_excel(caminho_arquivo, index=False)
    return print(arquivos_relatorio)

def obter_status(data):

    if isinstance(data, str):
        data = pd.to_datetime(data, errors='coerce', dayfirst=True)
    if pd.isna(data):
        return ""
    data_atual = pd.to_datetime('today').date()
    if data.date() > data_atual:
        mes_programado = data.strftime("%B").capitalize().upper()
        meses_pt = {'January': 'JANEIRO', 'February': 'FEVEREIRO', 'March': 'MARÇO', 'April': 'ABRIL',
                    'May': 'MAIO', 'June': 'JUNHO', 'July': 'JULHO', 'August': 'AGOSTO', 'September': 'SETEMBRO',
                    'October': 'OUTUBRO', 'November': 'NOVEMBRO', 'December': 'DEZEMBRO'}
        mes_pt = meses_pt.get(mes_programado, mes_programado)
        return f"PROGRAMADO {mes_pt}"
    else:
        return None

def calcular_dias_uteis(data_inicial):

    if pd.isna(data_inicial):
        return None
    feriados = cal.holidays(data_inicial.year)
    feriados = [feriado[0] for feriado in feriados]  
    return np.busday_count(data_inicial.date(), hoje, holidays=feriados)

def processar_grupo(grupo):

    grupo = grupo.sort_values('DATA_EMPENHO')
    valor_acumulado = 0.0
    data_empenho_400 = None

    for idx, linha in grupo.iterrows():
        valor_acumulado += linha['EMPENHO_VALOR']
        if valor_acumulado >= valor_minimo and data_empenho_400 is None:
            data_empenho_400 = linha['DATA_EMPENHO']
        grupo.loc[idx, 'data_empenho_400'] = data_empenho_400

    if data_empenho_400 is not None:
        grupo['LEADTIME_PEDIDO'] = grupo['data_empenho_400'].apply(calcular_dias_uteis)

    return grupo

def validar_valor_minimo(row):

    if row['UF'] in ufs_valor_minimo_600:
        min_value = 600
    else:
        min_value = 1200

    if row['EMPENHO_VALOR'] < min_value:
        return f'ABAIXO DO VALOR MÍNIMO: R$ {min_value}'
    return row['STATUS'] 

def definir_fifo(leadtime):

    if leadtime >= 10:
        return "FORA DO PRAZO"
    elif 8 <= leadtime <= 9:
        return "DENTRO DO PRAZO - AVISAR COMERCIAL"
    else:
        return "DENTRO DO PRAZO"

def excluir_base(base):

    if os.path.exists(base):
        os.remove(base)

def formatar_relatorio_final():

    # Aqui você pode formatar a base de acordo com as necessidades que tiver

    return 'resumo'

def formatar_resumo_para_email(resumo):

    resumo_html = "<table align='left' style='width: 30%; border-collapse: collapse; margin: 20px 0; font-family: Arial, sans-serif; border: 2px solid #007BFF; border-radius: 8px; overflow: hidden;'>"

    resumo_html += """
    <tr style='background-color: #007BFF; color: white;'>
        <th style='padding: 12px 10px; text-align: center; border: 1px solid #007BFF;'>Status</th>
        <th style='padding: 12px 10px; text-align: center; border: 1px solid #007BFF;'>Qtd. Pedidos</th>
        <th style='padding: 12px 10px; text-align: center; border: 1px solid #007BFF;'>Valor Total</th>
    </tr>
    """
    
    for _, linha in resumo.iterrows():
        resumo_html += f"""
        <tr>
            <td style='padding: 4px 10px; text-align: center; border: 1px solid #ccc;'>{linha['STATUS']}</td>
            <td style='padding: 4px 10px; text-align: center; border: 1px solid #ccc;'>{linha['QTD_PEDIDOS']}</td>
            <td style='padding: 4px 10px; text-align: center; border: 1px solid #ccc;'>{linha['VALOR_TOTAL']}</td>
        </tr>
        """
    
    resumo_html += "</table><br clear='all'/>" 
    resumo_html += "<p>Atenciosamente,<br>Pedro Gudryan</p>"
    
    return resumo_html

def mensagem_inicial():
    
    return "Olá, tudo bem?<br><br>sua mensagem.<br><br>"

def mensagem_final():

    return "<br><br>Atenciosamente,<br>seu nome"

def enviar_email(resumo):

    try:

        outlook = win32.Dispatch('outlook.application')
        

        mail = outlook.CreateItem(0)
        mail.Subject = "seu titulo" # Titulo do email
        email_body = f"<html><body>{mensagem_inicial()}<br><br>{formatar_resumo_para_email(resumo)}<br><br>{mensagem_final()}</body></html>" # Corpo do email 
        mail.HTMLBody = email_body
        destinatarios = [
            "destinatario1", "destinatario2"
        ]

        destinatarios_str = "; ".join(destinatarios)
        mail.To = destinatarios_str
        
        nome_arquivo = 'nome do seu arquivo com o tipo de arquivo'
        caminho_arquivo = os.path.join(pasta_downloads, nome_arquivo)
        
        
        if os.path.exists(caminho_arquivo):
            mail.Attachments.Add(caminho_arquivo)
        else:
            print(f"Arquivo não encontrado: {caminho_arquivo}")
            return False
        
        # Enviar o e-mail
        mail.Send()
        
        return True
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        return False

atualizar_relatorios()
unificar_relatorios('aguard_agendamento','romaneio_manual')
resumo = formatar_relatorio_final()
enviar_email(resumo)



