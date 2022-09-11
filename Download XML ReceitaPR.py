from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options as ChromeOptions
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from time import sleep
from tqdm import tqdm
from pathlib import Path
from datetime import datetime, timedelta
from pyunpack import Archive
import os

solicitacao_df = pd.read_excel(r'Pastas Solicitações.xlsx')
chrome_options = ChromeOptions()
prefs = {'download.default_directory': r'COLOCAR CAMINHO PADRÃO PARA DOWNLOAD', 'profile.default_content_setting_values.automatic_downloads': 1}
chrome_options.add_experimental_option('prefs', prefs)
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.maximize_window()

hoje = datetime.now()
data = str(hoje - timedelta(days=20))
mes = data[5:7]
ano = data[0:4]


def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        sleep(1)
        dl_wait = False
        for name in os.listdir(path_to_downloads):
            if name.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


navegador.get('https://receita.pr.gov.br/')
navegador.find_element(By.ID, 'cpfusuario').send_keys('COLOCAR LOGIN AQUI') # Por questões de privacidade foi apagado o login
navegador.find_element(By.NAME, 'senha').send_keys('COLOCAR SENHA AQUI') # Por questões de privacidade foi apagado a senha
navegador.find_element(By.XPATH , '/html/body/div[2]/form[1]/div[4]/button').click()
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'fa-sign-out'))) # Espera a página carregar
navegador.find_element(By.XPATH, '//*[@id="menulateral"]/div/a[12]').click()
navegador.find_element(By.XPATH, '//*[@id="menulateral215"]/div/a').click()
navegador.find_element(By.XPATH, '//*[@id="menulateral238"]/div/a').click()
navegador.find_element(By.ID, 'menuLink750').click()
WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'fa-sign-out'))) # Espera a página carregar
navegador.find_element(By.ID, 'ext-gen1081').send_keys('26.899.016/0001-71')
navegador.find_element(By.CLASS_NAME, 'x-boundlist-item').click()
sleep(1)

local_final = Path(r'CAMINHO FINAL COM AS PASTAS DAS EMPRESAS QUE DESEJA SALVAR')
download = Path(r'COLOCAR CAMINHO PADRÃO PARA DOWNLOAD')
for i in tqdm(range(len(solicitacao_df))):
    cnpj = solicitacao_df['CNPJ'][i]
    cod = solicitacao_df['Nome pasta'][i]
    empresa = solicitacao_df['Razão Social'][i]
    navegador.find_element(By.ID, 'ext-gen1081').clear()
    sleep(1)
    navegador.find_element(By.ID, 'ext-gen1081').send_keys(cnpj)
    sleep(0.5)
    navegador.find_element(By.ID, 'ext-gen1081').send_keys(Keys.ENTER)
    sleep(0.5)
    navegador.find_element(By.ID, 'ucs20_ToolBarbtnAtualizar-btnInnerEl').click()
    sleep(0.5)
    for i in range(15):
        try:
            navegador.find_elements(By.CLASS_NAME, 'x-action-col-1')[i].click()
            download_wait(download)
        except:
            continue
    pasta_empresa = local_final / f'{empresa} - {cod}'
    pastas_empresa = pasta_empresa.iterdir()
    sleep(1)
    lista_pastas = [pastas.name for pastas in pastas_empresa]
    if f'{ano}.{mes}' in lista_pastas:
        pasta = (pasta_empresa / f'{ano}.{mes}')
    else:
        pasta = (pasta_empresa / f'{ano}.{mes}')
        pasta.mkdir()
    download = Path(r'COLOCAR CAMINHO PADRÃO PARA DOWNLOAD')
    arquivos = download.iterdir()
    sleep(1)
    for arquivo in arquivos:
        nome_arquivo = arquivo.name
        pasta_arquivo = (pasta / nome_arquivo)
        pasta_arquivo.mkdir()
        Archive(download / nome_arquivo).extractall(pasta_arquivo)
        os.remove(arquivo)
