from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.remote_connection import LOGGER
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from auto_download_undetected_chromedriver import download_undetected_chromedriver
from openpyxl import load_workbook
from datetime import datetime
import ctypes
import time
import logging
import csv
import os
import pandas as pd
import undetected_chromedriver as uc
import psutil
import sys

LOGGER.setLevel(logging.WARNING)
ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

data = datetime.now()
data_formatada = data.strftime('%H:%M:%S %d/%m/%y')

interromper_extracao = False

url = 'https://casadosdados.com.br/solucao/cnpj/pesquisa-avancada'
url_base = 'https://casadosdados.com.br/solucao/cnpj/'

xpath_pesquisar = 'a.button:nth-child(1)'
xpath_proxima_pagina = 'pagination-next'
xpath_cnpj = '//*[@id="__nuxt"]/div/div[2]/section/div/div/div[4]/div/div/div[1]/p[2]'
xpath_razao_social = '.columns:nth-child(1) > .column:nth-child(2) > p:nth-child(2)'
xpath_nome_fantasia = '//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[1]/div[3]/p[2]'
xpath_telefone = 'id("__nuxt")/DIV[1]/DIV[2]/SECTION[1]/DIV[1]/DIV[1]/DIV[4]/DIV[1]/DIV[3]/DIV[1]/P[2]/A[1]'
xpath_telefone2 = '//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[3]/div[1]/p[3]/a'
xpath_email = '.column:nth-child(2) > p > a'
xpath_data_abertura = '.column > a'
xpath_situacao_cadastral = '.columns:nth-child(1) > .column:nth-child(5) > p:nth-child(2)'
xpath_capital_social = '.columns:nth-child(1) > .column:nth-child(7) > p:nth-child(2)'
xpath_capital_social2 = '.column:nth-child(8) > p:nth-child(2)'
xpath_cnae = '.columns:nth-child(5) > .column:nth-child(1) > p:nth-child(2)'
xpath_mei = '.column:nth-child(10) > p:nth-child(2)'
xpath_mei2 = '.column:nth-child(9) > p:nth-child(2)'
xpath_logradouro = 'id("__nuxt")/DIV[1]/DIV[2]/SECTION[1]/DIV[1]/DIV[1]/DIV[4]/DIV[1]/DIV[2]/DIV[1]/P[2]'
xpath_numero = '.columns:nth-child(2) > .column:nth-child(2) > p:nth-child(2)'
xpath_cep = 'id("__nuxt")/DIV[1]/DIV[2]/SECTION[1]/DIV[1]/DIV[1]/DIV[4]/DIV[1]/DIV[2]/DIV[4]/P[2]'
xpath_bairro = 'id("__nuxt")/DIV[1]/DIV[2]/SECTION[1]/DIV[1]/DIV[1]/DIV[4]/DIV[1]/DIV[2]/DIV[5]/P[2]'
xpath_municipio = '.column:nth-child(6) > p > a'
xpath_uf = '//*[@id="__nuxt"]/div/div[2]/section[1]/div/div/div[4]/div[1]/div[2]/div[7]/p[2]/a'


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def limpar_tela():
    time.sleep(3)
    os.system('cls' if os.name == "nt" else "clear")


def configurar_driver():
    chrome_options = configurar_opcoes_chrome()
    webdriver_path = obter_caminho_webdriver()
    driver = uc.Chrome(options=chrome_options, executable_path=webdriver_path, headless=False)
    return driver


def configurar_opcoes_chrome():
    chrome_options = webdriver.ChromeOptions()
    arguments = [
        '--start-maximized',
        '--disable-extensions',
        "--disable-blink-features=AutoplayIgnoreWebAudio",
        '--blink-settings=imagesEnabled=false',
        '--disable-gpu',
        '--disable-accelerated-2d-canvas',
        '--disable-accelerated-jpeg-decoding',
        '--disable-infobars',
        '--disable-sync',
        '--disable-autofill',
        "--disable-blink-features=MediaEngagementBypassAutoplayPolicies",
        '--kiosk',
        '--disable-site-isolation-trials'
    ]
    for arg in arguments:
        chrome_options.add_argument(arg)
    return chrome_options


def obter_caminho_webdriver():
    user_dir = os.path.expanduser("~")
    webdriver_path = os.path.join(
        user_dir, 'Desktop', 'chromedriver.exe')
    os.environ['PATH'] = f'{os.environ["PATH"]};{os.path.abspath(webdriver_path)}'
    return webdriver_path


def tempo_espera(driver, tempo_maximo):
    start_time = time.time()
    while time.time() - start_time < tempo_maximo:
        if driver.execute_script('return window.buttonClicked'):
            botao_clicado = True
            print(f"Extração iniciada às {data_formatada}")
            return botao_clicado
        time.sleep(1)
        
        
def WaitButton(driver, tempo_maximo=600):
    botao_clicado = False
    try:
        botao_clicado = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, xpath_pesquisar))
        )
        driver.execute_script(f"""
            document.querySelector('{xpath_pesquisar}').addEventListener('click', function() {{
                window.buttonClicked = true;
            }});
        """)
        tempo_espera(driver, tempo_maximo)
    except TimeoutException:
        print(
            f"\nTempo máximo de espera ({tempo_maximo} segundos) atingido. Reinicie a sessão.")
    return botao_clicado


def extrair_dados():
    elementos_p = driver.find_elements(By.CSS_SELECTOR, '.box p')
    dados_temporarios = list()
    for elemento in elementos_p:
        strongs = elemento.find_elements(By.TAG_NAME, 'strong')
        if len(strongs) >= 2:
            numero_cnpj = ''.join(c for c in strongs[0].text if c.isdigit())
            razao_social = strongs[1].text.strip().replace(
                ' ', '-').replace(',', '')
            dado = f'{razao_social}-{numero_cnpj}'
            dados_temporarios.append(dado)
    return dados_temporarios


def clicar_proxima_pagina(xpath_proxima_pagina):
    try:
        botao_proxima_pagina = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CLASS_NAME, xpath_proxima_pagina)))

        if botao_proxima_pagina and 'is-disabled' not in botao_proxima_pagina.get_attribute('class'):
            driver.execute_script("arguments[0].click();", botao_proxima_pagina)
            time.sleep(2)
            return True
    except Exception as e:
        time.sleep(3)
        return False


def extrair_dados_de_todas_as_paginas(dados_armazenados):
    while True:
        dados_da_pagina = extrair_dados()
        if dados_da_pagina:
            dados_armazenados.extend(dados_da_pagina)
            pode_prosseguir = clicar_proxima_pagina(xpath_proxima_pagina)
            time.sleep(1)
            if not pode_prosseguir:
                time.sleep(3)
                print(
                    f"Primeira extração concluída! {data_formatada}\nIniciando a segunda...")
                break


def formatar_telefone(telefone):
    if telefone:
        telefone_formatado = ''.join(c for c in telefone.text if c.isdigit())
        if not telefone_formatado.startswith('55'):
            telefone_formatado = '55' + telefone_formatado
        return telefone_formatado
    else:
        return ''


def encontrar_elemento(driver,strategy, path):
    try:
        return WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((strategy, path)))
    except TimeoutException:
        return None


def pegar_informacoes(driver):
    elementos = list()
    for strategy, path in [
        (By.CSS_SELECTOR, xpath_razao_social),
        (By.XPATH, xpath_nome_fantasia),
        (By.XPATH, xpath_cnpj),
        (By.XPATH, xpath_telefone),
        (By.XPATH, xpath_telefone2),
        (By.CSS_SELECTOR, xpath_email),
        (By.CSS_SELECTOR, xpath_data_abertura),
        (By.CSS_SELECTOR, xpath_situacao_cadastral),
        (By.CSS_SELECTOR, xpath_capital_social) or "R$ 0",
        (By.CSS_SELECTOR, xpath_capital_social2) or "R$ 0",
        (By.CSS_SELECTOR, xpath_cnae),
        (By.CSS_SELECTOR, xpath_mei),
        (By.CSS_SELECTOR, xpath_mei2),
        (By.XPATH, xpath_logradouro),
        (By.CSS_SELECTOR, xpath_numero),
        (By.XPATH, xpath_cep),
        (By.XPATH, xpath_bairro),
        (By.CSS_SELECTOR, xpath_municipio),
        (By.XPATH, xpath_uf)
    ]:
    
        try:
            elemento = encontrar_elemento(driver, strategy, path) 
            elementos.append(elemento)
        except Exception as e:
            elementos.append(None)
            
    return tuple(elementos)


def obter_dados(elemento):
    return str(elemento.text) if elemento else ''


def obter_dados_formatados(elemento, formatacao):
    return formatacao(elemento) if elemento else ''


def coletar_dados(dados_armazenados):
    extrair_dados_de_todas_as_paginas(dados_armazenados)

    with open(resource_path('dados.csv'), 'w', newline='') as csvfile:
        fieldnames = ['empresa']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for dado in dados_armazenados:
            row_dict = {'empresa': dado}
            writer.writerow(row_dict)
   
            
def localizar_planilha():
    csv_file = 'dados.csv'
    file_path = resource_path(csv_file)
    df = pd.read_csv(file_path)
    return df


def configurar_driver_e_navegar(url, xpath_proxima_pagina):
    while clicar_proxima_pagina(xpath_proxima_pagina):
        time.sleep(3.5)

    with configurar_driver() as driver:
        driver.get(url)

        botao_clicado = WaitButton(driver)

        if botao_clicado:
            return driver
        else:
            print("\nBotão não foi clicado.")
            return None
        
        
def extrair_informacoes_empresa(driver, url_base, item):
    url_completa = url_base + item
    driver.get(url_completa)
    time.sleep(1.5)
    limpar_tela()
    print(f"\nExtraindo informações de: {item}")
    return url_completa, pegar_informacoes(driver)


def processar_informacoes(elementos):
    elementos_e_formatacoes = {
        'Razão Social': (elementos[0], obter_dados),
        'Nome Fantasia': (elementos[1], obter_dados),
        'CNPJ': (elementos[2], obter_dados),
        'Telefone': (elementos[3], formatar_telefone),
        'Telefone 2': (elementos[4], formatar_telefone),
        'E-Mail': (elementos[5], obter_dados),
        'Data Abertura': (elementos[6], obter_dados),
        'Situação Cadastral': (elementos[7], obter_dados),
        'Capital Social': (elementos[8], obter_dados),
        'Capital Social 2': (elementos[9], obter_dados),
        'CNAE': (elementos[10], obter_dados),
        'MEI': (elementos[11], obter_dados),
        'MEI 2': (elementos[12], obter_dados),
        'Logradouro': (elementos[13], obter_dados),
        'Número': (elementos[14], obter_dados),
        'CEP': (elementos[15], obter_dados),
        'Bairro': (elementos[16], obter_dados),
        'Município': (elementos[17], obter_dados),
        'UF': (elementos[18], obter_dados)
    }
    return {
        chave: obter_dados_formatados(elemento, formatacao)
        for chave, (elemento, formatacao) in elementos_e_formatacoes.items()
    }
    
def salvar_dados_em_excel(dados_armazenados_planilha, excel_path):
    if not os.path.exists('Extração'):
        os.makedirs('Extração')

    if os.path.exists(excel_path):
        df_existente = pd.read_excel(excel_path)
        df_final = pd.concat([df_existente, pd.DataFrame(
            dados_armazenados_planilha)], ignore_index=True)
    else:
        df_final = pd.DataFrame(dados_armazenados_planilha)

    df_final.to_excel(excel_path, index=False)
    
def processar_empresas(driver, df, url_base):
    dados_armazenados_planilha = list()
    
    for i, row in df.iterrows():

        item = row['empresa']

        try:
            url_completa, elementos = extrair_informacoes_empresa(driver, url_base, item)
            dados_formatados = processar_informacoes(elementos)
            dados_armazenados_planilha.append(dados_formatados)

            excel_path = './Extração/Extração CNPJ.xlsx'
            salvar_dados_em_excel(dados_armazenados_planilha, excel_path)
            dados_armazenados_planilha = list()

            print(f"{i + 1} empresas foram salvas na planilha!")
                
        except Exception as e:
            print(
                f"\nErro ao extrair informações de: {url_completa}", e)
            

def encerrar_processo_driver(driver):
    try:
        driver.quit()
    except Exception as e:
        print(f"Erro ao encerrar o WebDriver: {e}")
        
        
def encerrar_processos_google():
    for process in psutil.process_iter(['pid', 'name']):
        if 'chrome' in process.info['name'].lower():
            try:
                psutil.Process(process.info['pid']).terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass


if __name__ == "__main__":
    print("\nPreencha as informações para iniciar a extração.\n")

    driver = configurar_driver_e_navegar(url, xpath_proxima_pagina)
    
    time.sleep(2)

    if driver:
        coletar_dados(dados_armazenados=list())

    df = localizar_planilha()

    time.sleep(1)

    processar_empresas(driver, df, url_base)
            
    time.sleep(2)
    
    limpar_tela()

    print(
        f"Extração finalizada com sucesso!")

    encerrar_processo_driver(driver)

    time.sleep(1)

    encerrar_processos_google()

    sys.exit()