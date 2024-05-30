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

url = 'https://casadosdados.com.br/solucao/cnpj/pesquisa-avancada'
print("\nPreencha as informações para iniciar a extração.\n")

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
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def configurar_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--start-maximized')

    # Desativar extensões do Chrome
    chrome_options.add_argument('--disable-extensions')

    # Desativar a execução automática de reprodução de mídia com som
    chrome_options.add_argument(
        "--disable-blink-features=AutoplayIgnoreWebAudio")

    # Desativar o carregamento de imagens
    chrome_options.add_argument('--blink-settings=imagesEnabled=false')

    # Desativar o carregamento de fontes
    chrome_options.add_argument('--disable-gpu')

    chrome_options.add_argument('--disable-accelerated-2d-canvas')
    chrome_options.add_argument('--disable-accelerated-jpeg-decoding')

    # Desativar o carregamento de imagens de perfil
    chrome_options.add_argument('--disable-infobars')

    # Desativar sincronização de dados entre dispositivos
    chrome_options.add_argument('--disable-sync')

    # Desativar sugestões de preenchimento automático
    chrome_options.add_argument('--disable-autofill')

    # Desativar a execução automática de reprodução de mídia
    chrome_options.add_argument(
        "--disable-blink-features=MediaEngagementBypassAutoplayPolicies")

    # Esconde a barra de endereço (URL)
    chrome_options.add_argument('--kiosk')

    chrome_options.add_argument('--disable-site-isolation-trials')

    user_dir = os.path.expanduser("~")
    webdriver_path = os.path.join(
        user_dir, 'Desktop', 'extractor_cnpj', 'chromedriver.exe')
    os.environ['PATH'] = f'{os.environ["PATH"]};{os.path.abspath(webdriver_path)}'

    # Create the Chrome driver with the configured options
    driver = uc.Chrome(options=chrome_options,
                       executable_path=webdriver_path, headless=False)

    return driver


def WaitButton(driver, tempo_maximo=600):
    botao_clicado = False

    try:
        # Aguarde até que o botão seja clicável por até 30 segundos
        botao_clicado = WebDriverWait(driver, tempo_maximo).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, xpath_pesquisar))
        )

        # Adiciona um ouvinte de evento para detectar o clique do mouse
        driver.execute_script(f"""
            document.querySelector('{xpath_pesquisar}').addEventListener('click', function() {{
                window.buttonClicked = true;
            }});
        """)

        # Espera até que o tempo máximo seja atingido ou até que alguém clique no botão
        start_time = time.time()
        while time.time() - start_time < tempo_maximo:
            # Verifica se o botão foi clicado
            if driver.execute_script('return window.buttonClicked'):
                botao_clicado = True
                print(f"Extração iniciada às {data_formatada}")
                break

            # Aguarde um curto período antes de verificar novamente
            time.sleep(1)

    except TimeoutException:
        print(
            f"\nTempo máximo de espera ({tempo_maximo} segundos) atingido. Reinicie a sessão.")

    return botao_clicado


def extrair_dados():
    elementos_p = driver.find_elements(By.CSS_SELECTOR, '.box p')
    dados_temporarios = []

    for elemento in elementos_p:
        strongs = elemento.find_elements(By.TAG_NAME, 'strong')
        if len(strongs) >= 2:
            numero_cnpj = ''.join(c for c in strongs[0].text if c.isdigit())
            razao_social = strongs[1].text.strip().replace(
                ' ', '-').replace(',', '')
            dado = f'{razao_social}-{numero_cnpj}'
            dados_temporarios.append(dado)

    return dados_temporarios


# Função para clicar no botão da próxima página
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


def extrair_dados_de_todas_as_paginas():
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
        # Remova espaços e caracteres não numéricos do número de telefone
        telefone_formatado = ''.join(c for c in telefone.text if c.isdigit())

        # Adicione o prefixo internacional (55) se o número não começar com 55
        if not telefone_formatado.startswith('55'):
            telefone_formatado = '55' + telefone_formatado

        return telefone_formatado
    else:
        return ''


def pegar_informacoes():
    def encontrar_elemento(strategy, path):
        try:
            return WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((strategy, path)))
        except TimeoutException:
            return None

    razao_social = encontrar_elemento(By.CSS_SELECTOR, xpath_razao_social)
    nome_fantasia = encontrar_elemento(By.XPATH, xpath_nome_fantasia)
    cnpj = encontrar_elemento(By.XPATH, xpath_cnpj)
    telefone = encontrar_elemento(By.XPATH, xpath_telefone)
    telefone2 = encontrar_elemento(By.XPATH, xpath_telefone2)
    email = encontrar_elemento(By.CSS_SELECTOR, xpath_email)
    data_abertura = encontrar_elemento(By.CSS_SELECTOR, xpath_data_abertura)
    situacao_cadastral = encontrar_elemento(
        By.CSS_SELECTOR, xpath_situacao_cadastral)
    capital_social = encontrar_elemento(
        By.CSS_SELECTOR, xpath_capital_social) or "R$ 0"
    capital_social2 = encontrar_elemento(
        By.CSS_SELECTOR, xpath_capital_social2) or "R$ 0"
    cnae = encontrar_elemento(By.CSS_SELECTOR, xpath_cnae)
    mei = encontrar_elemento(By.CSS_SELECTOR, xpath_mei)
    mei2 = encontrar_elemento(By.CSS_SELECTOR, xpath_mei2)
    logradouro = encontrar_elemento(By.XPATH, xpath_logradouro)
    numero = encontrar_elemento(By.CSS_SELECTOR, xpath_numero)
    cep = encontrar_elemento(By.XPATH, xpath_cep)
    bairro = encontrar_elemento(By.XPATH, xpath_bairro)
    municipio = encontrar_elemento(By.CSS_SELECTOR, xpath_municipio)
    uf = encontrar_elemento(By.XPATH, xpath_uf)

    return razao_social, nome_fantasia, cnpj, telefone, telefone2, email, data_abertura, situacao_cadastral, capital_social, capital_social2, cnae, mei, mei2, logradouro, numero, cep, bairro, municipio, uf


def obter_dados(elemento):
    return str(elemento.text) if elemento else ''


def obter_dados_formatados(elemento, formatacao):
    return formatacao(elemento) if elemento else ''


# Loop para clicar na próxima página até que não haja mais páginas (Não tirar daqui, pois não funciona se tirar)
while clicar_proxima_pagina(xpath_proxima_pagina):
    time.sleep(3.5)
    pass

# Configuração do WebDriver
with configurar_driver() as driver:
    driver.get(url)

    # Aguarde até que o botão seja clicado ou tempo máximo atingido
    botao_clicado = WaitButton(driver)

    if botao_clicado:
        # Continue com o restante do seu código
        dados_armazenados = []
        if not dados_armazenados:
            dados_armazenados = []

        extrair_dados_de_todas_as_paginas()

        with open(resource_path('dados.csv'), 'w', newline='') as csvfile:
            fieldnames = ['empresa']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            for dado in dados_armazenados:
                row_dict = {'empresa': dado}
                writer.writerow(row_dict)
    else:
        print("\nBotão não foi clicado.")

# Nome do arquivo CSV
csv_file = 'dados.csv'

# Obtém o caminho absoluto usando a função resource_path
file_path = resource_path(csv_file)

# Leia os dados do arquivo CSV usando pandas
df = pd.read_csv(file_path)

time.sleep(1)

dados_armazenados_planilha = []

url_base = 'https://casadosdados.com.br/solucao/cnpj/'

for i, row in df.iterrows():

    item = row['empresa']

    url_completa = url_base + item

    driver.get(url_completa)

    try:
        data = datetime.now()
        data_atualizada = data.strftime('%H:%M:%S %d/%m/%y')

        print(
            f"\nExtraindo informações de: {item} às {data_atualizada}")
        # Verifique se os elementos estão presentes na página
        razao_social, nome_fantasia, cnpj, telefone, telefone2, email, data_abertura, situacao_cadastral, capital_social, capital_social2, cnae, mei, mei2, logradouro, numero, cep, bairro, municipio, uf = pegar_informacoes()

        # Mapeamento entre os elementos e suas formatações
        elementos_e_formatacoes = {
            'Razão Social': (razao_social, obter_dados),
            'Nome Fantasia': (nome_fantasia, obter_dados),
            'CNPJ': (cnpj, obter_dados),
            'Telefone': (telefone, formatar_telefone),
            'Telefone 2': (telefone2, formatar_telefone),
            'E-Mail': (email, obter_dados),
            'Data Abertura': (data_abertura, obter_dados),
            'Situação Cadastral': (situacao_cadastral, obter_dados),
            'Capital Social': (capital_social, obter_dados),
            'Capital Social 2': (capital_social2, obter_dados),
            'CNAE': (cnae, obter_dados),
            'MEI': (mei, obter_dados),
            'MEI 2': (mei2, obter_dados),
            'Logradouro': (logradouro, obter_dados),
            'Número': (numero, obter_dados),
            'CEP': (cep, obter_dados),
            'Bairro': (bairro, obter_dados),
            'Município': (municipio, obter_dados),
            'UF': (uf, obter_dados)
        }

        # Adiciona as informações à lista
        dados_armazenados_planilha.append({
            chave: obter_dados_formatados(elemento, formatacao)
            for chave, (elemento, formatacao) in elementos_e_formatacoes.items()
        })

    # Salva os dados a cada 5 linhas processadas
        if (i + 1) % 1 == 0:

            # Verifica se o diretório 'Extração' existe, se não, cria
            if not os.path.exists('Extração'):
                os.makedirs('Extração')

            excel_path = './Extração/Extração CNPJ.xlsx'

            # Verifica se o arquivo Excel já existe
            if os.path.exists(excel_path):
                # Se existir, lê o arquivo Excel e concatena com os novos dados
                df_existente = pd.read_excel(excel_path)
                df_final = pd.concat([df_existente, pd.DataFrame(
                    dados_armazenados_planilha)], ignore_index=True)
            else:
                df_final = pd.DataFrame(dados_armazenados_planilha)

            # Salva os dados no arquivo Excel
            df_final.to_excel(excel_path, index=False)
            dados_armazenados_planilha = []
            print(f"{i + 1} empresas foram salvas na planilha!")

    except Exception as e:
        print(
            f"\nErro ao extrair informações de: {url_completa} às {data_atualizada}")

print(
    f"Segundo e último processo de extração finalizado com sucesso às {data_atualizada}!")

try:
    driver.quit()
except Exception as e:
    print(f"Erro ao encerrar o WebDriver: {e}")

time.sleep(1)
# Encerre todos os processos do Google Chrome
for process in psutil.process_iter(['pid', 'name']):
    if 'chrome' in process.info['name'].lower():
        try:
            psutil.Process(process.info['pid']).terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

sys.exit()
