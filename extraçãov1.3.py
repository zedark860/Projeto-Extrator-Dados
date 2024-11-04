from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
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
import os
import pandas as pd
import undetected_chromedriver as uc

LOGGER.setLevel(logging.WARNING)
ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

class DataExtractor:
    def _init_(self, url, url_base, xpath_proxima_pagina):
        self.url = url
        self.url_base = url_base
        self.xpath_proxima_pagina = xpath_proxima_pagina
        self.driver = self.configurar_driver()
        self.data_formatada = datetime.now().strftime('%H:%M:%S %d/%m/%y')
        self.interromper_extracao = False
        self.xpaths = {
            'pesquisar': 'a.button:nth-child(1)',
            'cnpj': '//label[contains(text(), "CNPJ")]/following::p[1]',
            'razao_social': '//label[contains(text(), "Razão Social")]/following::p[1]',
            'socio': '//label[contains(text(), "Sócios")]/following::p[1]',
            'nome_fantasia': '//label[contains(text(), "Nome Fantasia")]/following::p[1]',
            'telefone': '//label[contains(text(), "Telefone")]/following::p[1]',
            'telefone2': '//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[20]/p/a[1]',
            'email': '//label[contains(text(), "Email")]/following::p[1]',
            'data_abertura': '//label[contains(text(), "Data de Abertura")]/following::p[1]',
            'situacao_cadastral': '//label[contains(text(), "Situação Cadastral")]/following::p[1]',
            'capital_social': '//label[contains(text(), "Capital Social")]/following::p[1]',
            'capital_social2': '//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[10]/p',
            'cnae': '//label[contains(text(), "CNAE Principal")]/following::p[1]',
            'mei': '//label[contains(text(), "Empresa MEI")]/following::p[1]',
            'mei2': 'id("__nuxt")/DIV[1]/SECTION[4]/DIV[2]/DIV[1]/DIV[1]/DIV[9]/P[1]',
            'logradouro': '//label[contains(text(), "Logradouro")]/following::p[1]',
            'numero': '//label[contains(text(), "Número")]/following::p[1]',
            'cep': '//label[contains(text(), "CEP")]/following::p[1]',
            'bairro': '//label[contains(text(), "Bairro")]/following::p[1]',
            'municipio': '//label[contains(text(), "Municipio")]/following::p[1]',
            'uf': '//label[contains(text(), "Estado")]/following::p[1]'
        }

    def configurar_driver(self):
        chrome_options = self.configurar_opcoes_chrome()
        webdriver_path = self.obter_caminho_webdriver()
        try:
            driver = uc.Chrome(options=chrome_options, executable_path=webdriver_path, headless=False)
            return driver
        except Exception as e:
            print(f"Erro ao configurar o driver: {e}")
            return None

    @staticmethod
    def configurar_opcoes_chrome():
        chrome_options = webdriver.ChromeOptions()
        arguments = [
            '--disable-extensions',
            '--disable-blink-features=AutoplayIgnoreWebAudio',
            '--blink-settings=imagesEnabled=false',
            '--disable-gpu',
            '--disable-accelerated-2d-canvas',
            '--disable-accelerated-jpeg-decoding',
            '--disable-infobars',
            '--disable-sync',
            '--disable-autofill',
            '--disable-blink-features=MediaEngagementBypassAutoplayPolicies',
            '--kiosk',
            '--disable-site-isolation-trials'
        ]
        for arg in arguments:
            chrome_options.add_argument(arg)
        return chrome_options

    @staticmethod
    def obter_caminho_webdriver():
        user_dir = os.path.expanduser("~")
        webdriver_path = os.path.join(user_dir, 'Desktop', 'chromedriver.exe')
        os.environ['PATH'] = f'{os.environ["PATH"]};{os.path.abspath(webdriver_path)}'
        return webdriver_path

    def tempo_espera(self, driver, tempo_maximo):
        start_time = time.time()
        while time.time() - start_time < tempo_maximo:
            if driver.execute_script('return window.buttonClicked'):
                botao_clicado = True
                print(f"Extração iniciada às {self.data_formatada}")
                return botao_clicado
            time.sleep(1)

    def WaitButton(self, driver, tempo_maximo=600):
        botao_clicado = False
        try:
            botao_clicado = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, self.xpaths['pesquisar']))
            )
            driver.execute_script(f"""
                document.querySelector('{self.xpaths['pesquisar']}').addEventListener('click', function() {{
                    window.buttonClicked = true;
                }});
            """)
            self.tempo_espera(driver, tempo_maximo)
        except TimeoutException:
            print(f"\nTempo máximo de espera ({tempo_maximo} segundos) atingido. Reinicie a sessão.")
        return botao_clicado

    def extrair_dados(self):
        elementos_p = self.driver.find_elements(By.CSS_SELECTOR, '.box p')
        dados_temporarios = []
        for elemento in elementos_p:
            strongs = elemento.find_elements(By.TAG_NAME, 'strong')
            if len(strongs) >= 2:
                numero_cnpj = ''.join(c for c in strongs[0].text if c.isdigit())
                razao_social = strongs[1].text.strip().replace(' ', '-').replace(',', '')
                dado = f'{razao_social}-{numero_cnpj}'
                dados_temporarios.append(dado)
        return dados_temporarios

    def clicar_proxima_pagina(self):
        try:
            botao_proxima_pagina = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, self.xpath_proxima_pagina)))
            if botao_proxima_pagina and 'is-disabled' not in botao_proxima_pagina.get_attribute('class'):
                self.driver.execute_script("arguments[0].click();", botao_proxima_pagina)
                time.sleep(1)
                return True
        except Exception as e:
            time.sleep(2)
            return False

    def extrair_dados_de_todas_as_paginas(self, dados_armazenados):
        while True:
            dados_da_pagina = self.extrair_dados()
            if dados_da_pagina:
                dados_armazenados.extend(dados_da_pagina)
                pode_prosseguir = self.clicar_proxima_pagina()
                time.sleep(1)
                if not pode_prosseguir:
                    time.sleep(1)
                    print(f"Primeira extração concluída! {self.data_formatada}\nIniciando a segunda...")
                    break

    @staticmethod
    def formatar_telefone(telefone):
        if telefone:
            telefone_formatado = ''.join(c for c in telefone.text if c.isdigit())
            if not telefone_formatado.startswith('55'):
                telefone_formatado = '55' + telefone_formatado
            return telefone_formatado
        else:
            return ''

    @staticmethod
    def encontrar_elemento(driver, strategy, path):
        try:
            return WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((strategy, path)))
        except TimeoutException:
            return None

    def pegar_informacoes(self, driver):
        elementos = []
        for strategy, path in [
            (By.XPATH, self.xpaths['razao_social']),
            (By.XPATH, self.xpaths['cnpj']),
            (By.XPATH, self.xpaths['socio']),
            (By.XPATH, self.xpaths['nome_fantasia']),
            (By.XPATH, self.xpaths['telefone']),
            (By.XPATH, self.xpaths['telefone2']),
            (By.XPATH, self.xpaths['email']),
            (By.XPATH, self.xpaths['data_abertura']),
            (By.XPATH, self.xpaths['situacao_cadastral']),
            (By.XPATH, self.xpaths['capital_social']),
            (By.XPATH, self.xpaths['capital_social2']),
            (By.XPATH, self.xpaths['cnae']),
            (By.XPATH, self.xpaths['mei']),
            (By.XPATH, self.xpaths['mei2']),
            (By.XPATH, self.xpaths['logradouro']),
            (By.XPATH, self.xpaths['numero']),
            (By.XPATH, self.xpaths['cep']),
            (By.XPATH, self.xpaths['bairro']),
            (By.XPATH, self.xpaths['municipio']),
            (By.XPATH, self.xpaths['uf'])
        ]:
            elemento = self.encontrar_elemento(driver, strategy, path)
            if strategy == By.XPATH and path in [self.xpaths['telefone'], self.xpaths['telefone2']]:
                elementos.append(self.formatar_telefone(elemento))
            else:
                elementos.append(elemento.text if elemento else '')
        return elementos

    def salvar_informacoes(self, informacoes, dados_extraidos, contador):
        try:
            caminho_pasta = os.getcwd() + '\\Extracão'
            if not os.path.isdir(caminho_pasta):
                os.makedirs(caminho_pasta)
            
            nome_arquivo = f'{caminho_pasta}\\dados_extraidos_{self.data_formatada.replace(":", "_").replace("/", "-")}.xlsx'
            if not os.path.exists(nome_arquivo):
                header = ['Razão Social', 'CNPJ', 'Sócio', 'Nome Fantasia', 'Telefone', 'Telefone 2', 'E-mail',
                        'Data de Abertura', 'Situação Cadastral', 'Capital Social', 'Capital Social 2', 'CNAE', 'MEI', 'MEI 2',
                        'Logradouro', 'Número', 'CEP', 'Bairro', 'Município', 'UF']
                pd.DataFrame([informacoes], columns=header).to_excel(nome_arquivo, index=False)
            else:
                workbook = load_workbook(nome_arquivo)
                sheet = workbook.active
                sheet.append(informacoes)
                workbook.save(nome_arquivo)
            dados_extraidos.append(informacoes)
            print(f"\n{contador + 1} empresas foram salvas na planilha às {self.data_formatada}. (Empresa {informacoes[0]})\n")
        except Exception as e:
            print(f"Erro ao salvar empresa {informacoes[0]}: {e}")

    def extrair_e_salvar_dados(self, driver, urls_armazenados):
        print(f"Extraindo dados detalhados de {len(urls_armazenados)} CNPJs...")
        dados_extraidos = []
        for contador, dado in enumerate(urls_armazenados):
            url = self.url_base + dado
            driver.get(url)
            informacoes = self.pegar_informacoes(driver)
            self.salvar_informacoes(informacoes, dados_extraidos, contador)

    def main(self):
        try:
            driver = self.driver
            driver.get(self.url)

            if self.WaitButton(driver):
                urls_armazenados = []
                self.extrair_dados_de_todas_as_paginas(urls_armazenados)
                self.extrair_e_salvar_dados(driver, urls_armazenados)
            else:
                print("Falha na espera do botão. A extração não foi iniciada.")
                
            driver.quit()
        except Exception as e:
            print(f"Erro inesperado: {e}")

if _name_ == '_main_':
    url = 'https://casadosdados.com.br/solucao/cnpj/pesquisa-avancada'
    url_base = 'https://casadosdados.com.br/solucao/cnpj/'
    xpath_proxima_pagina = 'pagination-next'
    extractor = DataExtractor(url, url_base, xpath_proxima_pagina)
    extractor.main()

    ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
