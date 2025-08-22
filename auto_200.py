"""
Módulo de automação para extração do sistema SSW
Script para download automático de relatórios 200 Manisfesto Operacionais
Realiza login, preenchimento das funções e realiza download do arquivo em excel

Requer:
- Arquivo credenciais.env com as variáveis do SSW
- Selenium WebDriver para Microsoft Edge
- Acesso a internet para interação com o sistema SSW
"""


from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import calendar
import pandas as pd
import os
from dotenv import load_dotenv
import time
import locale


# Código para setar as datas em português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8' )
except locale.Error:
    locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252' )

# Código para setar o diretório de dowloadod e carregamento de crendenciais
download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\DB_COMUM\\DB_200_manifestos')
load_dotenv('credenciais.env')

# Código que define o comando de realizar login no sistema SSW
# Para saber qual o id do campo da pagina, utilize o inspecionar elemento do navegador
def logar(driver, stop_event):
    if stop_event and stop_event.is_set():
        return
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))

    # Preenche os campos de login
    driver.find_element(By.NAME, "f1").send_keys(os.getenv("SSW_EMPRESA"))
    driver.find_element(By.NAME, "f2").send_keys(os.getenv("SSW_CNPJ"))
    driver.find_element(By.NAME, "f3").send_keys(os.getenv("SSW_USUARIO"))
    driver.find_element(By.NAME, "f4").send_keys(os.getenv("SSW_SENHA"))

    login_button = driver.find_element(By.ID, "5") # Localiza o botão de login
    driver.execute_script("arguments[0].click();", login_button) # Executa o clique no botão de login
    time.sleep(5)  # Aguarda o carregamento da página após o login

def manifestos_set (driver, data_inicio, data_fim, stop_event):
    if stop_event and stop_event.is_set():
        return
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f3").send_keys("200")  # Preenche o campo para redirecionar para o relatório de manifestos
    time.sleep(2)  # Aguarda o carregamento da janela
    
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])  # Muda para a aba do relatório de manifestos

    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.ID, "1"))) # Espera o primeiro campo de 'periodo de emissão'
    time.sleep(1)

    driver.find_element(By.ID, "1").send_keys(data_inicio)  # Preenche o primeiro campo de data
    time.sleep(3)

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "2")))  # Espera o segundo campo de data aparecer
    campo_data_fim = driver.find_element(By.ID, "2")
    driver.execute_script("arguments[0].value = '';", campo_data_fim)  # Clica no botão de pesquisa
    time.sleep(0.3)
    campo_data_fim.send_keys(data_fim)  # Preenche o segundo campo de data
    time.sleep(1)

    driver.find_element(By.ID, "11").clear()  # Direceciona ao campo 'tipo de arquivo' e limpa ele
    time.sleep(0.3)

    driver.find_element(By.ID, "11").send_keys("E")  # Preenche o campo 'tipo de arquivo' com 'E'
    time.sleep(5)

    #actions = ActionChains(driver)
    #actions.send_keys("7").perform()
    time.sleep(5)
def renomear_arquivo(pasta_download, nome_novo):
    # Código para renomear o arquivo mais recente na pasta de download
    try:
        # Cria a variavel 'arquivos' que contem todos os arquivos da pasta de download
        arquivos = [os.path.join(pasta_download, f)
                    for f in os.listdir(pasta_download)
                    if f.lower() != 'desktop.ini']
        
        if not arquivos:
            print("Nenhum arquivo encontrado na pasta de download.")
            return
        
        # Encontra o arquivo mais recente com base na data de criação
        
        arquivo_mais_recente = max(arquivos, key = os.path.getmtime)

        print(f"Arquivo mais recente encontrado: {os.path.basename(arquivo_mais_recente)}")
        _, extensao = os.path.splitext(arquivo_mais_recente)

        novo_nome_completo = os.path.join(pasta_download, nome_novo + extensao)

        if os.path.exists(novo_nome_completo):
            os.remove(novo_nome_completo)
    
        os.rename(arquivo_mais_recente, novo_nome_completo)
        print(f"Sucesso! Arquivo renomeado para: {os.path.basename(novo_nome_completo)}")
        return True
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return False
    
def main (stop_event = None):
    
    #Função principal que orquestra o processo de automação
    #Extrai os manifestos dos últimos 3 meses

    #Args:
    #    stop_event: Evento para sinalizar a parada da execução.
    

    hoje = datetime.now()

    edge_options = Options()
    edge_prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True
    }
    edge_options.add_experimental_option("prefs", edge_prefs)

    for i in range(3):
        if stop_event and stop_event.is_set():
            print("Sinal de parada recebido. Interrompendo a execução.")
            return
        
        driver = None

        
        #Código responsavel por calcular o primeiro e o último dia do mês
        #Assim como configurar as strings de data e nome do arquivo
        
        try:
            primeiro_dia_mes_atual = hoje.replace(day=1)
            ano = primeiro_dia_mes_atual.year
            mes = primeiro_dia_mes_atual.month - i

            if mes <= 1:
                mes += 12 + mes
                ano -= 1

            primeiro_dia = datetime(ano, mes, 1)

            
            if mes == datetime.now().month and ano == datetime.now().year:
                ultimo_dia = datetime.now()
            else:
                ultimo_dia_numero = calendar.monthrange(ano, mes)[1]
                ultimo_dia = datetime(ano, mes, ultimo_dia_numero)

            data_inicio_str = primeiro_dia.strftime('%d%m%y')
            data_fim_str = ultimo_dia.strftime('%d%m%y')

            nome_arquivo_mes = primeiro_dia.strftime('%b').upper() + str(ano)
            print(f"\n--- Iniciando o processo para o período {data_inicio_str} a {data_fim_str} ---\n")

            driver = webdriver.Edge(options=edge_options)

            # Código que chama a função de login no sistema SSW
            logar(driver, stop_event)
            if stop_event and stop_event.is_set(): raise InterruptedError

            # Código que chama a função de setar os manifestos
            manifestos_set(driver, data_inicio_str, data_fim_str, stop_event)
            if stop_event and stop_event.is_set(): raise InterruptedError

            # Código que chama a função de renomear o arquivo baixado
            renomear_arquivo(download_folder, nome_arquivo_mes)
            if stop_event and stop_event.is_set(): raise InterruptedError
            
        
        except InterruptedError:
            print("Execução interrompida pelo usuário.")
            break

        except Exception as e:
            print(f"Ocorreu um erro geral na automação: {mes} {ano}: {e}")

        finally:
            if driver:
                print("Encerrando a sessão do navegador.")
                driver.quit()

if __name__ == "__main__":
    main()
