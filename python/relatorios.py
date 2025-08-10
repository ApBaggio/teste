import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import schedule

# Função para baixar o relatório do Profig
def download_report():
    # Configurações do WebDriver
    driver = webdriver.Chrome()  # Altere para o WebDriver correspondente ao seu navegador
    driver.get("https://profig.institutogourmet.com/#/reports-central")

    # Login se necessário
    # login_field = driver.find_element(By.ID, "login_id")
    # password_field = driver.find_element(By.ID, "password_id")
    # login_field.send_keys("SEU_USUARIO")
    # password_field.send_keys("SUA_SENHA")
    # driver.find_element(By.ID, "login_button_id").click()

    # Navegando até o relatório desejado
    wait = WebDriverWait(driver, 10)
    report_link = wait.until(EC.element_to_be_clickable((By.XPATH, 'XPATH_DO_RELATORIO_DO_MES_ANTERIOR')))
    report_link.click()

    # Aguardar o download completar
    time.sleep(10)  # Tempo suficiente para o download completar

    # Fechar o navegador
    driver.quit()

    # Retornar o caminho do arquivo baixado
    download_path = os.path.join(os.path.expanduser("~"), "Downloads")  # Path padrão de downloads
    report_file = max([os.path.join(download_path, f) for f in os.listdir(download_path)], key=os.path.getctime)
    return report_file

# Função para atualizar a planilha de comissões
def update_commissions(report_file):
    # Carregar o relatório do Profig
    report_df = pd.read_excel(report_file)

    # Carregar a planilha de comissões existente
    commissions_file = '\\sgourmet\Compartilhada\Diretoria\RH\COMISSOES\COMISSAO FLORIANOPOLIS1.xlxs'
    commissions_df = pd.read_excel(commissions_file, sheet_name='Planilha1')

    # Atualizar a planilha de comissões com os dados do relatório
    updated_commissions_df = commissions_df.append(report_df, ignore_index=True)

    # Salvar a planilha atualizada
    with pd.ExcelWriter(commissions_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        updated_commissions_df.to_excel(writer, sheet_name='Planilha1', index=False)

# Função para executar a tarefa mensal
def job():
    report_file = download_report()
    update_commissions(report_file)



# Mantendo o script rodando
while True:
    schedule.run_pending()
    time.sleep(60)
