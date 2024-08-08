# Programa que acessa o portal da Sefaz e faz download do arquivo CSV para a pasta Downloads
# Devolve as notas geradas desde o início do mês ATUAL até ontem
import os
import pandas as pd
import pyautogui
import time
import tkinter as tk
from io import StringIO
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By


# Função para renomear os arquivos baixados
def renomeia_arquivo(download_dir, novo_nome):
    # Exclui o arquivo caso ele já exista
    if os.path.exists(os.path.join(download_dir, novo_nome)):
        os.remove(os.path.join(download_dir, novo_nome))

    caminho_antigo = os.path.join(download_dir, 'arquivo.csv')
    novo_caminho = os.path.join(download_dir, novo_nome)
    os.rename(caminho_antigo, novo_caminho)


# Função para acessar a página de download dos CSVs
def acessa_portal(driver):
    # Lidando com o pop-up
    pyautogui.click(1151, 376)  # Clica em 'Acessar o DT-e'
    time.sleep(2)  # Aguarda 2 segundos para o popup aparecer
    pyautogui.click(1134, 403)  # Clica em 'OK', presumindo que o certificado correto já está selecionado

    # Navegando no DT-e
    driver.find_element(By.XPATH,  # Clica no CNPJ correto
                        '//*[text()="06.201.124-3"]').click()

    driver.find_element(By.XPATH,  # Clica em 'NF-e'
                        '//*[@id="areaPainel"]/tbody/tr[1]/td/div/div[1]/div[4]').click()

    driver.find_element(By.XPATH,  # Clica em 'NF-e - Download de Arquivos'
                        '//*[@id="areaPainel"]/tbody/tr[1]/td/div/div[2]/div[5]/div[9]').click()


# Função para baixar os CSVs
def baixa_csv(driver, data1, data2):
    caixa_data1 = driver.find_element(By.XPATH,  # Seleciona a caixa para fazer o input da data 1
                                      '//*[@id="datepicker_de"]')
    caixa_data1.send_keys(data1)

    caixa_data2 = driver.find_element(By.XPATH,  # Seleciona a caixa para fazer o input da data 2
                                      '//*[@id="datepicker_ate"]')
    caixa_data2.send_keys(data2)

    driver.find_element(By.XPATH,  # Clica em 'Solicitar'
                        '/html/body/table[2]/tbody/tr[1]/td/form/table/tbody/tr[23]/td[1]/input').click()

    driver.find_element(By.XPATH,  # Clica em 'Exportar para CSV
                        '/html/body/table[2]/tbody/tr[1]/td/table[2]/tbody/tr[2]/td[1]/input').click()

    time.sleep(3)  # Aguarda o download do arquivo CSV


# Função para adicionar o delimitador ';' que falta após o nome da última coluna no CSV de CT-es
def ajusta_csv(caminho):
    # Abre o arquivo CSV como texto
    with open(caminho, 'r', encoding='latin1') as f:
        # Lê o conteúdo do arquivo
        text = f.read()

    # Divide o texto em linhas
    lines = text.splitlines()

    # Adiciona o caractere de delimitador após o nome da última coluna no cabeçalho
    lines[1] = lines[1] + ';'

    # Junta as linhas novamente em um único texto
    text = '\n'.join(lines)

    # Cria um objeto StringIO a partir do texto
    csv_file = StringIO(text)

    # Lê o arquivo CSV a partir do objeto StringIO
    df = pd.read_csv(csv_file, encoding='latin1', delimiter=';', skiprows=1)

    return df


# Função para receber as datas como input
def get_data():
    # Cria uma janela
    window = tk.Tk()

    # Cria dois input box
    entr1 = tk.Entry(window, width=50)
    entr2 = tk.Entry(window, width=50)

    # Faz o texto indicativo
    entr1.insert(0, "Digite a data de início no formato dd/mm/aaaa")
    entr2.insert(0, "Digite a data final no formato dd/mm/aaaa")

    # Função para limpar o texto indicativo quando a caixa de entrada recebe o foco
    def clear_entry(event, entry):
        entry.delete(0, tk.END)

    # Adiciona o evento <FocusIn> às caixas de entrada
    entr1.bind("<FocusIn>", lambda event: clear_entry(event, entr1))
    entr2.bind("<FocusIn>", lambda event: clear_entry(event, entr2))

    # Posiciona as caixas
    entr1.pack()
    entr2.pack()

    datas = [None, None]

    def botao_ok():
        datas[0] = entr1.get()
        datas[1] = entr2.get()

        window.destroy()

    ok = tk.Button(window, text="OK", command=botao_ok)
    ok.pack()

    # Ajusta a posição da janela para o centro da tela
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    window.mainloop()

    return datas


datas = get_data()

d1 = datas[0]
d2 = datas[1]

# Define a pasta de download do arquivo como 'sefaz'
pasta_atual = os.getcwd()
download_dir = os.path.dirname(os.path.dirname(pasta_atual))

options = Options()
options.add_experimental_option("prefs", {
  "download.default_directory": download_dir,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

# Cria o objeto driver

dr = webdriver.Chrome(options=options)
dr.maximize_window()  # Deixa a janela em fullscreen para garantir que as coordenadas do clique estão certas
dr.get('http://online.sefaz.am.gov.br/inicioDte.asp')

acessa_portal(dr)

dr.find_element(By.XPATH,   # Seleciona NF-e
                '/html/body/table[2]/tbody/tr[1]/td/form/table/tbody/tr[8]/td[2]/input[1]').click()

baixa_csv(dr, d1, d2)

renomeia_arquivo(download_dir, 'planilha_nfe.csv')

dr.find_element(By.XPATH,   # Clica em 'Voltar'
                '/html/body/table[2]/tbody/tr[1]/td/table[2]/tbody/tr[4]/td[3]/input').click()

dr.find_element(By.XPATH,   # Seleciona CT-e
                '/html/body/table[2]/tbody/tr[1]/td/form/table/tbody/tr[8]/td[2]/input[3]').click()

baixa_csv(dr, d1, d2)

renomeia_arquivo(download_dir, 'planilha-cte.csv')

# Transforma os CSVs em Dataframes
caminho_nfe = os.path.join(download_dir, 'planilha_nfe.csv')
caminho_cte = os.path.join(download_dir, 'planilha-cte.csv')

df_nfe = pd.read_csv(caminho_nfe, encoding='latin1', delimiter=';', skiprows=1)
df_cte = ajusta_csv(caminho_cte)

os.remove(caminho_nfe)
os.remove(caminho_cte)

# Define o caminho da planilha final
caminho_planilha = os.path.join(download_dir, 'planilha_notas.xlsx')

writer = pd.ExcelWriter(caminho_planilha, engine='xlsxwriter')

# Salva os dois CSVs na mesma planilha em abas diferentes
df_nfe.to_excel(writer, sheet_name='Planilha NFe', index=False)
df_cte.to_excel(writer, sheet_name='Planilha CTe', index=False)

writer._save()
