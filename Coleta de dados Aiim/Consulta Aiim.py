from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

# Instalar e inicializar o ChromeDriver
servico = Service(ChromeDriverManager().install())
opcoes = webdriver.ChromeOptions()
opcoes.add_argument('--headless')

driver = webdriver.Chrome(service=servico, options=opcoes)

# Carregar a planilha Excel
workbook = openpyxl.load_workbook('computadores.xlsx')  # Substitua 'computadores.xlsx' pelo nome do seu arquivo Excel
sheet = workbook.active

# Definir o tempo de espera máximo
wait = WebDriverWait(driver, 10)

# Definir a data atual
DATE = datetime.today()

# Funções para análise dos dados coletados
def salvar_dados(linha, dados):
    for col, valor in enumerate(dados, start=1):
        sheet.cell(row=linha, column=col).value = valor

# Função principal para buscar informações de computadores
def buscar_informacoes_computadores():
    linha_planilha = 2  # Começar a partir da linha 2 da planilha
    url_base = 'https://www.sitedecomputadores.com/'  # Exemplo de URL de site de computadores

    # Visitar o site
    driver.get(url_base)

    try:
        # Buscar um elemento específico (exemplo: nome do produto, preço, etc.)
        # Exemplo: esperar por um elemento de produto
        produtos = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.produto')))
        
        for produto in produtos:
            # Capturar o nome e preço do produto
            nome_produto = produto.find_element(By.CSS_SELECTOR, '.nome').text
            preco_produto = produto.find_element(By.CSS_SELECTOR, '.preco').text

            # Outros dados que podem ser coletados (descrição, especificações, etc.)
            descricao = produto.find_element(By.CSS_SELECTOR, '.descricao').text
            especificacoes = produto.find_element(By.CSS_SELECTOR, '.especificacoes').text
            
            # Salvar os dados na planilha
            salvar_dados(linha_planilha, [nome_produto, preco_produto, descricao, especificacoes])
            linha_planilha += 1  # Avançar para a próxima linha

    except Exception as e:
        print(f"Erro durante a coleta: {e}")

    finally:
        # Salvar o arquivo Excel atualizado
        workbook.save('Computadores_Coletados.xlsx')
        driver.quit()
        print("Processo de coleta concluído.")

# Executar a função
buscar_informacoes_computadores()