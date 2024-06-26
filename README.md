# Automacão Palin
Automação para equipe do Pré-Vendas da Palin &amp; Martins
## Automação de Extração de Dados


Este script Python foi desenvolvido para automatizar o processo de extração de dados definidas. Ele utiliza a biblioteca Selenium para interagir com o navegador e openpyxl para manipular planilhas Excel. O objetivo é extrair informações desse site.

### Pré-requisitos

1. [Python 3.x](https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe)
2. [documentação Python](https://docs.python.org/pt-br/3/tutorial/)
3. [Vs Code](https://code.visualstudio.com/)
4. Bibliotecas a ser baixadas no Visual Code
5. [Selenium](https://selenium-python.readthedocs.io/)
6. [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
7. [WebDriver para o navegador Chrome](https://www.selenium.dev/pt-br/documentation/webdriver/)

### Instalação

1. Clone o repositório para sua máquina local ou virtual para rodar em segndo plano.
2. Instale as dependências executando o comando `pip install openpyxl selenium`.
3. Certifique-se de ter o WebDriver para o navegador Chrome instalado e configurado no seu PATH.

### Utilização

1. Adicione os números que deseja pesquisar na planilha Excel 'nome_definido_por_vc.xlsx'.
2. Não esqueca de formatar o o jeito da busca Ex: `"9999999" sem virgulas, pontos e o mesmo serve para as letras, Ex de letra: "Luis Dias"`
3. Aguarde até que o processo seja concluído. Os dados serão salvos na planilha 'nome_definido_por_vc.xlsx'.
4. Verifique os resultados na planilha gerada.

Passos e Explicações:
Imports de Bibliotecas:

selenium.webdriver: Para interagir com o navegador via Selenium.
openpyxl: Para manipular arquivos Excel.
time e datetime: Para manipulação de tempo e datas.

Inicialização e Navegação no Site:

Inicializa o WebDriver do Chrome e navega até a página inicial do site da Secretaria da Fazenda de SP.
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.styles import PatternFill
import openpyxl
import time
import datetime

# Inicializar o WebDriver
driver = webdriver.Chrome()
driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')
Carregamento da Planilha Excel:
Carrega um arquivo Excel preexistente para armazenar os dados extraídos do site.

# Carregar a planilha Excel
workbook = openpyxl.load_workbook('aiims.xlsx')  # Substitua 'aiims.xlsx' pelo nome do seu arquivo Excel
sheet = workbook.active
Definição de Cores para Formatação na Planilha:
Define diferentes cores para destacar informações na planilha Excel.

# Definir cores para preenchimento na planilha
cor_clickup = PatternFill(patternType='solid', fgColor='F0D402')
cor_outros = PatternFill(patternType='solid', fgColor='FF5B5B')
cor_naotem = PatternFill(patternType='solid', fgColor='55A3F9')
Configuração de Variáveis e Listas:
Configura variáveis como a data atual (DATE) e uma lista de DRTs específicos (outros) para comparação posterior.

# Data atual
DATE = datetime.date.today().strftime("%d/%m/%Y")
# Lista de DRTs específicos
outros = {
    "LITORAL", "OSASCO", 
    "CAPITAL I", "CAPITAL II", "CAPITAL III", 
    "GUARULHOS", 
    "DTE-II – FISCALIZAÇÃO ESPECIAL", "DTE-I – FISCALIZAÇÃO ESPECIAL",
    "Compliance MNM", "Compliance M&E"
}
Iteração sobre os Dados na Planilha:
Itera sobre as linhas da planilha, inserindo cada AIIM na página de pesquisa do site e coletando informações correspondentes.

# Iterar sobre as linhas da planilha
try:
for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
aiim = str(row[0])
Manipulação do WebDriver e Extração de Dados:
Utiliza Selenium para preencher campos de pesquisa, clicar em botões e extrair informações da página conforme necessário.

# Preencher AIIM na página de pesquisa
aiim_input = wait.until(EC.element_to_be_clickable((By.NAME, 'ctl00$ConteudoPagina$TxtNumAIIM')))
aiim_input.clear()
aiim_input.send_keys(aiim)

# Clicar no botão de pesquisa
pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Clique para pesquisar por numero do aiim (sem o digito verificador)']")))
pesquisar.click()

# Manipulação de exceções e coleta de dados
try:
# Manipulação de exceções e coleta de dados aqui...
except Exception as e:
# Manipulação de exceções e tratamento de erros aqui...
Atualização da Planilha com os Dados Coletados:
Atualiza a planilha Excel com os dados extraídos, formatando células de acordo com as condições especificadas.

finally:
# Atualização da planilha com os dados coletados
# Formatação das células de acordo com os resultados obtidos
Finalização e Salvamento da Planilha:
Salva a planilha Excel com os dados atualizados e fecha o navegador.

finally:
# Salvar o arquivo Excel
workbook.save('AiimsColetados.xlsx')

# Fechar o navegador após o uso
driver.quit()

# Mensagem de conclusão
print("Processo concluído com êxito!")

Essa estrutura de documentação fornece uma visão clara de cada parte do código e como elas contribuem para o objetivo final do script. Cada seção é descrita de forma sucinta e focada no que está sendo realizado, facilitando a compreensão e a manutenção futura do código.



## API utilizada para eles

[DataStone API](https://backoffice.datastone.com.br/docs/)
