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

        # Inicializar o WebDriver
        driver = webdriver.Chrome()
        driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')
        
        # Carregar a planilha Excel
        workbook = openpyxl.load_workbook('aiims.xlsx')  # Substitua 'aiims.xlsx' pelo nome do seu arquivo Excel
        sheet = workbook.active
        
        linha_planilha = 2  # Começar na segunda linha, supondo que a primeira linha seja cabeçalho
        
        # Definir tempo de espera máximo
        time.sleep(0.5)
        wait = WebDriverWait(driver, 0.5)
        
        # Define data de quando foi utilizado 
        DATE = datetime.date.today().strftime("%d/%m/%Y")
        
        cor_clickup = PatternFill(patternType='solid', fgColor='F0D402')
        cor_outros = PatternFill(patternType='solid', fgColor='FF5B5B')
        cor_naotem = PatternFill(patternType='solid', fgColor='55A3F9')
        
        
        
        outros = {
                "LITORAL", "OSASCO", 
                "CAPITAL I", "CAPITAL II", "CAPITAL III", 
                "GUARULHOS", 
                "DTE-II – FISCALIZAÇÃO ESPECIAL", "DTE-I – FISCALIZAÇÃO ESPECIAL",
                "Compliance MNM", "Compliance M&E"
                }
        
        
        nomeColunas = ["N°", "DRT", "D.AIIM", 
                    "CONTRIBUINTE", "CNPJ", 
                    "TELEFONE", "E-MAIL", 
                    "CNAE", "D.DIA", "SITUAÇÂO"
                    ]
        
        try:
            for col, nomeColunas in enumerate(nomeColunas, start=1):
                sheet.cell(row=1, column=col).value = nomeColunas
            # Loop pelas células com dados na planilha
            for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
                aiim = row[0]
                aiim = str(aiim)
        
        sheet.cell(row=linha_planilha, column=2).value = "DRT"

        # Espere até que o campo de AIIM seja clicável e insira o valor
        aiim_input = wait.until(EC.element_to_be_clickable((By.NAME, 'ctl00$ConteudoPagina$TxtNumAIIM')))
        aiim_input.clear()  # Limpa o campo antes de inserir o próximo AIIM
        aiim_input.send_keys(str(aiim))

        # Espere até que o botão de pesquisa seja clicável e clique nele
        pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Clique para pesquisar por numero do aiim (sem o digito verificador)']")))
        pesquisar.click()
        
        try:
            alert_wait = WebDriverWait(driver, timeout=0.5)
            alert = alert_wait.until(EC.alert_is_present())
            driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')
        except Exception as e:
            ERRO = "" # String Vazia
            DRT = ""
            nomeColunas = "DRT, D.AIIM, CONTRIBUINTE, CNPJ, TELEFONE, E-MAIL, CNAE, D.DIA, SITUAÇÂO"
            try:
                # Espera até que o elemento com o ID 'ConteudoPagina_lblDRT' seja visível e obtenha o texto
                elemento_drt = wait.until(EC.visibility_of_element_located((By.ID, 'ConteudoPagina_lblDRT')))
                DRT = elemento_drt.text

                # Pesquisar nome
                elemento_nome = driver.find_element(By.ID, 'ConteudoPagina_lblNomeAutuado')
                NOME = elemento_nome.text

                # Pesquisar data
                elemento_data = driver.find_element(By.CSS_SELECTOR, 'td.td1#dataEvento')
                DATA = elemento_data.text
                
            except Exception as e:
                # Se ocorrer uma exceção durante a busca de informações, registre o erro na planilha
                ERRO = str(e)
            finally:
                # Exibe a data do Dia
                sheet.cell(row=linha_planilha, column=9).value = DATE
                # Verificar se há erro e escrever na planilha
                if ERRO == "":
                    sheet.cell(row=linha_planilha, column=2).value = DRT
                    sheet.cell(row=linha_planilha, column=3).value = DATA
                    sheet.cell(row=linha_planilha, column=4).value = NOME
                    sheet.cell(row=linha_planilha, column=10).value = "Passar Click Up"
                    for col in range(1, 11):
                        sheet.cell(row=linha_planilha, column=col).fill = cor_clickup
                else:
                    sheet.cell(row=linha_planilha, column=2).value = "Erro"
                    sheet.cell(row=linha_planilha, column=10).value = "Não tem ainda"
                    for col in range(1, 11):
                        sheet.cell(row=linha_planilha, column=col).fill = cor_naotem
                    
                # Verificar se DRT está na lista outros
                if DRT in outros:
                    sheet.cell(row=linha_planilha, column=10).value = "Outros"
                    for col in range(1, 11):
                        sheet.cell(row=linha_planilha, column=col).fill = cor_outros
                    

                # Incrementar o contador de linha em qualquer caso
                linha_planilha += 1

                # Volte para a página inicial após cada iteração
                driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')
            

        finally:
            # Salvar o arquivo Excel
            workbook.save('AiimsColetados.xlsx')  # Salvar com um novo nome para evitar a substituição do original
        
            # Fechar o navegador após o uso
            driver.quit()
        
            # Enviar mensagem de êxito
            print("Processo concluído com êxito!")

Essa estrutura de documentação fornece uma visão clara de cada parte do código e como elas contribuem para o objetivo final do script. Cada seção é descrita de forma sucinta e focada no que está sendo realizado, facilitando a compreensão e a manutenção futura do código.



## API utilizada para eles

[DataStone API](https://backoffice.datastone.com.br/docs/)
