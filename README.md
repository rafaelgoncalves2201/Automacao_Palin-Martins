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

### Detalhes do Script

- O codigo ascessa o site escolhido e faz buscas na página de consulta/pode ser utilizado para outras automações no mesmo estilo.
- Ele extrai informações como local, nome e data para cada coisa que ele consulta.
- Os dados são registrados na planilha, incluindo a data da execução do codigo
- Aqueles que resultem em erros durante a pesquisa são marcados como "Erro" na planilha porem fica a escolha.


## API utilizada para eles

[DataStone API](https://backoffice.datastone.com.br/docs/)
