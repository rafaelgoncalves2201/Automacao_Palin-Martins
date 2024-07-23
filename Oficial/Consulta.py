import os
import sys
import threading
from tkinter import ttk
from datetime import datetime
from tkinter import filedialog
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

def normalizar_descricao(desc):
    padrao_decurso_prazo = r'\bDecurso de Prazo\b.*'
    padrao_distribuicao_julgamento = r'\bDistribuição da Defesa para Julgamento\b.*'
    padrao_peticao = r'\bProtocolo de Petição\b.*'
    
    if re.search(padrao_decurso_prazo, desc):
        return re.sub(padrao_decurso_prazo, 'Decurso de Prazo', desc).strip()
    
    if re.search(padrao_distribuicao_julgamento, desc):
        return re.sub(padrao_distribuicao_julgamento, 'Distribuição da Defesa para Julgamento', desc).strip()
    
    if re.search(padrao_peticao, desc):
        return re.sub(padrao_peticao, 'Protocolo de Petição', desc).strip()
    
    return desc.strip()

def formatar_aiim(aiim):
    apenas_numeros = re.sub(r'\D', '', aiim)
    if len(apenas_numeros) >= 8:
        apenas_numeros = apenas_numeros[:7] + apenas_numeros[8:]
    return apenas_numeros

def exibir(tipo, texto):
    cor = "black" if tipo in ["Sucesso", "Processo"] else "black"
    mensagem_label.config(text=texto, fg=cor)

def escolher_planilha():
    global file_path
    file_path = filedialog.askopenfilename(title="Selecione a planilha Excel", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        exibir("Sucesso", "Planilha selecionada com sucesso.")

def mostrar_progress_bar():
    progress_bar.pack(pady=20)
    janela.update_idletasks()

def esconder_progress_bar():
    progress_bar.pack_forget()

def atualizar_progress_bar(valor):
    progress_bar['value'] = valor
    janela.update_idletasks()

def executar_processo():
    if not file_path:
        exibir("Erro", "Por favor, selecione uma planilha primeiro.")
        return

    exibir("Processo", "Processo iniciado, por favor aguarde...")
    mostrar_progress_bar()
    
    global cancelar
    cancelar = False

    thread = threading.Thread(target=processo_automacao)
    thread.start()

def processo_automacao():
    global cancelar
    driver = None

    if not file_path:
        return

    try:
        servico = Service(ChromeDriverManager().install())
        opcoes = webdriver.ChromeOptions()
        opcoes.add_argument('--headless=new')
        driver = webdriver.Chrome(service=servico, options=opcoes)
        driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')

        workbook = load_workbook(file_path)
        sheet = workbook.active
        linha_planilha = 2
        wait = WebDriverWait(driver, 0.1)

        DATE = datetime.today()
        cor_clickup = PatternFill(patternType='solid', fgColor='F0D402')
        cor_outros = PatternFill(patternType='solid', fgColor='FF5B5B')
        cor_naotem = PatternFill(patternType='solid', fgColor='55A3F9')

        aiims_valido = {"Notificação do AIIM", "Inscrição na Dívida Ativa/ AIIM inscrito em dívida ativa",
                        "AIIM enviado para a Unidade Fiscal da Cobrança.", "Decurso de Prazo"}
        aiims_invalido = {"AIIM liquidado", "Protocolo da Defesa", "Protocolo de Petição",
                          "Entrada do processo na Delegacia Tributária de Julgamento.",
                          "Publicação no Diário Eletrônico", "Distribuição da Defesa para Julgamento", "Protocolo de Petição"}
        outros = {"LITORAL", "OSASCO", "CAPITAL I", "CAPITAL II", "CAPITAL III", "GUARULHOS",
                  "DTE-II – FISCALIZAÇÃO ESPECIAL", "DTE-I – FISCALIZAÇÃO ESPECIAL",
                  "Compliance MNM", "Compliance M&E"}

        nomeColunas = ["N°", "DRT", "D.AIIM", "CONTRIBUINTE", "ANDAMENTO DO AIIM",
                       "CNPJ", "TELEFONE", "E-MAIL", "CNAE", "D.DIA", "SITUAÇÂO"]

        for col, nomeColuna in enumerate(nomeColunas, start=1):
            sheet.cell(row=1, column=col).value = nomeColuna

        total_rows = len(list(sheet.iter_rows(min_row=2, max_col=1, values_only=True)))
        for index, row in enumerate(sheet.iter_rows(min_row=2, max_col=1, values_only=True), start=1):
            if cancelar:
                break  # Interrompe o loop se o cancelamento for solicitado

            aiim = row[0]
            if aiim is None or aiim == "":
                continue

            sheet.cell(row=linha_planilha, column=2).value = "DRT"
            aiim_form = formatar_aiim(str(aiim))
            aiim_input = wait.until(EC.element_to_be_clickable((By.NAME, 'ctl00$ConteudoPagina$TxtNumAIIM')))
            aiim_input.clear()
            aiim_input.send_keys(aiim_form)

            pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Clique para pesquisar por numero do aiim (sem o digito verificador)']")))
            pesquisar.click()

            try:
                alert_wait = WebDriverWait(driver, timeout=0.5)
                alert = alert_wait.until(EC.alert_is_present())
                driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')
                continue

            except Exception:
                ERRO = ""
                DRT = ""
                NOME = ""
                DATA = ""
                DESC = ""
                try:
                    elemento_drt = wait.until(EC.visibility_of_element_located((By.ID, 'ConteudoPagina_lblDRT')))
                    DRT = elemento_drt.text

                    elemento_nome = driver.find_element(By.ID, 'ConteudoPagina_lblNomeAutuado')
                    NOME = elemento_nome.text

                    elemento_data = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="dataEvento"]')))
                    if elemento_data:
                        DATA = datetime.strptime(elemento_data[-1].text, "%d/%m/%Y")

                    elemento_desc = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="descricaoEvento"]')))
                    if elemento_desc:
                        DESC = elemento_desc[-1].text

                except Exception as e:
                    ERRO = str(e)

                finally:
                    sheet.cell(row=linha_planilha, column=10).value = DATE.strftime("%d/%m/%Y")

                    desc = normalizar_descricao(DESC)
                    if ERRO == "":
                        sheet.cell(row=linha_planilha, column=2).value = DRT
                        sheet.cell(row=linha_planilha, column=3).value = DATA.strftime("%d/%m/%Y")
                        sheet.cell(row=linha_planilha, column=4).value = NOME
                        sheet.cell(row=linha_planilha, column=5).value = DESC
                        
                        if desc in aiims_valido:
                            sheet.cell(row=linha_planilha, column=11).value = "Passar ClickUp"
                            for col in range(1, 12):
                                sheet.cell(row=linha_planilha, column=col).fill = cor_clickup
                        else:
                            sheet.cell(row=linha_planilha, column=11).value = "Suspenso"
                            for col in range(1, 12):
                                sheet.cell(row=linha_planilha, column=col).fill = cor_outros
                        if desc in aiims_invalido:
                            sheet.cell(row=linha_planilha, column=11).value = "Suspenso"
                            for col in range(1, 12):
                                sheet.cell(row=linha_planilha, column=col).fill = cor_outros
                        if DRT in outros:
                            sheet.cell(row=linha_planilha, column=11).value = "Outros"
                            for col in range(1, 12):
                                sheet.cell(row=linha_planilha, column=col).fill = cor_outros
                    else:
                        sheet.cell(row=linha_planilha, column=2).value = "Erro"
                        sheet.cell(row=linha_planilha, column=11).value = "Não tem ainda"
                        for col in range(1, 12):
                            sheet.cell(row=linha_planilha, column=col).fill = cor_naotem

                    linha_planilha += 1
                    driver.get('https://www.fazenda.sp.gov.br/epat/extratoprocesso/PesquisarExtrato.aspx')

            atualizar_progress_bar((index / total_rows) * 100)

        if not cancelar:
            exibir("Pronto", "Processo finalizado salve a planilha")
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Salvar planilha como")
            if save_path:
                workbook.save(save_path)
            janela.event_generate("<<ProcessoConcluido>>")

    except Exception as e:
        janela.event_generate("<<ProcessoErro>>", data=str(e))
    finally:
        if driver:
            driver.quit()
            esconder_progress_bar()
            if not cancelar:
                exibir("Pronto", "Tudo Salvo")
                janela.event_generate("<<ProcessoSalvo>>")
        else:
            esconder_progress_bar()
            if cancelar:
                exibir("Aviso", "Processo cancelado. Nada foi salvo.")
            else:
                exibir("Pronto", "Tudo Salvo")
                janela.event_generate("<<ProcessoSalvo>>")

def atualizar_status(event):
    exibir("Pronto", "Processo finalizado salve a planilha")
    
def atualizar_salvamento(event):
    exibir("Pronto", "Tudo Salvo")

def cancelar():
    global cancelar
    cancelar = True
    exibir("Aviso", "Processo cancelado. Por favor, aguarde...")
    if not progress_bar['value'] == 100:
        progress_bar['value'] = 100
        janela.update_idletasks()

def mostrar_erro(event):
    exibir("Erro", f"Erro inesperado: {event.data}")
    
def nova_consulta():
    global file_path
    file_path = None
    exibir("Calma", "Nova consulta iniciada. Adicione uma nova planilha.")
    mensagem_label.config(text="", fg="black")
    progress_bar['value'] = 0
    janela.update_idletasks()
    
# Criação da janela principal
janela = tk.Tk()
janela.title("Automação para Consulta de Aiims")
janela.geometry("800x600")
janela.configure(bg="#f8f9fa")

# Carregar ícone
def carregar_icone():
    if getattr(sys, 'frozen', False):
        icone_path = os.path.join(sys._MEIPASS, "imgs", "palinico.ico")
    else:
        icone_path = os.path.join("imgs", "palinico.ico")
    return icone_path

icone_path = carregar_icone()
janela.iconbitmap(icone_path)

header_frame = tk.Frame(janela, bg="#007bff", padx=20, pady=10)
header_frame.pack(fill="x")

header_label = tk.Label(header_frame, text="Automação para Consulta de Aiims", bg="#007bff", fg="white", font=("Helvetica", 16, "bold"))
header_label.pack()

body_frame = tk.Frame(janela, bg="#f8f9fa")
body_frame.pack(padx=20, pady=20, fill="both", expand=True)

instruction_label = tk.Label(body_frame, text="Por favor, adicione a planilha Excel contendo os Aiims:", bg="#f8f9fa", fg="#495057", font=("Helvetica", 12))
instruction_label.pack(pady=10)

button_frame = tk.Frame(body_frame, bg="#f8f9fa")
button_frame.pack(pady=20)

planilha_btn = tk.Button(button_frame, text="Adicionar Planilha", command=escolher_planilha, bg="#007bff", fg="white", font=("Helvetica", 12), relief="flat", padx=10, pady=5)
planilha_btn.pack(pady=10)

iniciar_btn = tk.Button(button_frame, text="Iniciar", command=executar_processo, bg="#007bff", fg="white", font=("Helvetica", 12), relief="flat", padx=10, pady=5)
iniciar_btn.pack(pady=10)

cancelar_btn = tk.Button(body_frame, text="Cancelar processo", command=cancelar, bg="#f0371d", fg="white", font=("Helvetica", 12), relief="flat", padx=10, pady=5)
cancelar_btn.pack(pady=10)

nova_consulta_btn = tk.Button(janela, text="Iniciar nova consulta", bg="#007bff", fg="white", font=("Helvetica", 12), relief="flat", padx=10, pady=5)
nova_consulta_btn.pack(pady=10)

mensagem_label = tk.Label(body_frame, text="", bg="#f8f9fa", font=("Helvetica", 12))
mensagem_label.pack(pady=10)

# Barra de Progresso
progress_bar = ttk.Progressbar(body_frame, orient="horizontal", length=400, mode="determinate")

file_path = None
cancelar = False

janela.bind("<<ProcessoConcluido>>", atualizar_status)
janela.bind("<<ProcessoSalvo>>", atualizar_salvamento)
janela.bind("<<ProcessoErro>>", mostrar_erro)

janela.mainloop()