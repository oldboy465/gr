import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import re

def limpar_observacao(obs_bruta):
    """
    Limpa o campo de observação, removendo textos residuais de tabelas 
    que o leitor de PDF acaba misturando no meio do texto.
    """
    if not obs_bruta:
        return ""
    
    # Remove termos intrusos gerados pela formatação do SIGEF
    termos_remover = [
        r"Listar Guia Recebimento", 
        r"Detalhe", 
        r"Empenho Original", 
        r"Valor\s+[\d\.,]+", 
        r"Recolhedor", 
        r"Número Processo",
        r"Nota",
        r"\""
    ]
    
    obs_limpa = obs_bruta
    for termo in termos_remover:
        obs_limpa = re.sub(termo, "", obs_limpa, flags=re.IGNORECASE)
    
    # Remove quebras de linha e múltiplos espaços
    obs_limpa = re.sub(r"\s+", " ", obs_limpa).strip()
    
    # Remove vírgulas ou caracteres soltos no início da string
    obs_limpa = re.sub(r"^[,\.\-]\s*", "", obs_limpa).strip()
    
    return obs_limpa

def extrair_dados_uema(pdf_path, output_path, progress_var, janela):
    """Extrai os campos específicos do PDF e salva em Excel."""
    lista_dados = []

    with pdfplumber.open(pdf_path) as pdf:
        total_paginas = len(pdf.pages)
        
        for i, pagina in enumerate(pdf.pages):
            # Atualiza a barra de progresso
            progress_var.set((i / total_paginas) * 100)
            janela.update_idletasks()
            
            texto = pagina.extract_text()
            if not texto:
                continue

            # Extração refinada via Regex
            match_numero = re.search(r"Número.*?(\d{4}GR\d+)", texto, re.IGNORECASE | re.DOTALL)
            match_data = re.search(r"Data Referência.*?(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE | re.DOTALL)
            match_domicilio = re.search(r"Domic[ií]lio Origem.*?([\d\s-]{10,})", texto, re.IGNORECASE | re.DOTALL)
            
            # Pega o primeiro valor financeiro encontrado após a palavra Valor
            match_valor = re.search(r"Valor.*?([\d\.,]{4,})", texto, re.IGNORECASE | re.DOTALL)
            
            # Pega o nome do usuário que lançou
            match_usuario = re.search(r"Lançado em.*?por\s+([A-ZÀ-Ÿa-z\s]+)", texto, re.IGNORECASE)
            
            # Pega o bloco da observação (entre a palavra Observação e Lançamentos)
            match_obs = re.search(r"Observação(.*?)(?=Lançamentos|Transação Origem|N°\s+Evento|\Z)", texto, re.IGNORECASE | re.DOTALL)

            dados_pagina = {
                "Número": match_numero.group(1).strip() if match_numero else None,
                "Data Referência": match_data.group(1).strip() if match_data else None,
                "Valor": match_valor.group(1).strip() if match_valor else None,
                "Domicílio Origem": re.sub(r"\s+", " ", match_domicilio.group(1)).strip() if match_domicilio else None,
                "Observação": limpar_observacao(match_obs.group(1)) if match_obs else None,
                "Usuário": match_usuario.group(1).strip() if match_usuario else None
            }
            
            # Só adiciona se encontrou pelo menos o Número da GR
            if dados_pagina["Número"]:
                lista_dados.append(dados_pagina)

    # Preenche a barra de progresso ao finalizar
    progress_var.set(100)
    janela.update_idletasks()

    if not lista_dados:
        raise ValueError("Nenhum dado válido encontrado no PDF com o formato esperado.")

    # Criação do DataFrame e exportação
    df = pd.DataFrame(lista_dados)
    df.to_excel(output_path, index=False)
    return len(df)

# --- Funções da Interface Gráfica ---

def selecionar_pdf():
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
    )
    if caminho_pdf:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, caminho_pdf)

def selecionar_saida():
    caminho_saida = filedialog.asksaveasfilename(
        title="Salvar planilha como",
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")]
    )
    if caminho_saida:
        entry_saida.delete(0, tk.END)
        entry_saida.insert(0, caminho_saida)

def executar_processamento():
    pdf_path = entry_pdf.get()
    output_path = entry_saida.get()

    if not pdf_path:
        messagebox.showwarning("Aviso", "Por favor, selecione o arquivo PDF de origem.")
        return
    if not output_path:
        messagebox.showwarning("Aviso", "Por favor, escolha onde salvar o arquivo gerado.")
        return

    btn_executar.config(text="Processando...", state=tk.DISABLED)
    barra_progresso['value'] = 0
    janela.update()

    try:
        qtd_registros = extrair_dados_uema(pdf_path, output_path, progresso, janela)
        messagebox.showinfo("Sucesso!", f"Processamento concluído com sucesso!\n\n{qtd_registros} guias extraídas para o Excel.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{str(e)}")
    finally:
        btn_executar.config(text="Extrair Dados", state=tk.NORMAL)
        progresso.set(0)

# --- Configuração da Janela Principal ---

janela = tk.Tk()
janela.title("Extrator de GRS - UEMA (Avançado)")
janela.geometry("580x300")
janela.resizable(False, False)

style = ttk.Style()
style.theme_use('clam')

frame = ttk.Frame(janela, padding="20 20 20 20")
frame.pack(fill=tk.BOTH, expand=True)

# Linha 1: Arquivo PDF
ttk.Label(frame, text="Arquivo PDF de Origem:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
entry_pdf = ttk.Entry(frame, width=55)
entry_pdf.grid(row=1, column=0, sticky=tk.EW, padx=(0, 10))
ttk.Button(frame, text="Procurar PDF...", command=selecionar_pdf).grid(row=1, column=1)

# Linha 2: Arquivo de Saída
ttk.Label(frame, text="Salvar Planilha (Excel) em:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=(15, 5))
entry_saida = ttk.Entry(frame, width=55)
entry_saida.grid(row=3, column=0, sticky=tk.EW, padx=(0, 10))
ttk.Button(frame, text="Salvar como...", command=selecionar_saida).grid(row=3, column=1)

# Linha 3: Barra de Progresso
progresso = tk.DoubleVar()
barra_progresso = ttk.Progressbar(frame, variable=progresso, maximum=100)
barra_progresso.grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=(20, 0))

# Linha 4: Botão de Execução
btn_executar = ttk.Button(frame, text="Extrair Dados", command=executar_processamento)
btn_executar.grid(row=5, column=0, columnspan=2, pady=(15, 0), ipadx=30, ipady=8)

janela.mainloop()