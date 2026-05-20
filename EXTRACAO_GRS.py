import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
import webbrowser

# --- Lógica de Extração (Intacta) ---

def limpar_observacao(obs_bruta):
    """
    Limpa o campo de observação, removendo textos residuais de tabelas 
    que o leitor de PDF acaba misturando no meio do texto.
    """
    if not obs_bruta:
        return ""
    
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
    
    obs_limpa = re.sub(r"\s+", " ", obs_limpa).strip()
    obs_limpa = re.sub(r"^[,\.\-]\s*", "", obs_limpa).strip()
    
    return obs_limpa

def extrair_dados_uema(pdf_path, output_path, barra_progresso, janela):
    """Extrai os campos específicos do PDF e salva em Excel."""
    lista_dados = []

    with pdfplumber.open(pdf_path) as pdf:
        total_paginas = len(pdf.pages)
        
        for i, pagina in enumerate(pdf.pages):
            # CustomTkinter usa valores de 0.0 a 1.0 para progresso
            progresso_atual = (i + 1) / total_paginas
            barra_progresso.set(progresso_atual)
            janela.update()
            
            texto = pagina.extract_text()
            if not texto:
                continue

            match_numero = re.search(r"Número.*?(\d{4}GR\d+)", texto, re.IGNORECASE | re.DOTALL)
            match_data = re.search(r"Data Referência.*?(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE | re.DOTALL)
            match_domicilio = re.search(r"Domic[ií]lio Origem.*?([\d\s-]{10,})", texto, re.IGNORECASE | re.DOTALL)
            match_valor = re.search(r"Valor.*?([\d\.,]{4,})", texto, re.IGNORECASE | re.DOTALL)
            match_usuario = re.search(r"Lançado em.*?por\s+([A-ZÀ-Ÿa-z\s]+)", texto, re.IGNORECASE)
            match_obs = re.search(r"Observação(.*?)(?=Lançamentos|Transação Origem|N°\s+Evento|\Z)", texto, re.IGNORECASE | re.DOTALL)

            dados_pagina = {
                "Número": match_numero.group(1).strip() if match_numero else None,
                "Data Referência": match_data.group(1).strip() if match_data else None,
                "Valor": match_valor.group(1).strip() if match_valor else None,
                "Domicílio Origem": re.sub(r"\s+", " ", match_domicilio.group(1)).strip() if match_domicilio else None,
                "Observação": limpar_observacao(match_obs.group(1)) if match_obs else None,
                "Usuário": match_usuario.group(1).strip() if match_usuario else None
            }
            
            if dados_pagina["Número"]:
                lista_dados.append(dados_pagina)

    barra_progresso.set(1.0)
    janela.update()

    if not lista_dados:
        raise ValueError("Nenhum dado válido encontrado no PDF com o formato esperado.")

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
        entry_pdf.delete(0, ctk.END)
        entry_pdf.insert(0, caminho_pdf)

def selecionar_saida():
    caminho_saida = filedialog.asksaveasfilename(
        title="Salvar planilha como",
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")]
    )
    if caminho_saida:
        entry_saida.delete(0, ctk.END)
        entry_saida.insert(0, caminho_saida)

def mostrar_ajuda():
    janela_ajuda = ctk.CTkToplevel(janela)
    janela_ajuda.title("Informações de Desenvolvimento")
    janela_ajuda.geometry("380x250")
    janela_ajuda.resizable(False, False)
    janela_ajuda.attributes("-topmost", True) # Mantém a janela sempre na frente
    
    texto_info = (
        "Desenvolvido para automação e extração de dados.\n\n"
        "Autor: Philipe Sampaio\n"
        "Git: oldboy465\n"
        "Wpp: (98) 98250-6920\n\n"
        "Ferramenta de editar pdfs:"
    )
    
    # Label com o texto normal
    ctk.CTkLabel(janela_ajuda, text=texto_info, justify="center").pack(pady=(20, 5), padx=20)
    
    # Label com o link clicável
    link_url = "https://oldboy465.github.io/pdf92/"
    lbl_link = ctk.CTkLabel(janela_ajuda, text=link_url, text_color="#1f538d", 
                            cursor="hand2", font=ctk.CTkFont(underline=True))
    lbl_link.pack(pady=(0, 20))
    
    # Evento de clique para abrir o navegador
    lbl_link.bind("<Button-1>", lambda e: webbrowser.open_new(link_url))

def executar_processamento():
    pdf_path = entry_pdf.get()
    output_path = entry_saida.get()

    if not pdf_path:
        messagebox.showwarning("Aviso", "Por favor, selecione o arquivo PDF de origem.")
        return
    if not output_path:
        messagebox.showwarning("Aviso", "Por favor, escolha onde salvar o arquivo gerado.")
        return

    btn_executar.configure(text="Processando...", state="disabled")
    barra_progresso.set(0)
    janela.update()

    try:
        qtd_registros = extrair_dados_uema(pdf_path, output_path, barra_progresso, janela)
        messagebox.showinfo("Sucesso!", f"Processamento concluído com sucesso!\n\n{qtd_registros} guias extraídas para o Excel.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{str(e)}")
    finally:
        btn_executar.configure(text="Extrair Dados", state="normal")
        barra_progresso.set(0)

# --- Configuração da Janela Principal (CustomTkinter) ---

ctk.set_appearance_mode("System")  # Segue o tema do Windows (Dark/Light)
ctk.set_default_color_theme("blue")

janela = ctk.CTk()
janela.title("Extrator de GRS - UEMA (Avançado)")
janela.geometry("650x380")
janela.resizable(False, False)

# Título Superior
titulo = ctk.CTkLabel(janela, text="Extrator de GRS para Excel", font=ctk.CTkFont(size=20, weight="bold"))
titulo.pack(pady=(20, 5))

# Frame Principal
frame = ctk.CTkFrame(janela)
frame.pack(pady=10, padx=20, fill="both", expand=True)

# Grid Layout dentro do frame
frame.columnconfigure(0, weight=1)

# Linha 1: Arquivo PDF
ctk.CTkLabel(frame, text="Arquivo PDF de Origem:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(15, 0))
entry_pdf = ctk.CTkEntry(frame, width=380, placeholder_text="Selecione o PDF base...")
entry_pdf.grid(row=1, column=0, sticky="w", padx=15, pady=5)
btn_pdf = ctk.CTkButton(frame, text="Procurar PDF...", width=120, command=selecionar_pdf)
btn_pdf.grid(row=1, column=1, padx=(0, 15), pady=5)

# Linha 2: Arquivo de Saída
ctk.CTkLabel(frame, text="Salvar Planilha (Excel) em:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, sticky="w", padx=15, pady=(15, 0))
entry_saida = ctk.CTkEntry(frame, width=380, placeholder_text="Onde salvar o Excel...")
entry_saida.grid(row=3, column=0, sticky="w", padx=15, pady=5)
btn_saida = ctk.CTkButton(frame, text="Salvar como...", width=120, command=selecionar_saida)
btn_saida.grid(row=3, column=1, padx=(0, 15), pady=5)

# Linha 3: Barra de Progresso
barra_progresso = ctk.CTkProgressBar(frame, width=500)
barra_progresso.grid(row=4, column=0, columnspan=2, pady=(25, 10))
barra_progresso.set(0) # Inicia vazia

# Linha 4: Botões de Ação
frame_botoes = ctk.CTkFrame(frame, fg_color="transparent")
frame_botoes.grid(row=5, column=0, columnspan=2, pady=10)

btn_executar = ctk.CTkButton(frame_botoes, text="Extrair Dados", width=200, height=40, font=ctk.CTkFont(weight="bold"), command=executar_processamento)
btn_executar.pack(side="left", padx=10)

btn_ajuda = ctk.CTkButton(frame_botoes, text="Ajuda", width=100, height=40, fg_color="#555555", hover_color="#333333", command=mostrar_ajuda)
btn_ajuda.pack(side="left", padx=10)

janela.mainloop()