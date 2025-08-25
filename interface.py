import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
from pathlib import Path
from utils_path import app_data_dir, resource_dir
import sys, subprocess, json

DATA_DIR = app_data_dir()

TEMP_PLANILHA = DATA_DIR / "planilha_filtrada.xlsx"
CONFIG_PATH = DATA_DIR / "config.json"

ATENDENTES = [
    "Leticia Pereira Dos Anjos",
    "Daniel da Silva Monteiro",
    "Ana Paula de Sousa Macedo",
    "Bruno Oliveira da Silva"
]

df = None

def carregar_planilha():
    global df

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, header=None)
        num_linhas, num_cols = df.shape

        header_selector["values"] = [f"Linha {i+1}" for i in range(num_linhas)]
        header_selector.current(0)

        messagebox.showinfo("Sucesso", f"Planilha carregada com {num_cols} colunas e {num_linhas} linhas!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar planilha: {e}")

def aplicar_cabecalho(event=None):
    global df
    if df is None:
        return

    idx = header_selector.current()
    df.columns = df.iloc[idx].tolist()  # define header
    df = df.drop(range(idx+1))          # remove linhas do cabeçalho

    # Atualiza listbox de colunas
    listbox_colunas.delete(0, tk.END)
    for col in df.columns.tolist():
        listbox_colunas.insert(tk.END, col)

    # >>> CPF selector (novo)
    cpf_selector["values"] = df.columns.tolist()
    if df.columns.tolist():
        # tenta escolher automaticamente uma coluna com "cpf" no nome
        cols_lower = [c.lower() for c in df.columns.tolist()]
        try:
            idx = cols_lower.index("cpf")
            cpf_selector.current(idx)
        except ValueError:
            cpf_selector.current(0)

def iniciar_processo():
    global df
    if df is None:
        messagebox.showwarning("Atenção", "Carregue a planilha primeiro!")
        return

    selecionadas = listbox_colunas.curselection()
    if not selecionadas:
        messagebox.showwarning("Atenção", "Selecione pelo menos uma coluna!")
        return

    colunas_selecionadas = [listbox_colunas.get(i) for i in selecionadas]
    dados_filtrados = df[colunas_selecionadas].copy()

    col_cpf = cpf_selector.get()
    if not col_cpf:
        messagebox.showwarning("Atenção", "Selecione a coluna de CPF!")
        return

    if col_cpf not in colunas_selecionadas:
        colunas_selecionadas.append(col_cpf)

    atendente = atendente_selector.get()

    dados_filtrados.to_excel(TEMP_PLANILHA, index=False)

    config = {
        "colunas": colunas_selecionadas,
        "coluna_cpf": col_cpf,
        "atendente": atendente
    }
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    messagebox.showinfo("Resumo", f"""
    Colunas escolhidas: {colunas_selecionadas}
    Coluna CPF: {col_cpf}
    Atendente: {atendente}
    """)

    def run_script():
        if getattr(sys, 'frozen', False):
            editar_exec = resource_dir() / "editar.exe"
            subprocess.run([str(editar_exec)])
        else:
            subprocess.run([sys.executable, "editar.py"])

    threading.Thread(target=run_script).start()


# ------------------------- INTERFACE -------------------------

root = tk.Tk()
root.title("Automação do Follow-up")
root.geometry("700x500")

frame_top = tk.Frame(root)
frame_top.pack(pady=10)

btn_carregar = tk.Button(frame_top, text="Selecionar Planilha", command=carregar_planilha)
btn_carregar.pack()

# Cabeçalho
tk.Label(root, text="Selecione a linha de cabeçalho:").pack()
header_selector = ttk.Combobox(root, state="readonly")
header_selector.pack(pady=5)
header_selector.bind("<<ComboboxSelected>>", aplicar_cabecalho)

# Colunas múltiplas
tk.Label(root, text="Selecione colunas:").pack()
listbox_colunas = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=10)
listbox_colunas.pack(pady=5)

# Coluna CPF  (NOVO)
tk.Label(root, text="Selecione a coluna de CPF:").pack()
cpf_selector = ttk.Combobox(root, state="readonly")
cpf_selector.pack(pady=5)

# Atendente
tk.Label(root, text="Selecione o atendente:").pack()
atendente_selector = ttk.Combobox(root, values=ATENDENTES, state="readonly")
atendente_selector.set(ATENDENTES[0])
atendente_selector.pack(pady=5)

btn_iniciar = tk.Button(root, text="Iniciar Script", command=iniciar_processo, bg="green", fg="white")
btn_iniciar.pack(pady=20)

root.mainloop()
