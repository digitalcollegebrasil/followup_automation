import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import threading

def carregar_planilha():
    global df, colunas_selecionadas

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)
        colunas = df.columns.tolist()

        listbox_colunas.delete(0, tk.END)
        for col in colunas:
            listbox_colunas.insert(tk.END, col)
        messagebox.showinfo("Sucesso", f"{len(colunas)} colunas carregadas!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar planilha: {e}")

def iniciar_processo():
    global df

    selecionadas = listbox_colunas.curselection()
    if not selecionadas:
        messagebox.showwarning("Atenção", "Selecione pelo menos uma coluna!")
        return

    colunas_selecionadas = [listbox_colunas.get(i) for i in selecionadas]
    dados_filtrados = df[colunas_selecionadas].copy()

    temp_path = "planilha_filtrada.xlsx"
    dados_filtrados.to_excel(temp_path, index=False)

    def run_script():
        import subprocess
        subprocess.run(["python", "editar.py"])

    threading.Thread(target=run_script).start()

root = tk.Tk()
root.title("Automação do Follow-up")
root.geometry("600x400")

frame_top = tk.Frame(root)
frame_top.pack(pady=10)

btn_carregar = tk.Button(frame_top, text="Selecionar Planilha", command=carregar_planilha)
btn_carregar.pack()

listbox_colunas = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=15)
listbox_colunas.pack(pady=10)

btn_iniciar = tk.Button(root, text="Iniciar Script", command=iniciar_processo, bg="green", fg="white")
btn_iniciar.pack(pady=10)

root.mainloop()
