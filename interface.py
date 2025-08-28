import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
from utils_path import app_data_dir, resource_dir
import sys, subprocess, json
import tkinter.font as tkfont
from pathlib import Path
import re

# ===================== Util: fonte Inter (TTF privado / Windows) =====================
def _load_private_ttf(ttf_path: Path) -> bool:
    try:
        if not sys.platform.startswith("win"):
            return False
        import ctypes
        FR_PRIVATE = 0x10
        ttf_path = str(Path(ttf_path).resolve())
        added = ctypes.windll.gdi32.AddFontResourceExW(ttf_path, FR_PRIVATE, 0)
        return added > 0
    except Exception:
        return False

def _pick_inter_family():
    families = set(tkfont.families())
    for cand in ("Inter", "Inter Variable", "InterVariable", "Inter VF"):
        if cand in families:
            return cand
    for fb in ("Segoe UI", "Arial"):
        if fb in families:
            return fb
    return "TkDefaultFont"

# ===================== Paths / Consts =====================
DATA_DIR = app_data_dir()

TEMP_PLANILHA = DATA_DIR / "planilha_filtrada.xlsx"
CONFIG_PATH   = DATA_DIR / "config.json"

ATENDENTES = [
    "Leticia Pereira Dos Anjos",
    "Ana Celia da Silva Tavares",
]

SEDES = ["Aldeota", "Sul"]

# ===================== Estado global =====================
df = None            # DataFrame já com cabeçalho aplicado
df_raw = None        # DataFrame bruto da aba (sem cabeçalho)
XLSX_PATH = None
SHEET_NAMES = []

# ===================== Helpers =====================
def _linha_preview(row_values, max_cols=6, max_chars=80):
    vals = [str(v) for v in row_values[:max_cols]]
    s = " | ".join(vals)
    if len(s) > max_chars:
        s = s[:max_chars - 3] + "..."
    return s

def _popular_previa_cabecalhos(df_local):
    num_linhas, _ = df_local.shape
    opcoes = []
    for i in range(num_linhas):
        preview = _linha_preview(df_local.iloc[i].tolist())
        opcoes.append(f"Linha {i+1}: {preview}")
    header_selector["values"] = opcoes
    if opcoes:
        header_selector.current(0)

def _normalize_alunoid_series(s: pd.Series) -> pd.Series:
    # remove .0, espaços
    return (
        s.astype(str)
         .str.strip()
         .str.replace(r'\.0+$', '', regex=True)
    )

def _normalize_cpf_series(s: pd.Series) -> pd.Series:
    # mantém apenas dígitos
    return (
        s.astype(str)
         .str.replace(r'\D+', '', regex=True)
    )

# ===================== ações de UI =====================
def carregar_planilha():
    """Escolhe o arquivo .xlsx e popula o select de abas."""
    global XLSX_PATH, SHEET_NAMES, df, df_raw

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        XLSX_PATH = file_path
        xfile = pd.ExcelFile(XLSX_PATH)
        SHEET_NAMES = xfile.sheet_names

        if not SHEET_NAMES:
            messagebox.showerror("Erro", "Nenhuma aba encontrada no arquivo.")
            return

        sheet_selector["values"] = SHEET_NAMES
        sheet_selector.current(0)

        # Carrega a primeira aba por padrão
        carregar_aba()

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar planilha: {e}")

def carregar_aba(event=None):
    """Carrega o DataFrame da aba selecionada (sem cabeçalho) e gera prévias."""
    global df, df_raw
    if not XLSX_PATH:
        return

    aba = sheet_selector.get()
    if not aba:
        return

    try:
        df_raw = pd.read_excel(XLSX_PATH, sheet_name=aba, header=None, dtype=str)
        df = None  # ainda não aplicamos cabeçalho

        _popular_previa_cabecalhos(df_raw)

        # limpa seletores dependentes
        alunoid_selector["values"] = []
        alunoid_selector.set("")
        cpf_selector["values"] = []
        cpf_selector.set("")

        num_linhas, num_cols = df_raw.shape
        messagebox.showinfo("Sucesso", f"Aba '{aba}' carregada com {num_cols} colunas e {num_linhas} linhas!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar aba '{aba}': {e}")

def aplicar_cabecalho(event=None):
    """Usa a linha escolhida como cabeçalho e atualiza os seletores de AlunoID/CPF."""
    global df, df_raw
    if df_raw is None:
        return

    idx = header_selector.current()
    if idx is None or idx < 0:
        return

    header = df_raw.iloc[idx].tolist()
    df = df_raw.iloc[idx+1:].copy()
    df.columns = header
    df.reset_index(drop=True, inplace=True)

    cols = df.columns.tolist()

    # Preenche o seletor de AlunoID
    alunoid_selector["values"] = cols

    # Preenche o seletor de CPF
    cpf_selector["values"] = cols

    # Sugestão automática
    if cols:
        cols_lower = [c.lower() if isinstance(c, str) else "" for c in cols]
        # tenta AlunoID
        try:
            idx_ai = next(i for i, c in enumerate(cols_lower)
                          if "alunoid" in c or c == "id" or c == "aluno_id" or c == "calunoid")
            alunoid_selector.current(idx_ai)
        except StopIteration:
            pass

        # tenta CPF
        try:
            idx_cpf = next(i for i, c in enumerate(cols_lower) if c.strip() == "cpf")
            cpf_selector.current(idx_cpf)
        except StopIteration:
            pass

def iniciar_processo():
    """Gera uma planilha só com a coluna-chave (AlunoID ou CPF) e salva config."""
    global df
    if df is None:
        messagebox.showwarning("Atenção", "Carregue a planilha e defina o cabeçalho!")
        return

    col_alunoid = alunoid_selector.get().strip()
    col_cpf     = cpf_selector.get().strip()

    usar_alunoid = bool(col_alunoid and col_alunoid in df.columns)
    usar_cpf     = (not usar_alunoid) and bool(col_cpf and col_cpf in df.columns)

    if not usar_alunoid and not usar_cpf:
        messagebox.showwarning(
            "Atenção",
            "Escolha a coluna de AlunoID ou, se não houver, escolha a coluna de CPF."
        )
        return

    # monta df mínimo e dedup automático por chave
    if usar_alunoid:
        df_work = df[[col_alunoid]].copy()
        df_work[col_alunoid] = _normalize_alunoid_series(df_work[col_alunoid])
        df_work = df_work.dropna(subset=[col_alunoid])
        df_work = df_work[df_work[col_alunoid].astype(str).str.strip() != ""]
        df_work = df_work.drop_duplicates(subset=[col_alunoid], keep="first")
        colunas_config = [col_alunoid]  # compat com editar.py que ainda lê "colunas"
    else:
        df_work = df[[col_cpf]].copy()
        df_work[col_cpf] = _normalize_cpf_series(df_work[col_cpf])
        df_work = df_work.dropna(subset=[col_cpf])
        df_work = df_work[df_work[col_cpf].astype(str).str.strip() != ""]
        df_work = df_work.drop_duplicates(subset=[col_cpf], keep="first")
        colunas_config = [col_cpf]      # compat com editar.py

    atendente = atendente_selector.get()
    aba_escolhida = sheet_selector.get() or ""

    # credenciais e sede
    head_office    = sede_selector.get().strip()
    sponte_email   = email_entry.get().strip()
    sponte_password= password_entry.get().strip()

    # salva planilha mínima
    TEMP_PLANILHA.parent.mkdir(parents=True, exist_ok=True)
    df_work.to_excel(TEMP_PLANILHA, index=False)

    # monta config
    config = {
        "colunas": colunas_config,       # compat com seu editar.py
        "atendente": atendente,
        "aba": aba_escolhida,
        "head_office": head_office,
        "sponte_email": sponte_email,
        "sponte_password": sponte_password
    }
    if usar_alunoid:
        config["coluna_alunoid"] = col_alunoid
    else:
        # sem AlunoID → usamos CPF; o editar.py faz a busca do AlunoID via API
        config["coluna_cpf"] = col_cpf

    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    pw_mask = "•" * len(sponte_password) if sponte_password else "(vazio)"
    chave_txt = col_alunoid if usar_alunoid else col_cpf
    messagebox.showinfo(
        "Resumo",
        f"""
Aba: {aba_escolhida}
Chave usada: {'AlunoID' if usar_alunoid else 'CPF'} → {chave_txt}
Registros (após deduplicação): {len(df_work)}

Atendente: {atendente}
Sede: {head_office or '(não informada)'}
E-mail Sponte: {sponte_email or '(não informado)'}
Senha Sponte: {pw_mask}
"""
    )

    # Dispara o editar.py (ou editar.exe quando empacotado)
    def run_script():
        if getattr(sys, 'frozen', False):
            editar_exec = resource_dir() / "editar.exe"
            subprocess.run([str(editar_exec)])
        else:
            subprocess.run([sys.executable, "editar.py"])
    threading.Thread(target=run_script, daemon=True).start()

# ===================== UI =====================
root = tk.Tk()
root.title("Automação do Follow-up")
root.geometry("980x760")

# Fonte Inter
try:
    FONT_PATH = resource_dir() / "assets" / "Inter-VariableFont_opsz,wght.ttf"
    _load_private_ttf(FONT_PATH)
    inter_family = _pick_inter_family()
except Exception:
    inter_family = "TkDefaultFont"

# Tema & estilos
style = ttk.Style(root)
try:
    style.theme_use("clam")
except Exception:
    pass

BG = "#ffffff"  # fundo branco (troque se quiser outro tom)
root.configure(bg=BG)

# Fonte padrão + fundos para ttk
root.option_add("*Font", (inter_family, 11))
style.configure(".", font=(inter_family, 11))
style.configure("TFrame", background=BG)
style.configure("TLabel", background=BG)
style.configure("Headline.TLabel", background=BG, font=(inter_family, 16, "bold"))
style.configure("TLabelframe", background=BG)
style.configure("TLabelframe.Label", background=BG)
style.configure("TButton", padding=6)
style.configure("TEntry", fieldbackground="#ffffff", background="#ffffff")

# Header com logo e título
header = ttk.Frame(root)
header.pack(fill="x", pady=(8, 4), padx=12)

logo_img = None
try:
    logo_path = resource_dir() / "assets" / "logo.png"
    logo_img = tk.PhotoImage(file=str(logo_path))
    if logo_img.width() > 220:
        factor = max(1, logo_img.width() // 220)
        logo_img = logo_img.subsample(factor, factor)
except Exception:
    pass

if logo_img:
    lbl_logo = ttk.Label(header, image=logo_img, style="TLabel")
    lbl_logo.image = logo_img
    lbl_logo.pack(side="left")
else:
    ttk.Label(header, text="Digital College", style="Headline.TLabel").pack(side="left")

ttk.Label(header, text="Automação do Follow-up", style="Headline.TLabel").pack(side="left", padx=12)

# Container rolável (Canvas + Frame)
class Scrollable(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=BG)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # frame interno também com fundo branco
        self.body = ttk.Frame(self, style="TFrame")
        self.body_id = self.canvas.create_window((0, 0), window=self.body, anchor="nw")

        self.body.bind("<Configure>", self._on_body_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # mouse wheel
        self._bind_mousewheel(self.canvas)

    def _on_body_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.body_id, width=event.width)

    def _bind_mousewheel(self, widget):
        if sys.platform.startswith("win"):
            widget.bind_all("<MouseWheel>", self._on_mousewheel_win)
        elif sys.platform == "darwin":
            widget.bind_all("<MouseWheel>", self._on_mousewheel_mac)
        else:
            widget.bind_all("<Button-4>", self._on_mousewheel_linux)
            widget.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _on_mousewheel_win(self, event):
        self.canvas.yview_scroll(-int(event.delta / 120), "units")

    def _on_mousewheel_mac(self, event):
        self.canvas.yview_scroll(-int(event.delta), "units")

    def _on_mousewheel_linux(self, event):
        direction = -1 if event.num == 4 else 1
        self.canvas.yview_scroll(direction, "units")

scroll = Scrollable(root)
scroll.pack(fill="both", expand=True, padx=12, pady=(0, 12))

# Arquivo
frame_file = ttk.Frame(scroll.body)
frame_file.pack(pady=8, fill="x")
btn_carregar = ttk.Button(frame_file, text="Selecionar Planilha", command=carregar_planilha)
btn_carregar.pack(side="left")

# Sede / Login
frame_auth = ttk.LabelFrame(scroll.body, text="Acesso Sponte", style="TLabelframe")
frame_auth.pack(fill="x", padx=0, pady=6)

row1 = ttk.Frame(frame_auth); row1.pack(fill="x", pady=3)
ttk.Label(row1, text="Sede: ", width=18).pack(side="left")
sede_selector = ttk.Combobox(row1, values=SEDES, state="readonly", width=30)
sede_selector.set(SEDES[0])
sede_selector.pack(side="left", padx=4)

row2 = ttk.Frame(frame_auth); row2.pack(fill="x", pady=3)
ttk.Label(row2, text="E-mail Sponte: ", width=18).pack(side="left")
email_entry = ttk.Entry(row2, width=40)
email_entry.pack(side="left", padx=4)

row3 = ttk.Frame(frame_auth); row3.pack(fill="x", pady=3)
ttk.Label(row3, text="Senha Sponte: ", width=18).pack(side="left")
password_entry = ttk.Entry(row3, show="•", width=40)
password_entry.pack(side="left", padx=4)

# Seletor de aba
ttk.Label(scroll.body, text="Selecione a aba da planilha:").pack(anchor="w", pady=(10, 0))
sheet_selector = ttk.Combobox(scroll.body, state="readonly", width=60)
sheet_selector.pack(pady=5, anchor="w")
sheet_selector.bind("<<ComboboxSelected>>", carregar_aba)

# Cabeçalho
ttk.Label(scroll.body, text="Selecione a linha de cabeçalho (com prévia):").pack(anchor="w")
header_selector = ttk.Combobox(scroll.body, state="readonly", width=110)
header_selector.pack(pady=5, anchor="w")
header_selector.bind("<<ComboboxSelected>>", aplicar_cabecalho)

# Seletor de AlunoID
ttk.Label(scroll.body, text="Coluna de AlunoID (se existir):").pack(anchor="w")
alunoid_selector = ttk.Combobox(scroll.body, state="readonly", width=60)
alunoid_selector.pack(pady=5, anchor="w")

# Seletor de CPF
ttk.Label(scroll.body, text="Coluna de CPF (se não houver AlunoID):").pack(anchor="w")
cpf_selector = ttk.Combobox(scroll.body, state="readonly", width=60)
cpf_selector.pack(pady=5, anchor="w")

# Atendente
ttk.Label(scroll.body, text="Selecione o atendente:").pack(anchor="w")
atendente_selector = ttk.Combobox(scroll.body, values=ATENDENTES, state="readonly", width=60)
atendente_selector.set(ATENDENTES[0])
atendente_selector.pack(pady=5, anchor="w")

btn_iniciar = ttk.Button(scroll.body, text="Iniciar Script", command=iniciar_processo)
btn_iniciar.pack(pady=20)

root.mainloop()
