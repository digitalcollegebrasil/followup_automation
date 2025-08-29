import os
import time
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import json
from utils_path import app_data_dir
from pathlib import Path
from zeep import Client

load_dotenv()

# ---------- WSDL ----------
wsdl = 'https://api.sponteeducacional.net.br/WSAPIEdu.asmx?WSDL'
client = Client(wsdl=wsdl)

credenciais = {
    'Aldeota': {'codigo_cliente': '72546', 'token': 'QZUSqqgsLA63'},
    'Sul':     {'codigo_cliente': '74070', 'token': 'jVNLW7IIUXOh'}
    # 'Bezerra': {...}  # adicione se precisar usar API p/ essa sede
}

DATA_DIR = app_data_dir()
DATA_DIR.mkdir(parents=True, exist_ok=True)
planilha_path = DATA_DIR / "planilha_filtrada.xlsx"
config_path  = DATA_DIR / "config.json"

def click_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView();", element)
    driver.execute_script("arguments[0].click();", element)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")

# ---------- Carrega config ----------
try:
    if not config_path.exists():
        raise FileNotFoundError(f"Arquivo de configuração não encontrado: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    for k in ("colunas", "atendente", "coluna_alunoid"):
        if k not in config:
            raise KeyError(f"Chave '{k}' ausente no config.json")

    colunas_selecionadas = config["colunas"]
    atendente_escolhido = config["atendente"]
    coluna_alunoid = config["coluna_alunoid"]

    # NOVO: sede / email / senha vindos da interface, com fallback para .env
    head_office = config.get("head_office") or os.getenv("HEAD_OFFICE")
    email_address = config.get("sponte_email") or os.getenv("SPONTE_EMAIL")
    password_value = config.get("sponte_password") or os.getenv("SPONTE_PASSWORD")

    if not head_office:
        raise KeyError("Sede (head_office) não informada no config e não encontrada no .env")
    if not email_address or not password_value:
        raise KeyError("Credenciais do Sponte ausentes: informe e-mail e senha pela interface ou .env")

    print(f"Sede: {head_office}")
    print("Configuração recebida:")
    print(f"→ Colunas: {colunas_selecionadas}")
    print(f"→ Coluna AlunoID: {coluna_alunoid}")
    print(f"→ Atendente: {atendente_escolhido}")
    # NÃO imprimir senha

except Exception as e:
    raise SystemExit(f"[ERRO] Falha ao ler config.json/credenciais: {e}")

if not planilha_path.exists():
    raise SystemExit(f"[ERRO] Planilha não encontrada: {planilha_path}. Gere pela interface primeiro.")

# -------- Buscar CPF do Aluno --------
def get_aluno_cpf(cpf, codigo_cliente, token):
    parametros_busca = f'CPF={cpf}'
    try:
        response = client.service.GetAlunos(nCodigoCliente=codigo_cliente, sToken=token, sParametrosBusca=parametros_busca)
        
        if response:
            for aluno in response:
                return aluno['AlunoID']
        else:
            return None
    except Exception as e:
        print(f"Erro ao obter aluno para o CPF {cpf}: {e}")
        return None

# ---------- Carrega planilha ----------
try:
    dados = pd.read_excel(planilha_path, dtype=str)
    # se houver data_nascimento, formata; se não, ignora
    if "data_nascimento" in [c.lower() if isinstance(c, str) else c for c in dados.columns]:
        try:
            dados['data_nascimento'] = pd.to_datetime(dados['data_nascimento']).dt.strftime('%d/%m/%Y')
        except Exception:
            pass
    print("Cabeçalho da planilha:", dados.columns.tolist())
    print("Planilha carregada com sucesso!")
except Exception as e:
    raise SystemExit(f"Erro ao carregar a planilha: {e}")

# ---------- Login ----------
driver = None
try:
    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://www.sponteeducacional.net.br/Home.aspx")

    email = driver.find_element(By.ID, "txtLogin")
    email.send_keys(email_address)
    password = driver.find_element(By.ID, "txtSenha")
    password.send_keys(password_value)
    login_button = driver.find_element(By.ID, "btnok")
    login_button.click()
    time.sleep(8)

    print(head_office)
    enterprise = driver.find_element(By.ID, "ctl00_spnNomeEmpresa").get_attribute("innerText").strip().replace(" ", "")
    print(enterprise)

    combinacoes = {
        ("Aldeota", "DIGITALCOLLEGESUL-74070"): (1, "Acessando a sede Aldeota."),
        ("Aldeota", "DIGITALCOLLEGEBEZERRADEMENEZES-488365"): (1, "Acessando a sede Aldeota."),
        ("Sul", "DIGITALCOLLEGEALDEOTA-72546"): (3, "Acessando a sede Sul."),
        ("Sul", "DIGITALCOLLEGEBEZERRADEMENEZES-488365"): (3, "Acessando a sede Sul."),
        ("Bezerra", "DIGITALCOLLEGEALDEOTA-72546"): (4, "Acessando a sede Bezerra."),
        ("Bezerra", "DIGITALCOLLEGESUL-74070"): (4, "Acessando a sede Bezerra."),
        ("Aldeota", "DIGITALCOLLEGEALDEOTA-72546"): (None, "O script já está na Aldeota."),
        ("Sul", "DIGITALCOLLEGESUL-74070"): (None, "O script já está no Sul."),
        ("Bezerra", "DIGITALCOLLEGEBEZERRADEMENEZES-488365"): (None, "O script já está na Bezerra."),
    }

    val, message = combinacoes.get((head_office, enterprise), (None, "Ação não realizada: combinação não reconhecida."))
    print(message)

    if val is not None:
        driver.execute_script(f"$('#ctl00_hdnEmpresa').val({val});javascript:__doPostBack('ctl00$lnkChange','');")
        time.sleep(3)

    # ---------- Loop pelos registros ----------
    cols_lower = [str(c).lower() for c in dados.columns]
    if coluna_alunoid not in dados.columns:
        # tenta match case-insensitive
        try:
            real_name = next(c for c in dados.columns if str(c).lower() == str(coluna_alunoid).lower())
            coluna_alunoid = real_name
        except StopIteration:
            raise SystemExit(f"Coluna '{coluna_alunoid}' não encontrada na planilha.")

    for idx, row in dados.iterrows():
        aluno_id = row.get(coluna_alunoid)
        if pd.isna(aluno_id) or str(aluno_id).strip() == "":
            print(f"[AVISO] Linha {idx}: AlunoID vazio. Pulando...")
            continue

        aluno_id = str(aluno_id).strip()
        print(f"\n--- Processando linha {idx} ---")
        print(row.to_dict())   # mostra os dados no formato dict (coluna -> valor)

        # Abre direto a ficha do aluno usando o ID
        driver.get(f"https://www.sponteeducacional.net.br/SPCad/AlunoCadastro.aspx?cad=true&id={aluno_id}&ce=1")
        time.sleep(5)

        # Aba Follow-up
        followup_span = driver.find_element(By.XPATH, "//*[@id='__tab_tab_TabPanel9']")
        click_element(driver, followup_span)
        time.sleep(3)

        btn_incluir = driver.find_element(By.ID, "tab_TabPanel9_btnIncluirFollowUp_div")
        click_element(driver, btn_incluir)
        time.sleep(3)

        # Atendente (mantido o mapeamento por sede; pode usar 'atendente_escolhido' do config, se desejar)
        atendentes_default = {
            "Aldeota": "Leticia Pereira Dos Anjos",
            "Sul": "Leticia Pereira Dos Anjos",
        }
        atendente_nome = atendente_escolhido or atendentes_default.get(head_office, "Leticia Pereira Dos Anjos")

        # --- Entrar no iframe que acabou de abrir ---
        iframe = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'FollowUpCadastro.aspx')]"))
        )
        driver.switch_to.frame(iframe)

        sel_at = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "cmbAtendente"))
        )
        Select(sel_at).select_by_visible_text(atendente_nome)
        time.sleep(2)

        # Tipo de contato: WhatsApp
        sel_tc = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "cmbTipoContato"))
        )
        Select(sel_tc).select_by_visible_text("WhatsApp")
        time.sleep(2)

        # Tipo de Agendamento: Cobrança
        sel_ta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "cmbTipoAgendamento"))
        )
        Select(sel_ta).select_by_visible_text("Cobrança")
        time.sleep(2)

        # Grau de Interesse: Muito Interessado
        sel_gi = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "cmbGrauInteresse"))
        )
        Select(sel_gi).select_by_visible_text("Muito Interessado")
        time.sleep(2)

        # mapa de meses em PT-BR (salve o arquivo em UTF-8)
        MESES_PT = {
            1: "JANEIRO",
            2: "FEVEREIRO",
            3: "MARÇO",
            4: "ABRIL",
            5: "MAIO",
            6: "JUNHO",
            7: "JULHO",
            8: "AGOSTO",
            9: "SETEMBRO",
            10: "OUTUBRO",
            11: "NOVEMBRO",
            12: "DEZEMBRO",
        }

        mes_atual = MESES_PT[datetime.now().month]
        assunto_texto = f"COBRANÇA PARCELA - {mes_atual}"

        assunto_field = driver.find_element(By.ID, "txtAssunto")
        assunto_field.clear()
        assunto_field.send_keys(assunto_texto)
        time.sleep(2)

        # Se quiser reativar os cliques de salvar, descomente:
        # salvar_modal_button = driver.find_element(By.XPATH, "//div[@id='updRodapeFixo']//div[@id='btnSalvar_div']")
        # click_element(driver, salvar_modal_button)
        # time.sleep(3)

        # driver.switch_to.default_content()

        # salvar_principal_button = driver.find_element(By.XPATH, "//div[@id='updRodapeRelativo']//div[@id='btnSalvar_div']")
        # click_element(driver, salvar_principal_button)
        # time.sleep(3)

except Exception as e:
    print(f"Erro: {e}")

finally:
    try:
        if driver:
            driver.quit()
    except Exception:
        pass
