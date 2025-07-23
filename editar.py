import os
import time
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from zeep import Client
from openpyxl import Workbook

load_dotenv()

head_office = os.getenv("HEAD_OFFICE")
email_address = os.getenv("SPONTE_EMAIL")
password_value = os.getenv("SPONTE_PASSWORD")

print(f"Sede: {head_office}")

wsdl = 'https://api.sponteeducacional.net.br/WSAPIEdu.asmx?WSDL'
client = Client(wsdl=wsdl)

credenciais = {
    'Aldeota': {
        'codigo_cliente': '72546',
        'token': 'QZUSqqgsLA63'
    },
    'Sul': {
        'codigo_cliente': '74070',
        'token': 'jVNLW7IIUXOh'
    }
}

def remove_value_attribute(driver, element):
    driver.execute_script("arguments[0].removeAttribute('value')", element)

def set_input_value(driver, element, value):
    driver.execute_script("arguments[0].value = arguments[1]", element, value)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")

def click_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView();", element)
    driver.execute_script("arguments[0].click();", element)

script_dir = os.path.dirname(os.path.abspath(__file__))
planilha_path = os.path.join(script_dir, "planilha_filtrada.xlsx")

already_registered_path = os.path.join(script_dir, f"registros_cadastrados_{head_office}.xlsx")
error_path = os.path.join(script_dir, f"registros_com_erro_{head_office}.xlsx")

if os.path.exists(already_registered_path):
    try:
        registros_cadastrados = pd.read_excel(already_registered_path)
        print("Planilha de registros já cadastrados carregada com sucesso!")
    except Exception as e:
        print(f"Erro ao carregar a planilha de registros cadastrados: {e}")
else:
    registros_cadastrados = pd.DataFrame(columns=["nome_completo", "cpf", "email", "data_nascimento", "cep", "logradouro", "numero", "bairro", "telefone"])

if os.path.exists(error_path):
    try:
        registros_com_erro = pd.read_excel(error_path)
        print("Planilha de registros com erro carregada com sucesso!")
    except Exception as e:
        print(f"Erro ao carregar a planilha de registros com erro: {e}")
else:
    registros_com_erro = pd.DataFrame(columns=["nome_completo", "cpf", "email", "data_nascimento", "cep", "logradouro", "numero", "bairro", "telefone"])

def get_aluno_id(nome, codigo_cliente, token):
    parametros_busca = f'Nome={nome}'
    try:
        response = client.service.GetAlunos(nCodigoCliente=codigo_cliente, sToken=token, sParametrosBusca=parametros_busca)
        
        if response:
            for aluno in response:
                return aluno['AlunoID']
        else:
            return None
    except Exception as e:
        print(f"Erro ao obter aluno para o nome {nome}: {e}")
        return None

def salvar_registro_aluno(row):
    novo_registro = {
        "nome_completo": row['nome_completo'],
        "cpf": row['cpf'],
        "email": row['email'],
        "data_nascimento": row['data_nascimento'],
        "cep": row['cep'],
        "logradouro": row['logradouro'],
        "numero": row['numero'],
        "bairro": row['bairro'],
        "telefone": row['telefone'],
    }
    registros_cadastrados.loc[len(registros_cadastrados)] = novo_registro
    registros_cadastrados.to_excel(already_registered_path, index=False)

def salvar_registro_error(row):
    novo_registro = {
        "nome_completo": row['nome_completo'],
        "cpf": row['cpf'],
        "email": row['email'],
        "data_nascimento": row['data_nascimento'],
        "cep": row['cep'],
        "logradouro": row['logradouro'],
        "numero": row['numero'],
        "bairro": row['bairro'],
        "telefone": row['telefone'],
    }
    registros_com_erro.loc[len(registros_com_erro)] = novo_registro
    registros_com_erro.to_excel(error_path, index=False)

try:
    dados = pd.read_excel(planilha_path)
    dados['data_nascimento'] = pd.to_datetime(dados['data_nascimento']).dt.strftime('%d/%m/%Y')
    print(dados)
    print("Cabeçalho da planilha:", dados.columns.tolist())
    print("Planilha carregada com sucesso!")
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")

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

    resultado = combinacoes.get((head_office, enterprise), (None, "Ação não realizada: combinação não reconhecida."))
    val, message = resultado
    print(message)

    if val is not None:
        driver.execute_script(f"$('#ctl00_hdnEmpresa').val({val});javascript:__doPostBack('ctl00$lnkChange','');")
        time.sleep(3)

    for _, row in dados.iterrows():
        aluno_id = get_aluno_id(row['nome_completo'], credenciais[head_office]['codigo_cliente'], credenciais[head_office]['token'])

        driver.get(f"https://www.sponteeducacional.net.br/SPCad/AlunoCadastro.aspx?cad=true&id={aluno_id}&ce=1")
        time.sleep(5)
        
        # interessado_span = driver.find_element(By.XPATH, "//*[@id='__tab_tab_TabPanel6']")
        # click_element(driver, interessado_span)
        # time.sleep(5)

        # atendente_select = driver.find_element(By.XPATH, "//*[(@id='select2-tab_TabPanel6_tabInteressado_tbinteressado_cmbAtendente-container')]")
        # atendente_select.click()
        # time.sleep(5)

        # atendentes = {
        #     "Aldeota": "Mara Michele Dos Santos Milfort",
        #     "Sul": "Mayara de Araujo Barros",
        # }

        # atendente_nome = atendentes.get(head_office, "Mara Michele Dos Santos Milfort")
        # atendente_option = driver.find_element(By.XPATH, f"//li[contains(text(), '{atendente_nome}')]")
        # atendente_option.click()

        # turma_de_interesse_container = driver.find_element(By.XPATH, "//*[(@id='__tab_tab_TabPanel6_tabInteressado_tbCursos')]")
        # click_element(driver, turma_de_interesse_container)
        # time.sleep(5)

        # tipo_curso_select = driver.find_element(By.XPATH, "//*[(@id='select2-tab_TabPanel6_tabInteressado_tbCursos_cmbTipoCurso-container')]")
        # tipo_curso_select.click()
        # time.sleep(5)

        # tipos_curso = {
        #     "Aldeota": "Curso Livre Formação",
        #     "Sul": "Curso Livre Formação",
        # }
        
        # tipo_curso_nome = tipos_curso.get(head_office, tipos_curso[head_office])
        # tipo_curso_option = driver.find_element(By.XPATH, f"//li[contains(text(), '{tipo_curso_nome}')]")
        # tipo_curso_option.click()
        # time.sleep(5)

        # turma_interesse_select = driver.find_element(By.XPATH, "//*[(@id='select2-tab_TabPanel6_tabInteressado_tbCursos_cmbTurma-container')]")
        # turma_interesse_select.click()
        # time.sleep(5)

        # turma_interesse_option = driver.find_element(By.XPATH, f"//li[contains(text(), '{turma}')]")
        # click_element(driver, turma_interesse_option)
        # time.sleep(5)

        # curso_interesse = driver.find_element(By.XPATH, "//a[contains(text(), 'Livre Formação em Full Stack 20231')]")
        # checkbox = curso_interesse.find_element(By.XPATH, "./preceding-sibling::input[@type='checkbox']")
        # click_element(driver, checkbox)
        # time.sleep(5)

        # documentos_aba = driver.find_element(By.XPATH, "//*[@id='__tab_tab_TabPanel6_tabInteressado_tabDocumento']")
        # click_element(driver, documentos_aba)
        # time.sleep(5)

        # documentos_select = driver.find_element(By.XPATH, "//*[(@id='select2-tab_TabPanel6_tabInteressado_tabDocumento_cmbDocumento-container')]")
        # documentos_select.click()
        # time.sleep(5)

        # documentos_option = driver.find_element(By.XPATH, "//li[contains(text(), 'TERMO DE COMPROMISSO - GERAÇÃO TECH 2025')]")
        # documentos_option.click()
        # time.sleep(5)

        # salvar_button = driver.find_element(By.XPATH, "//*[(@id='btnSalvar_div')]")
        # salvar_button.click()
        # time.sleep(8)

except Exception as e:
    print(f"Erro: {e}")
    salvar_registro_error(row)

finally:
    driver.quit()
