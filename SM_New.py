import sys
from pathlib import Path
import flet as ft
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui as py
import os
from datetime import datetime
import chromedriver_autoinstaller
import base64
from dotenv import load_dotenv
import time
import threading
import webbrowser
import numpy as np
from Dropdown import Status_Selection


# Detecta se está rodando dentro de um executável PyInstaller
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).resolve().parent

# Pastas incluídas no executável
CAMINHO_PLANILHAS = BASE_DIR / 'planilhas'
CAMINHO_DADOS = BASE_DIR / 'dados'

# Pasta onde os arquivos de saída serão criados (mesmo diretório do .exe)
CAMINHO_SAIDA = Path(__file__).resolve().parent / 'saida'

# Função de log para diagnóstico
def save_log(message):
    log_path = os.path.join(os.path.expanduser("~"), "robos_log.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {message}\n")

def main(page: ft.Page):
    # Configurações da janela - Centralizando na tela
    page.window.width = page.window.max_width = page.window.min_width = 375
    page.window.height = page.window.max_height = page.window.min_height = 500
    page.window.center()  # This will center the window on the screen
    page.window.title_bar_hidden = True
    page.adaptive, page.padding = True, 0
    page.window.icon = os.path.join(os.path.dirname(__file__), "assets", "desk_logo.ico")
    page.title = "StatusManager"
    
    # Contador de progresso
    progress_counter = ft.Text(
        value="0/0",
        color=ft.Colors.WHITE,
        size=12,
        weight=ft.FontWeight.NORMAL
    )
    
    # Criando campos de texto
    login_field = ft.TextField(
        label='Login',
        border_color=ft.Colors.WHITE,
        label_style=ft.TextStyle(color=ft.Colors.WHITE, size=12),
        color=ft.Colors.WHITE,
        prefix_icon=ft.Icons.PERSON,
        width=280
    )
    
    password_field = ft.TextField(
        label='Senha',
        border_color=ft.Colors.WHITE,
        label_style=ft.TextStyle(color=ft.Colors.WHITE, size=12),
        color=ft.Colors.WHITE,
        password=True,
        prefix_icon=ft.Icons.LOCK,
        width=280
    )
    
    # Dropdown de status
    status_dropdown = Status_Selection.dropdown()
    
    # Checkbox de "Lembrar Usuário e Senha"
    remember_checkbox = ft.Checkbox(
        label="Lembrar Usuário e Senha",
        label_style=ft.TextStyle(size=12, color=ft.Colors.WHITE),
        value=False
    )
    
    # Função para salvar credenciais no .env quando "Lembrar" estiver marcado
    def save_credentials(username, password, remember=False):
        if remember and username and password:
            userdir_get = os.path.expanduser("~")
            save_log(f"Tentando salvar credenciais para o usuário: {username}")
            
            # Tentar vários possíveis caminhos para o OneDrive
            possible_paths = [
                os.path.join(userdir_get, "OneDrive - Desktop Sigmanet", "Área de Trabalho", "Robos"),
                os.path.join(userdir_get, "OneDrive", "Área de Trabalho", "Robos"),
                os.path.join(userdir_get, "Desktop", "Robos"),
                os.path.join(userdir_get, "Área de Trabalho", "Robos"),
                os.path.join(userdir_get, "OneDrive - Desktop SA","Área de Trabalho", "Robos"),
                os.path.join(userdir_get, "OneDrive\\Área de Trabalho", "Robos")
            ]
            
            robos_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    robos_path = path
                    save_log(f"Diretório Robos encontrado: {robos_path}")
                    break
            
            if robos_path is None:
                try:
                    # Cria um diretório padrão se nenhum for encontrado
                    robos_path = os.path.join(userdir_get, "Desktop", "Robos")
                    os.makedirs(robos_path, exist_ok=True)
                    save_log(f"Diretório Robos criado: {robos_path}")
                except Exception as e:
                    error_msg = f"Erro ao criar diretório para credenciais: {str(e)}"
                    save_log(error_msg)
                    py.alert("Erro: Não foi possível encontrar o diretório de trabalho")
                    return False
            
            logins_dir = os.path.join(robos_path, ".Logins")
            if not os.path.exists(logins_dir):
                try:
                    os.makedirs(logins_dir)
                    save_log(f"Diretório .Logins criado: {logins_dir}")
                except Exception as e:
                    error_msg = f"Erro ao criar diretório .Logins: {str(e)}"
                    save_log(error_msg)
                    return False
            
            login_encoded = base64.b64encode(username.encode('utf-8')).decode('utf-8')
            password_encoded = base64.b64encode(password.encode('utf-8')).decode('utf-8')
            env_path = os.path.join(logins_dir, ".env")
            
            existing_vars = {}
            if os.path.exists(env_path):
                try:
                    with open(env_path, "r") as f:
                        for line in f:
                            line = line.strip()
                            if line and not line.startswith('#'):
                                if '=' in line:
                                    key, value = line.split('=', 1)
                                    existing_vars[key] = value
                except Exception as e:
                    error_msg = f"Erro ao ler arquivo .env existente: {str(e)}"
                    save_log(error_msg)
            
            existing_vars["LOGINADM"] = login_encoded
            existing_vars["SENHAADM"] = password_encoded
            
            try:
                with open(env_path, "w") as f:
                    for key, value in existing_vars.items():
                        f.write(f"{key}={value}\n")
                save_log("Credenciais salvas com sucesso")
                return True
            except Exception as e:
                error_msg = f"Erro ao salvar credenciais: {str(e)}"
                save_log(error_msg)
                return False
        return False
    
    # Função para carregar credenciais salvas
    def load_saved_credentials():
        userdir_get = os.path.expanduser("~")
        save_log("Tentando carregar credenciais salvas")
        
        possible_logins_paths = [
            os.path.join(userdir_get, "OneDrive - Desktop Sigmanet", "Área de Trabalho", "Robos", ".Logins"),
            os.path.join(userdir_get, "OneDrive", "Área de Trabalho", "Robos", ".Logins"),
            os.path.join(userdir_get, "Desktop", "Robos", ".Logins"),
            os.path.join(userdir_get, "Área de Trabalho", "Robos", ".Logins"),
            os.path.join(userdir_get, "OneDrive - Desktop SA","Área de Trabalho", "Robos", ".Logins"),
            os.path.join(userdir_get, "OneDrive\\Área de Trabalho", "Robos", ".Logins")
        ]
        
        for path in possible_logins_paths:
            save_log(f"Verificando pasta de credenciais: {path}")
            env_path = os.path.join(path, ".env")
            if os.path.exists(env_path):
                save_log(f"Arquivo .env encontrado: {env_path}")
                try:
                    load_dotenv(dotenv_path=env_path, override=True)
                    login_adm = os.getenv("LOGINADM")
                    senha_adm = os.getenv("SENHAADM")
                    
                    if login_adm and senha_adm:
                        try:
                            login_decoded = base64.b64decode(login_adm).decode('utf-8')
                            senha_decoded = base64.b64decode(senha_adm).decode('utf-8')
                            login_field.value = login_decoded
                            password_field.value = senha_decoded
                            remember_checkbox.value = True
                            save_log("Credenciais carregadas com sucesso")
                            return True
                        except Exception as e:
                            error_msg = f"Erro ao decodificar credenciais: {str(e)}"
                            save_log(error_msg)
                except Exception as e:
                    error_msg = f"Erro ao carregar arquivo .env: {str(e)}"
                    save_log(error_msg)
        
        save_log("Nenhuma credencial salva encontrada")
        return False
    
    # Função para abrir o WhatsApp
    def open_whatsapp(e):
        # Substitua o número abaixo pelo seu número de WhatsApp no formato internacional
        webbrowser.open('https://wa.me/5519920026971')
        save_log("Link do WhatsApp aberto")
    
    # Função para abrir a pasta do projeto
    def open_project_folder(e):
        userdir_get = os.path.expanduser("~")
        save_log("Tentando abrir pasta do projeto")
        
        possible_paths = [
            os.path.join(userdir_get, "OneDrive - Desktop Sigmanet", "Área de Trabalho", "Robos", "Code", "Alteração"),
            os.path.join(userdir_get, "OneDrive", "Área de Trabalho", "Robos", "Code", "Alteração"),
            os.path.join(userdir_get, "Desktop", "Robos", "Code", "Alteração"),
            os.path.join(userdir_get, "Área de Trabalho", "Robos", "Code", "Alteração"),
            os.path.join(userdir_get, "OneDrive - Desktop SA","Área de Trabalho", "Robos", "Code", "Alteração"),
            os.path.join(userdir_get, "OneDrive\\Área de Trabalho", "Robos", "Code", "Alteração")
        ]
        
        for path in possible_paths:
            save_log(f"Verificando pasta do projeto: {path}")
            if os.path.exists(path):
                os.startfile(path)
                save_log(f"Pasta do projeto aberta: {path}")
                return
        
        # Se não encontrou, tenta criar
        try:
            default_path = os.path.join(userdir_get, "Desktop", "Robos", "Code", "Alteração")
            os.makedirs(default_path, exist_ok=True)
            os.startfile(default_path)
            save_log(f"Pasta do projeto criada e aberta: {default_path}")
        except Exception as e:
            error_msg = f"Erro ao abrir pasta do projeto: {str(e)}"
            save_log(error_msg)
            py.alert("Erro: Não foi possível encontrar ou criar o diretório do projeto")
    
    # Função para atualizar o contador de progresso na interface
    # Função para atualizar a barra e o contador de progresso na interface
    def update_progress(current, total):
        # Atualiza o texto do contador
        progress_counter.value = f"{current}/{total}"
        
        # Atualiza a barra de progresso
        if total > 0:
            progress_bar.value = current / total
        else:
            progress_bar.value = 0
            
        page.update()
    
    # Função integrada de login que usa os campos da interface
    def login_function(e):
        # Executar em uma thread separada para manter a interface responsiva
        threading.Thread(target=process_login, daemon=True).start()
    
    def process_login():
        username = login_field.value
        password = password_field.value
        selected_status = status_dropdown.value
    
        if not selected_status:
            py.alert("Erro: É necessário escolher uma opção: 'Disponível', 'Defeito/Inutilizado' ou 'Procurar'")
            return
        
        userdir_get = os.path.expanduser("~")
        save_log(f"Diretório do usuário: {userdir_get}")
        
        # Tentar vários possíveis caminhos para o OneDrive
        possible_paths = [
            os.path.join(userdir_get, "OneDrive - Desktop Sigmanet", "Área de Trabalho", "Robos"),
            os.path.join(userdir_get, "OneDrive", "Área de Trabalho", "Robos"),
            os.path.join(userdir_get, "Desktop", "Robos"),
            os.path.join(userdir_get, "Área de Trabalho", "Robos"),
            os.path.join(userdir_get, "OneDrive - Desktop SA","Área de Trabalho", "Robos"),
            os.path.join(userdir_get, "OneDrive\\Área de Trabalho", "Robos")
        ]

        script_path = None
        for path in possible_paths:
            save_log(f"Verificando caminho: {path} - Existe: {os.path.exists(path)}")
            if os.path.exists(path):
                script_path = path
                save_log(f"Caminho encontrado: {script_path}")
                break

        if script_path is None:
            # Se não encontrou o caminho, criar a pasta
            try:
                default_path = os.path.join(userdir_get, "Desktop", "Robos")
                os.makedirs(default_path, exist_ok=True)
                script_path = default_path
                save_log(f"Criado caminho padrão: {script_path}")
            except Exception as e:
                error_msg = f"Erro ao criar diretório: {str(e)}"
                save_log(error_msg)
                py.alert(f'''Erro: Não foi possível encontrar ou criar o diretório
                         Crie ou mova esta pasta para uma na pasta de trabalho com o nome "Robos"''')
                return
        
        # Garantir que o diretório de validação existe
        validation_dir = os.path.join(script_path, "Code", "Alteração")
        if not os.path.exists(validation_dir):
            try:
                os.makedirs(validation_dir, exist_ok=True)
                save_log(f"Criado diretório de validação: {validation_dir}")
            except Exception as e:
                error_msg = f"Erro ao criar diretório de validação: {str(e)}"
                save_log(error_msg)
                py.alert("Erro: Não foi possível criar o diretório de validação")
                return
        
        excel_path = os.path.join(validation_dir, "Dados.xlsx")
        save_log(f"Caminho do Excel: {excel_path}")
        
        if not os.path.exists(excel_path):
            try:
                # Criar um DataFrame vazio
                df = pd.DataFrame(columns=[
                    "Ativo/Mac/Serie", "Tipo", "Horario Processado", "Ativo", "Antigo Status", "Observação"
                ])
                df.to_excel(excel_path, index=False)
                save_log(f"Arquivo Excel criado: {excel_path}")
            except Exception as e:
                error_msg = f"Erro ao criar arquivo Excel: {str(e)}"
                save_log(error_msg)
                py.alert(f"Erro ao criar o arquivo Excel: {str(e)}")
                return
        else:
            try:
                with open(excel_path, 'r+b') as f:
                    pass
            except IOError:
                py.alert("Erro: O arquivo 'Dados.xlsx' está aberto.\nPor favor, salve e feche o arquivo antes de continuar.")
                return
        
        save_credentials(username, password, remember_checkbox.value)
        
        try:
            save_log("Iniciando processo de automação com Chrome")
            chromedriver_autoinstaller.install()
            options = webdriver.ChromeOptions()
            options.add_experimental_option("detach", True)
            options.add_argument("--start-maximized")
            options.add_argument("--allow-running-insecure-content")
            options.add_argument("--allow-insecure-localhost")
            # options.add_argument("--headless")
            chromedriver = webdriver.Chrome(options=options)
            wait = WebDriverWait(chromedriver, 30)

            #----------------------------------------------------------------------------------------------Alterar essa Função------------------------------------------------------------------------------------------------------
            
            if not username or not password:
                save_log("Credenciais não fornecidas, tentando carregar do arquivo")
                logins_dir = os.path.join(script_path, ".Logins")
                if os.path.exists(logins_dir):
                    os.chdir(logins_dir)
                    load_dotenv(dotenv_path=".env", override=True)
                    login_adm = os.getenv("LOGINADM")
                    senha_adm = os.getenv("SENHAADM")
                    
                    if login_adm and senha_adm:
                        login_decoded = base64.b64decode(login_adm).decode('utf-8')
                        senha_decoded = base64.b64decode(senha_adm).decode('utf-8')
                        save_log("Credenciais carregadas do arquivo")
                    else:
                        save_log("Credenciais não encontradas no arquivo")
                        py.alert("Erro: Credenciais não encontradas")
                        chromedriver.quit()
                        return
                else:
                    save_log("Diretório de credenciais não encontrado")
                    py.alert("Erro: Credenciais não fornecidas")
                    chromedriver.quit()
                    return
            else:
                login_decoded = username
                senha_decoded = password
                
            os.chdir(validation_dir)
            
            try:
                tabela = pd.read_excel("Dados.xlsx")
                save_log(f"Arquivo Excel lido com sucesso: {len(tabela)} linhas")
            except Exception as e:
                error_msg = f"Erro ao ler arquivo Excel: {str(e)}"
                save_log(error_msg)
                py.alert(f"Erro ao ler o arquivo Excel: {str(e)}")
                chromedriver.quit()
                return
            
            # Contar linhas com dados (excluindo cabeçalho)
            total_rows = len(tabela)
            # Inicializar contador de progresso
            current_row = 0
            # Atualizar o contador na interface
            update_progress(current_row, total_rows)
            
            chromedriver.get("http://adm.desktop.com.br/login.jsp?")   
            save_log("Acessando o SSO")
            # wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[1]/td[2]/font/input'))).send_keys(login_decoded)
            # wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[2]/font/input'))).send_keys(senha_decoded)
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/table[2]/tbody/tr[9]/td/button"))).click()
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[inputmode="email"]'))).send_keys(login_decoded)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="password"]'))).send_keys(senha_decoded)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))).click()

            save_log("Aguardando o input do usuário...")
            
            current_url = chromedriver.current_url
            if current_url.startswith("https://desktop.sso.e-trust.com.br/user/dashboard/"):
                pass
                save_log("Já na página dashboard, continuando...")
            
            else:
            
                wait.until(EC.url_matches('https://desktop.sso.e-trust.com.br/mfa/login/validate'))
                

            # chromedriver.execute_script("alert('Por gentileza, insira o código no campo solicitado abaixo. (Você terá 5 minutos)')")

            # py.alert(
            #     text='Por gentileza, insira o código de verificação na pagina. Se você preferir, solicite o código pelo email',
            #     title='Insira o código de verificação:',
            #     button='Combinado!'
            # )

            # chromedriver.find_element("xpath", '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[1]/td[2]/font/input').send_keys(login_decoded)
            # chromedriver.find_element("xpath", '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[2]/font/input').send_keys(senha_decoded)
            # chromedriver.find_element(By.XPATH, "//input[@type='Image']").click()

            
                



            try:
                wait = WebDriverWait(chromedriver, 300)
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'img[onclick="window.location.href=\'https://adm.desktop.com.br/Login_SSO.jsp\';"]'))).click()
                wait.until(EC.url_matches("https://adm.desktop.com.br/menu.jsp"))
                wait = WebDriverWait(chromedriver, 30)
            except Exception as e:
                save_log(f"Houve um erro durante o login: {e}")
                wait = WebDriverWait(chromedriver, 30)

                chromedriver.quit()
                exit()

            save_log("Login realizado, aguardando carregamento da página")
            chromedriver.get('https://adm.desktop.com.br/Ativos.jsp')
            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[4]/td/table/tbody/tr/td/font/a[4]").click()
            wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/input[2]"))).click()
            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/input[2]").click()
            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/input'))).click()
            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/input").click()
            save_log("Navegação inicial concluída, iniciando processamento dos itens")
            
            total_rows = len(tabela[tabela["Horario Processado"] != 'nan'])
            current_row = 0
            
            
            
            
            
            
            for i, série in enumerate(tabela["Ativo/Mac/Serie"]):
                if tabela.at[i, "Horario Processado"] == 'nan':
                    continue
                
                tipo = tabela.at[i, "Tipo"] if "Tipo" in tabela.columns else ""
                tipo = tipo.upper() if not pd.isnull(tipo) else ""
                hour = tabela.at[i, "Horario Processado"] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                ativo = tabela.at[i, "Ativo"] if "Ativo" in tabela.columns else ""
                
                save_log(f"Processando item {i+1}/{total_rows}: {série}, Tipo: {tipo}")
                
                if pd.isnull(série):
                    save_log(f"Item {i+1} sem série, pulando")
                    
                    continue
                
                try:
                    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/input'))).clear()
                    
                    # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/input").clear()
                    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/input'))).send_keys(série)
                    # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/input").send_keys(série)
                    
                    if tipo == "MAC":
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[2]'))).click()
                        # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[2]").click()
                    
                    elif tipo == "SERIE":
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[3]'))).click()
                        # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[3]").click()
                    
                    elif tipo == "ATIVO":
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[1]'))).click()
                        # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[1]").click()
                        
                    elif tipo == "FSAN":
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr/td/font/select[1]/option[4]'))).click()
                    
                    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/input'))).click()
                    # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/input").click()
                    
                    
                    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td/font'))).text
                    
                    try:
                        nao_encontrado = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td/font").text
                        if nao_encontrado == "Total de ativos atendendo aos requisitos da pesquisa: 0":
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            situacao = tabela.at[i, "Antigo Status"] = "Não encontrado"
                            hour = tabela.at[i, "Horario Processado"]
                            tabela.to_excel("Dados.xlsx", index=False)
                            save_log(f"Item {i+1} não encontrado")
                            
                            # Atualizar contador e progresso
                            current_row += 1
                            update_progress(current_row, total_rows)
                            continue
                    except:
                        pass
                    
                    try:
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[3]/font[1]')))
                        status_atual = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[3]/font[1]").text
                        
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[3]/font[1]')))
                        status_atual = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[3]/font[1]").text
                        save_log(f"Status atual do item {i+1}: {status_atual}")

                        # NOVA LÓGICA: Verificar se o status é "Alocado" e pular
                        if "Alocado" in status_atual:
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b')))
                            numero_ativo = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b").text
                            
                            tabela.at[i, "Ativo"] = numero_ativo
                            tabela.at[i, "Antigo Status"] = status_atual
                            tabela.to_excel("Dados.xlsx", index=False)
                            save_log(f"Item {i+1} está alocado, pulando")
                            
                            # Volta para a página de pesquisa para processar o próximo item
                            wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            
                                
                                
                                
                            # Atualizar contador e progresso
                            current_row += 1
                            update_progress(current_row, total_rows)
                            continue  # Pula para o próximo equipamento na lista
                        
                        if selected_status == "Disponível" and "Disponível" in status_atual:
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b')))
                            numero_ativo = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b").text
                            
                            tabela.at[i, "Ativo"] = numero_ativo
                            tabela.at[i, "Antigo Status"] = status_atual
                            obs = tabela.at[i, "Observação"]
                            obs = np.nan if str(obs).lower() == 'nan' else obs

                            
                            # Verifica se existe observação a ser adicionada
                            if pd.notna(obs) and obs.strip() != "":
                                # Clica para editar o equipamento
                                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[4]/a/img'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[4]/a/img").click()
                                
                                # Adiciona a observação
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input'))).clear()
                                # chromedriver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input').clear()
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input'))).send_keys(obs)
                                # chromedriver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input').send_keys(obs)
                                
                                # Clica para salvar as alterações
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/input'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/input").click()
                                time.sleep(2)
                                
                                # Volta para a página de pesquisa
                                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            else:
                                # Se não tem observação, apenas volta para a página de pesquisa
                                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            
                            tabela.at[i, "Novo Status"] = status_atual  # Manter o mesmo status
                            tabela.to_excel("Dados.xlsx", index=False)
                            
                            # Atualizar contador e progresso
                            current_row += 1
                            update_progress(current_row, total_rows)
                            continue  # Pula para o próximo equipamento na lista

                        if selected_status == "Defeito/Inutilizado" and "Defeito/Inutilizado" in status_atual:
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b')))
                            numero_ativo = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b").text
                            
                            tabela.at[i, "Ativo"] = numero_ativo
                            tabela.at[i, "Antigo Status"] = status_atual
                            obs = str(tabela.at[i, "Observação"])
                            obs = np.nan if str(obs).lower() == 'nan' else obs
                            
                            # Verifica se existe observação a ser adicionada
                            if pd.notna(obs) and obs.strip() != "":
                                # Clica para editar o equipamento
                                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[4]/a/img'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[4]/a/img").click()
                                
                                # Adiciona a observação
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input'))).clear()
                                # chromedriver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input').clear()
                                chromedriver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input').send_keys(obs)
                                
                                # Clica para salvar as alterações
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/input'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/input").click()
                                time.sleep(2)
                                
                                # Volta para a página de pesquisa
                                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            else:
                                # Se não tem observação, apenas volta para a página de pesquisa
                                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                                # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            
                            tabela.at[i, "Novo Status"] = status_atual  # Manter o mesmo status
                            tabela.to_excel("Dados.xlsx", index=False)
                            
                            # Atualizar contador e progresso
                            current_row += 1
                            update_progress(current_row, total_rows)
                            continue  # Pula para o próximo equipamento na lista
                        
                        
                        
                        
                        
                        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b')))
                        numero_ativo = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[1]/font/b").text
                        
                        tabela.at[i, "Ativo"] = numero_ativo
                        tabela.at[i, "Antigo Status"] = status_atual
                        obs = str(tabela.at[i, "Observação"])
                        obs = np.nan if str(obs).lower() == 'nan' else obs

                        

                        
                        
                        if selected_status == "Procurar":
                            wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img')))
                            chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                        else:
                            wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[4]/a/img')))
                            chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[4]/a/img").click()
                            
                            status_value = "L"  # Valor padrão para "Disponível"
                            if selected_status == "Defeito/Inutilizado":
                                status_value = "D"  # Valor para "Defeito/Inutilizado"
                            if selected_status == "Baixado/Perdido":
                                status_value = "B"  # Valor para "Baixado/Perdido"
                                
                                
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[1]/td[5]/font/select'))).send_keys(status_value)
                            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[1]/td[5]/font/select").send_keys(status_value)
                            
                            #path obs adm : /html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input
                            if pd.isnull(obs) or str(obs).strip().lower() in ['nan', 'null', 'none', '']:
                                pass

                                
                            else:
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input'))).clear()
                                # chromedriver.find_element(By.XPATH,'/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input').clear()
                                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input'))).send_keys(obs)
                                # chromedriver.find_element(By.XPATH,'/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input').send_keys(obs)
                            
                            
                            # wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input')))
                            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[2]/td[1]/font/input").send_keys(obs)
                            
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/input'))).click()
                            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/input").click()
                            time.sleep(2)

                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[3]/font[1]')))
                            status_new = chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[6]/td/table/tbody/tr[6]/td[3]/font[1]").text
                            tabela.at[i, "Novo Status"] = status_new
                            
                            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img'))).click()
                            # chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                            
                        tabela.to_excel("Dados.xlsx", index=False)
                        
                        # Atualizar contador e progresso
                        current_row += 1
                        update_progress(current_row, total_rows)
                        
                    except Exception as e:
                        # py.alert(f"Erro ao processar o item {série}: {str(e)}")
                        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img')))
                        chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                        # Atualizamos o contador mesmo em caso de erro
                        current_row += 1
                        update_progress(current_row, total_rows)
                        continue
                        
                except Exception as e:
                    # py.alert(f"Erro ao processar o item {série}: {str(e)}")
                    try:
                        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]/tbody/tr[8]/td/a[1]/img')))
                        chromedriver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[8]/td/a[1]/img").click()
                    except:
                        pass
                    # Atualizamos o contador mesmo em caso de erro
                    current_row += 1
                    update_progress(current_row, total_rows)
                    continue
            
            open_xlsx = os.path.join(script_path, "Code", "Alteração", "Dados.xlsx")
            py.alert("Após confirmar, você será redirecionado para o arquivo!")
            chromedriver.close()
            os.startfile(open_xlsx)
        except Exception as e:
            py.alert(f"Erro ao executar o script: {str(e)}")
            try:
                chromedriver.quit()
            except:
                pass
    
    # Botão de pasta apenas com ícone
    botao_pasta = ft.IconButton(
        icon=ft.Icons.FOLDER,
        icon_color=ft.Colors.WHITE,
        tooltip="Clique aqui para abrir a pasta aonde contem o arquivo excel.",
        on_click=open_project_folder
    )
    
    # Botão de WhatsApp
    
    whatsapp_button = ft.Container(
        content=ft.IconButton(
            icon=ft.Icons.CONTACT_SUPPORT,
            tooltip="Clique aqui para tirar dúvidas ou ver mais informações.",
            icon_color=ft.Colors.WHITE,
            icon_size=20,
            on_click=open_whatsapp
        ),
        bottom=-3,
        left=-1,
       
    )
    
    # Linha contendo o dropdown de status e o botão de abrir pasta
    status_pasta_row = ft.Row(
        controls=[
            status_dropdown,
            botao_pasta
        ],
        alignment=ft.MainAxisAlignment.START,
        spacing=20,
        width=290
    )
    
    # Container para a linha com o dropdown e o botão
    status_pasta_container = ft.Container(
        content=status_pasta_row,
        margin=ft.margin.only(left=30, top=300)
    )
    
    # Botão "Fazer Login"
    
    login_button = ft.ElevatedButton(
        text="Fazer Login",
        style=ft.ButtonStyle(
            bgcolor='#0e1111',
            color=ft.Colors.YELLOW_700,
            text_style=ft.TextStyle(size=15)
        ),
        on_click=login_function,
        width=325,
        height=45,
    )
    
    # Container para o contador de progresso
    # Barra de progresso
    progress_bar = ft.ProgressBar(
        width=220,
        color=ft.colors.YELLOW_700,
        bgcolor="#555555",
        value=0
    )

    # Contador de progresso
    progress_counter = ft.Text(
        value="0/0",
        color=ft.Colors.WHITE,
        size=12,
        weight=ft.FontWeight.NORMAL
    )    # margin=ft.margin.only(left=280, top=350),
        # alignment=ft.alignment.center
    
    

    # Adicione um texto informativo para o contador e reposicione:
# Primeiro, crie o texto informativo
    # Adicione um texto informativo para o progresso
    info_text = ft.Text(
        value="Progresso:",
        color=ft.Colors.WHITE,
        size=12,
        weight=ft.FontWeight.NORMAL
    )

    # Container para a barra de progresso
    progress_bar_container = ft.Container(
        content=progress_bar,
        margin=ft.margin.only(top=5)
    )

    # Coluna para os elementos de progresso
    progress_column = ft.Column(
        controls=[
            # Linha com o texto "Progresso:" e o contador
            ft.Row(
                controls=[
                    info_text,
                    progress_counter
                ],
                spacing=5,
                alignment=ft.MainAxisAlignment.START
            ),
            # Barra de progresso logo abaixo
            progress_bar_container
        ],
        spacing=0,
        horizontal_alignment=ft.CrossAxisAlignment.START
    )

    # Container principal para o progresso
    progress_counter_container = ft.Container(
        content=progress_column,
        margin=ft.margin.only(left=30, top=365)
    )

    # Container para aplicar a margem no botão de login
    login_button_container = ft.Container(
        content=login_button,
        margin=ft.margin.only(left=18, top=415)
    )
    
    # Função para animar o botão de fechar
    def hover_close(e):
        close.content.opacity = 1 if e.data == "true" else 0
        close.content.update()
    
    # Botão de fechar
    close = ft.Container(
        margin=ft.margin.only(left=335),
        content=ft.Icon(
            ft.Icons.CLOSE,
            color=ft.Colors.WHITE,
            opacity=0,
            animate_opacity=ft.Animation(duration=600, curve=ft.AnimationCurve.EASE_IN_OUT)
        ),
        ink=True,
        on_hover=hover_close,
        on_click=lambda _: page.window.close()
    )
    
    # Layout principal com uma pilha de elementos - Otimizado
    stack_main = ft.Stack(
        expand=True,
        controls=[
            # Background image
            ft.Container(
                content=ft.Image(
                    src=os.path.join(os.path.dirname(__file__), "assets", "Images", "bgdesktop.png"),
                    width=400,
                    height=550,
                    fit=ft.ImageFit.COVER,
                )
            ),
            # Footer
            ft.Container(
                content=ft.Text(
                    "© 2025 Desktop | Análises Logística | v1.0.0",
                    color=ft.Colors.WHITE,
                    size=12,
                    text_align=ft.TextAlign.CENTER,
                ),
                width=375,
                height=20,
                left=0,
                bottom=5,
                alignment=ft.alignment.center,
            ),
            # Logo
                
            ft.Container(
                content=ft.Image(
                    src=os.path.join(os.path.dirname(__file__), "assets", "Images", "logo.png"),
                    width=250,
                    height=75,
                    fit=ft.ImageFit.CONTAIN
                ),
                margin=ft.margin.only(top=25, left=47),
                
            ),
            ft.WindowDragArea(
                width=375, 
                height=100,
                content=ft.Container(bgcolor='transparent')
            ),
            close,
            whatsapp_button,
            # Fields in containers
            ft.Container(width=280, height=50, margin=ft.margin.only(left=30, top=125), content=login_field),
            ft.Container(width=280, height=50, margin=ft.margin.only(left=30, top=185), content=password_field),
            ft.Container(margin=ft.margin.only(left=30, top=245), content=remember_checkbox),
            status_pasta_container,
            progress_counter_container,
            login_button_container,
        ]
    )
    
    # Carregar credenciais salvas ao iniciar
    if load_saved_credentials():
        page.update()
    
    # Adiciona o conteúdo à página de forma otimizada
    page.add(stack_main)

if __name__ == "__main__":
    ft.app(target=main, assets_dir='assets')
    
    # '