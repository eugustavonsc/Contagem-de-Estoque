from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
import pyperclip
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv
import pandas as pd

# -----------------------------------------------------------------------------
# DOCUMENTAÇÃO: FUNÇÕES DE PREPARAÇÃO
# Estas funções preparam os dados para a automação.
# -----------------------------------------------------------------------------

def carregar_chassis(arquivo):
    """
    Lê um arquivo Excel e extrai uma lista de chassis.
    Procura por uma coluna que contenha 'chassi' no nome (ignorando maiúsculas/minúsculas).

    Args:
        arquivo (str): O caminho para o arquivo .xlsx ou .xls.

    Returns:
        list: Uma lista de strings, onde cada string é um chassi.

    Raises:
        FileNotFoundError: Se o arquivo especificado não for encontrado.
        ValueError: Se a coluna de chassi não for encontrada ou se o arquivo estiver vazio/com erro.
    """
    if not os.path.exists(arquivo):
        raise FileNotFoundError(f"Arquivo '{arquivo}' não encontrado!")
    try:
        df = pd.read_excel(arquivo, engine='openpyxl')
        col_chassi = next((col for col in df.columns if 'chassi' in col.lower()), None)
        if not col_chassi:
            raise ValueError("Coluna de chassi não encontrada no arquivo Excel!")
        chassis = df[col_chassi].dropna().astype(str).str.strip().tolist()
    except Exception as e:
        raise ValueError(f"Erro ao ler o arquivo Excel: {e}")
    if not chassis:
        raise ValueError("Nenhum chassi válido encontrado no arquivo!")
    return chassis

def selecionar_arquivo(titulo, tipos_de_arquivo):
    """
    Abre uma janela de diálogo para o usuário selecionar um arquivo.

    Args:
        titulo (str): O título da janela de diálogo.
        tipos_de_arquivo (list): Uma lista de tuplas definindo os tipos de arquivo permitidos.

    Returns:
        str: O caminho do arquivo selecionado, ou uma string vazia se nada for selecionado.
    """
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    file_path = filedialog.askopenfilename(title=titulo, filetypes=tipos_de_arquivo)
    root.destroy()
    return file_path

# -----------------------------------------------------------------------------
# DOCUMENTAÇÃO: INTERFACE GRÁFICA (Tkinter)
# Esta função gerencia a janela para o usuário carregar os arquivos.
# -----------------------------------------------------------------------------

def iniciar_interface():
    """
    Cria e exibe a interface gráfica para o usuário carregar o arquivo .env e o arquivo de chassis.
    A função armazena os dados coletados e os retorna somente após a janela ser fechada.

    Returns:
        dict: Um dicionário contendo os dados para a automação ('user_email', 'user_password', 'lista_chassis')
              e um status ('iniciar'). Retorna um dicionário vazio se a operação for cancelada.
    """
    # Dicionário para armazenar os dados que serão usados pela automação
    dados_automacao = {
        "user_email": "",
        "user_password": "",
        "lista_chassis": [],
        "iniciar": False # Flag para saber se o usuário clicou em "iniciar"
    }

    janela = tk.Tk()
    janela.title("Contagem de Estoque - Automação")
    janela.geometry("420x270")

    def carregar_env_manual():
        """Função interna para carregar o arquivo .env"""
        caminho_env = selecionar_arquivo(
            "Selecione o arquivo .env",
            [("Arquivo .env", ".env"), ("Todos os arquivos", "*.*")]
        )
        if caminho_env:
            load_dotenv(caminho_env, override=True)
            dados_automacao["user_email"] = os.getenv('USER_EMAIL', '')
            dados_automacao["user_password"] = os.getenv('USER_PASSWORD', '')
            label_usuario.config(text=f"Usuário: {dados_automacao['user_email']}")
            messagebox.showinfo(".env carregado", f"Arquivo .env carregado com sucesso!\nUsuário: {dados_automacao['user_email']}")
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo .env selecionado!")

    def iniciar_automacao_preparacao():
        """Função interna para preparar e fechar a janela para iniciar a automação."""
        # 1. Validar se o .env foi carregado
        if not dados_automacao["user_email"] or not dados_automacao["user_password"]:
            messagebox.showerror("Erro", "Usuário e senha não carregados! Por favor, selecione o arquivo .env.")
            return

        # 2. Pedir o arquivo de chassis
        arquivo_chassis = selecionar_arquivo(
            "Selecione o arquivo de chassis",
            [("Arquivos Excel", "*.xlsx;*.xls")]
        )
        if not arquivo_chassis:
            messagebox.showerror("Erro", "Nenhum arquivo de chassis selecionado!")
            return

        # 3. Carregar os chassis do arquivo
        try:
            dados_automacao["lista_chassis"] = carregar_chassis(arquivo_chassis)
            dados_automacao["iniciar"] = True # Sinaliza que tudo está pronto
            janela.destroy() # Fecha a janela para o script continuar
        except Exception as e:
            messagebox.showerror("Erro ao carregar arquivo", str(e))
            dados_automacao["iniciar"] = False

    # --- Widgets da Interface ---
    tk.Label(janela, text="Automação de Contagem de Estoque", font=("Arial", 14, "bold")).pack(pady=10)
    label_usuario = tk.Label(janela, text="Usuário: (Nenhum .env carregado)", font=("Arial", 10))
    label_usuario.pack(pady=5)
    tk.Button(janela, text="1. Selecionar arquivo .env", command=carregar_env_manual, font=("Arial", 11), bg="#2196F3", fg="white").pack(pady=5, padx=20, fill='x')
    tk.Button(janela, text="2. Selecionar arquivo de chassis e Iniciar", command=iniciar_automacao_preparacao, font=("Arial", 12), bg="#4CAF50", fg="white").pack(pady=20, padx=20, fill='x')
    tk.Label(janela, text="Certifique-se de preencher o .env com seu login.", font=("Arial", 9)).pack(side="bottom", pady=10)

    janela.mainloop() # Pausa o script aqui até a janela ser fechada

    # Após o mainloop terminar (janela.destroy() foi chamado), retorna os dados
    return dados_automacao

# -----------------------------------------------------------------------------
# DOCUMENTAÇÃO: LÓGICA DE AUTOMAÇÃO (Selenium)
# Esta função contém toda a lógica para interagir com o navegador.
# -----------------------------------------------------------------------------

def executar_automacao(user_email, user_password, lista_chassis):
    """
    Executa o processo de automação web com Selenium.

    Args:
        user_email (str): Email de login.
        user_password (str): Senha de login.
        lista_chassis (list): Lista de chassis a serem inseridos.
    """
    # --- Configuração do WebDriver ---
    print("Configurando o WebDriver...")
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-notifications')
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        actions = ActionChains(driver)
    except Exception as e:
        print(f"Erro crítico ao iniciar o WebDriver: {e}")
        print("Verifique sua conexão com a internet ou permissões de pasta.")
        return # Encerra a função se o driver não puder ser iniciado

    # --- Funções de Ação no Navegador ---

    def fazer_login():
        print("Acessando o site e realizando login...")
        driver.get('https://microworkcloud.com.br')
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'Username'))).send_keys(user_email)
        driver.find_element(By.ID, 'Password').send_keys(user_password)
        driver.find_element(By.NAME, 'entrar').click()
        # É uma boa prática esperar por um elemento da próxima página para confirmar o login
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        print("Login realizado com sucesso.")

    # ... (o resto das suas funções de automação permanecem as mesmas)
    def acessar_contagem_estoque():
        input("\nAcesse manualmente a tela de contagem de estoque e pressione ENTER para continuar...")

    def encontrar_campo_chassi():
        estrategias = [
            (By.CSS_SELECTOR, 'input[formcontrolname="Chassi"]'),
            (By.XPATH, '//input[contains(@placeholder, "código ou nome")]'),
            (By.XPATH, '//input[contains(@id, "chassi") or contains(@name, "chassi")]'),
            (By.CLASS_NAME, 'mw-input')
        ]
        for by, value in estrategias:
            try:
                return WebDriverWait(driver, 5).until(EC.presence_of_element_located((by, value)))
            except:
                continue
        raise Exception("Campo de chassi não encontrado após tentar várias estratégias")

    def registrar_chassi():
        # Lógica para registrar o chassi (pressionar TAB ou Enter, por exemplo)
        # Usar Keys.ENTER ou Keys.TAB é geralmente mais confiável
        try:
            actions.send_keys(Keys.ENTER).perform()
            return True
        except Exception as e:
            print(f"Não foi possível registrar com ENTER: {e}")
            return False

    def inserir_chassi(chassi):
        try:
            campo = encontrar_campo_chassi()
            # Usar pyperclip é bom, mas send_keys é mais nativo do Selenium
            campo.clear()
            campo.send_keys(chassi)
            time.sleep(0.5) # Pequena pausa para o sistema reagir
            if not registrar_chassi():
                print(f"Aviso: Não foi possível confirmar o chassi {chassi} automaticamente")
            time.sleep(0.8)
        except Exception as e:
            print(f"Erro ao inserir chassi {chassi}: {str(e)}")
            raise

    def clicar_botao_salvar():
        print("Aguardando o botão 'Salvar' ficar habilitado...")
        try:
            salvar_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Salvar'] and not(@disabled)]"))
            )
            salvar_btn.click()
            print("✅ Botão 'Salvar' clicado com sucesso.")
        except:
            print("❌ Erro: botão 'Salvar' não ficou habilitado a tempo.")

    # --- Execução Principal da Automação ---
    try:
        print("\nIniciando automação no navegador...")
        print(f"{len(lista_chassis)} chassis carregados para inserção.")
        fazer_login()
        acessar_contagem_estoque()
        print(f"\nIniciando inserção de {len(lista_chassis)} chassis...")
        for i, chassi in enumerate(lista_chassis, 1):
            print(f"{i}/{len(lista_chassis)} - Inserindo: {chassi}")
            inserir_chassi(chassi)
        print("\nTodos os chassis foram inseridos.")
        clicar_botao_salvar()
        print("\nProcesso concluído com sucesso!")
        time.sleep(10) # Aguarda 10 segundos antes de fechar o navegador
    except Exception as e:
        print(f"\nOcorreu um erro crítico durante a automação: {str(e)}")
        nome_arquivo_erro = 'erro_automacao.png'
        driver.save_screenshot(nome_arquivo_erro)
        print(f"Um screenshot do erro foi salvo como '{nome_arquivo_erro}'")
    finally:
        print("Fechando o navegador.")
        driver.quit()

# -----------------------------------------------------------------------------
# DOCUMENTAÇÃO: PONTO DE ENTRADA DO SCRIPT
# É aqui que o programa começa a ser executado.
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    # 1. Chama a interface para coletar os dados do usuário.
    # O script ficará pausado aqui até o usuário fechar a janela.
    print("Abrindo interface para configuração...")
    dados = iniciar_interface()

    # 2. Verifica se o usuário clicou em "Iniciar" e se os dados foram coletados.
    if dados and dados.get("iniciar"):
        print("\nConfiguração concluída. Preparando para iniciar a automação...")
        # 3. Inicia a automação com os dados que foram retornados pela interface.
        executar_automacao(
            dados["user_email"],
            dados["user_password"],
            dados["lista_chassis"]
        )
    else:
        # Caso o usuário feche a janela sem clicar em iniciar.
        print("\nOperação cancelada pelo usuário. O programa será encerrado.")