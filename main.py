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
import tink
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv


def carregar_chassis(arquivo):
    if not os.path.exists(arquivo):
        raise FileNotFoundError(f"Arquivo '{arquivo}' não encontrado!")

    with open(arquivo, 'r', encoding='utf-8') as f:
        chassis = [line.strip() for line in f.readlines() if line.strip()]

    if not chassis:
        raise ValueError("Nenhum chassi válido encontrado no arquivo!")

    return chassis

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo de chassis",
        filetypes=[("Arquivos de texto", "*.txt")]
    )
    root.destroy()
    return file_path

def selecionar_env():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo .env",
        filetypes=[("Arquivo .env", ".env"), ("Todos os arquivos", "*.*")]
    )
    root.destroy()
    return file_path

# Função para interface principal
def iniciar_interface():
    env_path = None
    user_email = ''
    user_password = ''

    def carregar_env_manual():
        nonlocal env_path, user_email, user_password
        env_path = selecionar_env()
        if env_path:
            load_dotenv(env_path, override=True)
            user_email = os.getenv('USER_EMAIL', '')
            user_password = os.getenv('USER_PASSWORD', '')
            label_usuario.config(text=f"Usuário: {user_email}")
            messagebox.showinfo(".env carregado", f"Arquivo .env carregado com sucesso!\nUsuário: {user_email}")
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo .env selecionado!")

    def iniciar_automacao():
        arquivo = selecionar_arquivo()
        if not arquivo:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
            return
        if not user_email or not user_password:
            messagebox.showerror("Erro", "Usuário e senha não carregados do .env!")
            return
        try:
            lista_chassis = carregar_chassis(arquivo)
        except Exception as e:
            messagebox.showerror("Erro ao carregar arquivo", str(e))
            return
        janela.destroy()
        executar_automacao(user_email, user_password, lista_chassis)

    janela = tk.Tk()
    janela.title("Contagem de Estoque - Automação")
    janela.geometry("420x270")
    tk.Label(janela, text="Automação de Contagem de Estoque", font=("Arial", 14, "bold")).pack(pady=10)
    label_usuario = tk.Label(janela, text=f"Usuário: {user_email}", font=("Arial", 10))
    label_usuario.pack(pady=5)
    tk.Button(janela, text="Selecionar arquivo .env", command=carregar_env_manual, font=("Arial", 11), bg="#2196F3", fg="white").pack(pady=5)
    tk.Button(janela, text="Selecionar arquivo de chassis e iniciar", command=iniciar_automacao, font=("Arial", 12), bg="#4CAF50", fg="white").pack(pady=20)
    tk.Label(janela, text="Certifique-se de preencher o .env com seu login.", font=("Arial", 9)).pack(side="bottom", pady=10)
    janela.mainloop()

def executar_automacao(user_email, user_password, lista_chassis):
    # Configurar o ChromeDriver automaticamente
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-notifications')

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    actions = ActionChains(driver)

    def fazer_login():
        driver.get('https://microworkcloud.com.br')
        email_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, 'Username')))
        email_field.send_keys(user_email)
        password_field = driver.find_element(By.ID, 'Password')
        password_field.send_keys(user_password)
        entrar_btn = driver.find_element(By.NAME, 'entrar')
        entrar_btn.click()
        time.sleep(10)

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
                return WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((by, value)))
            except:
                continue
        raise Exception("Campo de chassi não encontrado após tentar várias estratégias")

    def registrar_chassi():
        estrategias = [
            lambda: driver.find_element(By.TAG_NAME, 'body').click(),
            lambda: actions.send_keys(Keys.TAB).perform(),
            lambda: driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click(),
            lambda: driver.find_element(By.XPATH, '//button[contains(text(), "Confirmar")]').click()
        ]
        for estrategia in estrategias:
            try:
                estrategia()
                return True
            except:
                continue
        return False

    def inserir_chassi(chassi):
        try:
            campo = encontrar_campo_chassi()
            pyperclip.copy(chassi)
            campo.click()
            campo.clear()
            campo.send_keys(Keys.CONTROL, 'v')
            if not registrar_chassi():
                print("Aviso: Não foi possível confirmar o chassi automaticamente")
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

    try:
        print("Iniciando automação...")
        print(f"{len(lista_chassis)} chassis carregados")
        fazer_login()
        print("Login realizado")
        acessar_contagem_estoque()
        print("Acessou contagem de estoque (manual)")
        print(f"\nIniciando inserção de {len(lista_chassis)} chassis...")
        for i, chassi in enumerate(lista_chassis, 1):
            print(f"{i}/{len(lista_chassis)} - Inserindo: {chassi}")
            inserir_chassi(chassi)
        print("\nTodos os chassis foram inseridos. Aguardando 5 segundos antes de clicar no botão 'Salvar'...")
        time.sleep(5)
        clicar_botao_salvar()
        print("\nProcesso concluído com sucesso!")
        time.sleep(5)  # Aguarda 5 segundos antes de fechar o navegador
    except Exception as e:
        print(f"\nErro crítico: {str(e)}")
        driver.save_screenshot('erro_automacao.png')
        print("Screenshot salvo como 'erro_automacao.png'")
    finally:
        driver.quit()

if __name__ == "__main__":
    iniciar_interface()