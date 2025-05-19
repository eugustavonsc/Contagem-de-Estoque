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

# Configurações
user_email = 'gustavo.nascimento@cajueiromotos.com.br'
user_password = 'Gustavo@2025'
arquivo_chassis = '19052025_01_19.05moto.txt'

def carregar_chassis(arquivo):
    if not os.path.exists(arquivo):
        raise FileNotFoundError(f"Arquivo '{arquivo}' não encontrado!")

    with open(arquivo, 'r', encoding='utf-8') as f:
        chassis = [line.strip() for line in f.readlines() if line.strip()]

    if not chassis:
        raise ValueError("Nenhum chassi válido encontrado no arquivo!")

    return chassis

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

    # Aguarda o carregamento completo do sistema após login
    time.sleep(10)

def acessar_contagem_estoque():
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'favorito-titulo')))

    time.sleep(3)  # aguarda elementos carregarem totalmente

    for element in driver.find_elements(By.CLASS_NAME, 'favorito-titulo'):
        if element.text.strip().lower() in ['contagem de estoque', 'estoque de motos']:
            element.click()
            time.sleep(5)
            return

    raise Exception("Botão 'Contagem de Estoque' ou 'Estoque de Motos' não encontrado")

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
    lista_chassis = carregar_chassis(arquivo_chassis)
    print(f"{len(lista_chassis)} chassis carregados")

    fazer_login()
    print("Login realizado")

    acessar_contagem_estoque()
    print("Acessou contagem de estoque")

    input("\nSelecione a contagem manualmente e pressione ENTER para continuar...")

    print(f"\nIniciando inserção de {len(lista_chassis)} chassis...")

    for i, chassi in enumerate(lista_chassis, 1):
        print(f"{i}/{len(lista_chassis)} - Inserindo: {chassi}")
        inserir_chassi(chassi)

    print("\nTodos os chassis foram inseridos. Tentando clicar no botão 'Salvar'...")
    clicar_botao_salvar()
    print("\nProcesso concluído com sucesso!")

except Exception as e:
    print(f"\nErro crítico: {str(e)}")
    driver.save_screenshot('erro_automacao.png')
    print("Screenshot salvo como 'erro_automacao.png'")

finally:
    driver.quit()