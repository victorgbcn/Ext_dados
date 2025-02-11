import time
import pyautogui
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import smtplib
from email.message import EmailMessage
import os

# Função para tentativa de acesso
def tentar_acessar_site(url, tentativas=3):
    for tentativa in range(tentativas):
        try:
            driver = webdriver.Chrome()
            driver.get(url)
            driver.maximize_window()
            return driver
        except WebDriverException as e:
            print(f"Erro ao acessar o site: {e}. Tentativa {tentativa + 1} de {tentativas}.")
            time.sleep(3)

    # Caso houvesse um banco, basta gerar um update com o log
    print("Site fora do ar")
    return None


# Função para ler o arquivo txt com a senha do email
def ler_senha_arquivo(caminho_arquivo):
    try:
        with open(caminho_arquivo, "r") as file:
            senha = file.read()  # Lê a senha
        return senha
    except FileNotFoundError:
        print(f"Erro: O arquivo {caminho_arquivo} não foi encontrado.")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None


# Função para extrair os dados das cinco primeiras paginas do site
def extrair_dados(driver):
    produtos = []
    # while True: # Para extrair os dados de todas as paginas do site
    for pagina in range(1, 6):  # Navegar pelas cinco primeiras páginas
        try:
            time.sleep(3)
            items = driver.find_elements(By.CSS_SELECTOR, '[class="sc-knefzF gmRWWc"]')
            for item in items:
                try:
                    # Nome
                    nome = item.find_element(By.CSS_SELECTOR, '[data-testid="product-title"]').text
                    print(nome)

                    # Avaliações
                    avaliacoes = item.find_element(By.CSS_SELECTOR, '.sc-cXPBUD.SkEmc').text
                    avaliacoes = int(avaliacoes.split('(')[-1].split(')')[0])
                    print(avaliacoes)

                    # URL
                    link = item.find_element(By.CSS_SELECTOR, 'a').get_attribute("href")
                    print(link)

                    # Adiciona os dados na lista de produtos
                    produtos.append({"PRODUTO": nome, "QTD_AVAL": avaliacoes, "URL": link})
                    print(produtos)
                except Exception:
                    continue

            # Ir para a próxima página
            botao_proximo = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '.fDVVLS')))
            botao_proximo.click()
        except Exception as e:
            print(f"Erro na página {pagina}: {e}")
            break

    return produtos


# Função para salvar os dados no Excel
def criar_excel(dados):
    #Criar tabela de dados
    df = pd.DataFrame(dados)

    # Filtra para excluir os produtos com 0 avaliações
    df = df[df["QTD_AVAL"] > 0]

    # Ordena os dados pela coluna 'QTD_AVAL' de forma decrescente
    df = df.sort_values(by="QTD_AVAL", ascending=False)

    # Filtra os produtos com menos de 100 avaliações como piores e com 100 ou mais avaliações como melhores
    piores = df[df["QTD_AVAL"] < 100]
    melhores = df[df["QTD_AVAL"] >= 100]

    # Cria a pasta com a planilha
    pasta_planilha = "Teste Pratico planilha"
    os.makedirs(pasta_planilha, exist_ok=True)

    # Cria o caminho completo para o arquivo Excel
    caminho_planilha = os.path.join(pasta_planilha, "Notebooks.xlsx")

    # Salva os dados nas abas
    with pd.ExcelWriter(caminho_planilha) as writer:
        piores.to_excel(writer, sheet_name="Piores", index=False)
        melhores.to_excel(writer, sheet_name="Melhores", index=False)

    return caminho_planilha


# Função para enviar o e-mail com o anexo
def enviar_email(caminho_anexo, senha_email):
    # Função para criar mensagem de e-mail
    msg = EmailMessage()
    msg["Subject"] = "Relatório Notebooks"
    msg["From"] = "victorgc.nicolau@gmail.com"
    msg["To"] = "victorgc.nicolau@gmail.com"

    # Corpo do e-mail
    msg.set_content("""
    Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.

    Atenciosamente,
    Robô
    """)

    # Adiciona anexo no email
    with open(caminho_anexo, "rb") as file:
        msg.add_attachment(file.read(), maintype="application", subtype="xlsx", filename="Notebooks.xlsx")

    # Cria uma conexão com o servidor SMTP do Gmail
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login("victorgc.nicolau@gmail.com", senha_email)  # Senha lida do arquivo
        smtp.send_message(msg)

# Caminho para o arquivo de senha
caminho_arquivo_senha = "senha.txt"

# Ler a senha do arquivo
senha_email = ler_senha_arquivo(caminho_arquivo_senha)

if senha_email is None:
    print("Erro: A senha não foi carregada corretamente.")
    exit()

driver = tentar_acessar_site('https://www.magazineluiza.com.br/')
if not driver:
    exit()

try:
    driver.find_element(By.XPATH,'//*[@id="input-search"]').click()
    time.sleep(2)
    pyautogui.write('notebooks')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(3)

    produtos = extrair_dados(driver)
    gerar_relatorio = criar_excel(produtos)
    enviar_email(gerar_relatorio, senha_email)

    print("Automação concluída com sucesso!")
finally:
    driver.quit()