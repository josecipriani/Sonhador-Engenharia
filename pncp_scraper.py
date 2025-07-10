import os
import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service

try:
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
except ImportError:
    EdgeChromiumDriverManager = None

# Palavra a ser buscada
PALAVRA_CHAVE = "telha transparente"
# Caminho para o Edge WebDriver. Deixe como ``None`` para usar o caminho do
# sistema ou a instalação automática via ``webdriver-manager``.
CAMINHO_DRIVER = None
# Tempo máximo de espera para elementos
TEMPO_ESPERA = 15


def criar_driver():
    """Cria e retorna uma instância do Edge WebDriver."""
    options = Options()
    options.use_chromium = True
    options.add_argument("--start-maximized")

    caminho = os.getenv("EDGE_DRIVER_PATH", CAMINHO_DRIVER)

    if EdgeChromiumDriverManager is not None and caminho is None:
        service = Service(EdgeChromiumDriverManager().install())
    else:
        service = Service(caminho)

    return webdriver.Edge(service=service, options=options)


def esperar_e_clicar(driver, wait, by, valor):
    """Espera o elemento ficar clicável e clica nele."""
    botao = wait.until(EC.element_to_be_clickable((by, valor)))
    botao.click()
    time.sleep(1)


def main():
    driver = criar_driver()
    wait = WebDriverWait(driver, TEMPO_ESPERA)
    dados_resultado = []

    print("Acessando o site...")
    driver.get("https://pncp.gov.br/app/editais")
    time.sleep(6)

    input_busca = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//input[@placeholder='Buscar por palavras-chave']")
        )
    )
    input_busca.clear()
    input_busca.send_keys(PALAVRA_CHAVE)
    input_busca.send_keys(Keys.ENTER)
    time.sleep(4)

    esperar_e_clicar(driver, wait, By.XPATH, "//mat-checkbox//label[contains(., 'Encerradas')]")

    try:
        esperar_e_clicar(driver, wait, By.XPATH, "//mat-select[@formcontrolname='ordenarPor']")
        esperar_e_clicar(driver, wait, By.XPATH, "//mat-option//span[contains(text(), 'Mais relevante')]")
    except Exception:
        print("\u26a0\ufe0f N\u00e3o foi poss\u00edvel ordenar por 'Mais relevante'.")

    pagina = 1
    while True:
        print(f"\n\U0001F50D Coletando p\u00e1gina {pagina}...")
        time.sleep(3)
        editais = driver.find_elements(By.CSS_SELECTOR, "app-card-edital a")
        if not editais:
            print("❌ Nenhum edital encontrado nesta p\u00e1gina.")
            break

        links = [e.get_attribute("href") for e in editais if e.get_attribute("href")]
        print(f"🔗 {len(links)} editais encontrados.")

        for link in links:
            driver.execute_script("window.open(arguments[0]);", link)
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(4)
            try:
                esperar_e_clicar(driver, wait, By.XPATH, "//mat-tab-header//*[contains(text(), 'Itens')]")
                esperar_e_clicar(driver, wait, By.XPATH, "//mat-select[contains(@aria-label, 'itens por página')]")
                esperar_e_clicar(driver, wait, By.XPATH, "//mat-option//span[contains(text(), '50')]")
                time.sleep(2)

                encontrou = False
                textos_itens = driver.find_elements(By.CSS_SELECTOR, "table mat-cell")
                for t in textos_itens:
                    texto = t.text.lower()
                    if PALAVRA_CHAVE in texto:
                        dados_resultado.append({"edital": link, "item": t.text})
                        encontrou = True

                if encontrou:
                    print(f"✅ Edital relevante: {link}")

            except Exception as e:
                print(f"⚠️ Erro ao processar edital {link}: {e}")

            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        try:
            botao_proximo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Próxima página']")))
            if 'disabled' in botao_proximo.get_attribute("class"):
                print("📄 Fim das p\u00e1ginas.")
                break
            else:
                botao_proximo.click()
                pagina += 1
        except Exception as e:
            print(f"📄 Bot\u00e3o de pr\u00f3xima p\u00e1gina n\u00e3o encontrado ou erro: {e}")
            break

    if dados_resultado:
        df = pd.DataFrame(dados_resultado)
        df.to_excel("resultado_telhas.xlsx", index=False)
        print("✅ Resultado salvo em 'resultado_telhas.xlsx'")
    else:
        print("❌ Nenhum item com a palavra-chave foi encontrado.")

    driver.quit()


if __name__ == "__main__":
    main()
