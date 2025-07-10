# -*- coding: utf-8 -*-
# Linha extra 1
# Linha extra 2
# Linha extra 3

"""Ferramenta de raspagem de dados do PNCP.

Este script utiliza Selenium para navegar pelo site do Portal Nacional
 de Contratações Públicas (PNCP) em busca de editais que contenham
 determinada palavra-chave. O código foi organizado em classes e funções
 para facilitar a manutenção e possibilitar futuras extensões.

Principais recursos:
- Configuração flexível via linha de comando ou variáveis de ambiente
- Suporte opcional ao `webdriver-manager` para baixar o EdgeDriver
- Registro de logs com níveis de detalhe ajustáveis
- Captura automática de screenshots em caso de erros
- Armazenamento dos itens encontrados em planilha Excel
- Funções utilitárias para aguardar elementos e realizar ações

O objetivo é servir como exemplo didático de automação web com Selenium
aplicada ao PNCP. Ajustes podem ser necessários conforme a evolução do
site ou requisitos específicos de cada usuário.
"""

from __future__ import annotations

# ---------------------------------------------
# Este script foi desenvolvido para demonstrar
# como utilizar o Selenium em Python para
# automatizar consultas no PNCP.
#
# Embora funcione em sua forma atual, ele pode
# ser adaptado para casos de uso mais complexos.
# Sinta-se livre para modificar conforme a
# necessidade de cada projeto.
# ---------------------------------------------

import argparse
import logging
import os
import time
from dataclasses import dataclass
from typing import Generator, List, Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import TimeoutException

try:  # Dependência opcional para baixar o driver automaticamente
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
except Exception:  # pragma: no cover - execução offline
    EdgeChromiumDriverManager = None

# ---------------------------------------------------------------------------
# Configurações e estruturas de dados
# ---------------------------------------------------------------------------
@dataclass
class Config:
    """Parâmetros de configuração para o scraper."""

    palavra_chave: str = "telha transparente"
    caminho_driver: Optional[str] = None
    tempo_espera: int = 15
    headless: bool = False
    max_paginas: int = 0  # 0 significa todas
    arquivo_saida: str = "resultado_telhas.xlsx"
    screenshot_dir: str = "screenshots"
    log_level: str = "INFO"


@dataclass
class EditalItem:
    """Representa um item de um edital localizado pelo scraper."""

    edital: str
    item: str
    titulo: Optional[str] = None
    orgao: Optional[str] = None


# ---------------------------------------------------------------------------
# Funções utilitárias
# ---------------------------------------------------------------------------
def with_retry(tries: int, delay: float = 1.0):
    """Decorador simples para repetir operações que possam falhar."""

    def decorator(func):
        def wrapper(*args, **kwargs):
            last_err = None
            for attempt in range(1, tries + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as err:  # pragma: no cover - comunicação externa
                    last_err = err
                    logging.debug("Tentativa %s/%s falhou: %s", attempt, tries, err)
                    time.sleep(delay)
            raise last_err

        return wrapper

    return decorator


# ---------------------------------------------------------------------------
# Classe principal de raspagem
# ---------------------------------------------------------------------------
class PNCPScraper:
    """Classe que encapsula toda a lógica de raspagem do PNCP."""

    BASE_URL = "https://pncp.gov.br/app/editais"

    def __init__(self, config: Config) -> None:
        self.config = config
        self.driver = self._criar_driver()
        self.wait = WebDriverWait(self.driver, self.config.tempo_espera)
        self.resultados: List[EditalItem] = []
        os.makedirs(self.config.screenshot_dir, exist_ok=True)
        logging.basicConfig(
            level=getattr(logging, self.config.log_level.upper(), logging.INFO),
            format="%(asctime)s [%(levelname)s] %(message)s",
        )
        logging.info("Scraper iniciado com palavra-chave '%s'", self.config.palavra_chave)

    # ------------------------------------------------------------------
    # Criação e finalização do WebDriver
    # ------------------------------------------------------------------
    def _criar_driver(self) -> webdriver.Edge:
        options = Options()
        options.use_chromium = True
        options.add_argument("--start-maximized")
        if self.config.headless:
            options.add_argument("--headless=new")

        caminho = (
            os.getenv("EDGE_DRIVER_PATH")
            or self.config.caminho_driver
            or CAMINHO_DRIVER
        )
        if caminho and os.path.exists(caminho):
            service = Service(caminho)
        elif EdgeChromiumDriverManager is not None:
            logging.info("Baixando driver com webdriver-manager")
            service = Service(EdgeChromiumDriverManager().install())
        else:
            raise RuntimeError(
                "Edge WebDriver não encontrado. Informe o caminho ou instale webdriver-manager."
            )
        return webdriver.Edge(service=service, options=options)

    def fechar(self) -> None:
        if self.driver:
            self.driver.quit()

    # ------------------------------------------------------------------
    # Métodos auxiliares de interação
    # ------------------------------------------------------------------
    def _scroll_to_bottom(self) -> None:
        last_height = self.driver.execute_script("return document.body.scrollHeight")
        while True:
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

    def _esperar_e_clicar(self, by: By, valor: str) -> None:
        logging.debug("Clicando em %s", valor)
        for _ in range(3):
            try:
                elemento = self.wait.until(EC.element_to_be_clickable((by, valor)))
                self.driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                elemento.click()
                time.sleep(1)
                return
            except Exception as exc:
                logging.debug("Falha ao clicar: %s", exc)
                time.sleep(1)
        raise RuntimeError(f"Não foi possível clicar no elemento {valor}")

    def _fechar_dialogos(self) -> None:
        seletores = ["button[aria-label='Fechar']", "button.cookie-consent-accept"]
        for seletor in seletores:
            for el in self.driver.find_elements(By.CSS_SELECTOR, seletor):
                try:
                    el.click()
                    time.sleep(0.5)
                except Exception:
                    pass

    def _capturar_screenshot(self, prefixo: str) -> None:
        nome = f"{prefixo}_{int(time.time())}.png"
        caminho = os.path.join(self.config.screenshot_dir, nome)
        try:
            self.driver.save_screenshot(caminho)
            logging.debug("Screenshot salvo em %s", caminho)
        except Exception:
            logging.error("Falha ao salvar screenshot")

    # ------------------------------------------------------------------
    # Navegação principal
    # ------------------------------------------------------------------
    @with_retry(3)
    def abrir_site(self) -> None:
        logging.info("Abrindo %s", self.BASE_URL)
        self.driver.get(self.BASE_URL)
        time.sleep(5)
        self._fechar_dialogos()

    @with_retry(3)
    def realizar_busca(self) -> None:
        campo = self._obter_campo_busca()
        if not campo:
            raise RuntimeError("Campo de busca não localizado")
        campo.clear()
        campo.send_keys(self.config.palavra_chave)
        campo.send_keys(Keys.ENTER)
        time.sleep(2)
        self._fechar_dialogos()
        # habilita editais encerrados
        self._esperar_e_clicar(By.XPATH, "//mat-checkbox//label[contains(., 'Encerradas')]")
        try:
            self._esperar_e_clicar(By.XPATH, "//mat-select[@formcontrolname='ordenarPor']")
            self._esperar_e_clicar(By.XPATH, "//mat-option//span[contains(text(), 'Mais relevante')]")
        except Exception:
            logging.warning("Não foi possível definir ordenação por relevância")

    def _obter_campo_busca(self):
        locators = [
            (By.XPATH, "//input[@placeholder='Buscar por palavras-chave']"),
            (By.CSS_SELECTOR, "input[placeholder*='palavra']"),
        ]
        for by, value in locators:
            try:
                return self.wait.until(EC.presence_of_element_located((by, value)))
            except TimeoutException:
                continue
        return None

    def _iterar_paginas(self) -> Generator[str, None, None]:
        pagina = 1
        while True:
            if self.config.max_paginas and pagina > self.config.max_paginas:
                break
            logging.info("Lendo página %s", pagina)
            time.sleep(2)
            editais = self.driver.find_elements(By.CSS_SELECTOR, "app-card-edital a")
            if not editais:
                logging.info("Nenhum edital encontrado na página %s", pagina)
                break
            for e in editais:
                href = e.get_attribute("href")
                if href:
                    yield href
            try:
                botao_proximo = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Próxima página']"))
                )
                if "disabled" in botao_proximo.get_attribute("class"):
                    break
                botao_proximo.click()
                pagina += 1
            except Exception:
                break

    def _extrair_informacoes_basicas(self) -> dict | None:
        try:
            titulo = self.driver.find_element(By.CSS_SELECTOR, "h1").text
            orgao = self.driver.find_element(By.CSS_SELECTOR, "app-entidade a").text
            return {"titulo": titulo, "orgao": orgao}
        except Exception:
            return None

    def _coletar_itens_edital(self, link: str) -> None:
        self.driver.execute_script("window.open(arguments[0]);", link)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        try:
            self._fechar_dialogos()
            self._esperar_e_clicar(By.XPATH, "//mat-tab-header//*[contains(text(), 'Itens')]")
            self._esperar_e_clicar(By.XPATH, "//mat-select[contains(@aria-label, 'itens por página')]")
            self._esperar_e_clicar(By.XPATH, "//mat-option//span[contains(text(), '50')]")
            time.sleep(1)
            detalhes = self._extrair_informacoes_basicas()
            textos = self.driver.find_elements(By.CSS_SELECTOR, "table mat-cell")
            encontrou = False
            for t in textos:
                if self.config.palavra_chave.lower() in t.text.lower():
                    self.resultados.append(
                        EditalItem(
                            edital=link,
                            item=t.text,
                            titulo=detalhes.get("titulo") if detalhes else None,
                            orgao=detalhes.get("orgao") if detalhes else None,
                        )
                    )
                    encontrou = True
            if encontrou:
                logging.info("Edital relevante: %s", link)
        except Exception as err:
            logging.error("Erro ao processar edital %s: %s", link, err)
            self._capturar_screenshot("erro_edital")
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    @with_retry(2)
    def coletar_resultados(self) -> None:
        for link in self._iterar_paginas():
            self._coletar_itens_edital(link)
            self._scroll_to_bottom()

    def salvar(self) -> None:
        if not self.resultados:
            logging.info("Nenhum item encontrado para salvar")
            return
        df = pd.DataFrame([r.__dict__ for r in self.resultados])
        df.to_excel(self.config.arquivo_saida, index=False)
        logging.info("Dados salvos em %s", self.config.arquivo_saida)


# ---------------------------------------------------------------------------
# Funções de linha de comando
# ---------------------------------------------------------------------------
def parse_args() -> Config:
    parser = argparse.ArgumentParser(description="Scraper de editais do PNCP")
    parser.add_argument("--keyword", dest="palavra_chave", default="telha transparente", help="Palavra-chave de busca")
    parser.add_argument("--driver", dest="caminho_driver", default=None, help="Caminho para o EdgeDriver")
    parser.add_argument("--headless", action="store_true", help="Executar navegador em modo headless")
    parser.add_argument("--max-pages", dest="max_paginas", type=int, default=0, help="Limitar quantidade de páginas a visitar")
    parser.add_argument("--output", dest="arquivo_saida", default="resultado_telhas.xlsx", help="Arquivo Excel de saída")
    parser.add_argument("--wait", dest="tempo_espera", type=int, default=15, help="Tempo de espera para elementos")
    parser.add_argument("--log", dest="log_level", default="INFO", help="Nível de log (DEBUG, INFO, WARNING)")
    return Config(**vars(parser.parse_args()))


# Utilitário extra
log_environment = lambda: (logging.info("Selenium %s", webdriver.__version__), logging.info("HEADLESS=%s", os.getenv("HEADLESS")), logging.info("EDGE_DRIVER_PATH=%s", os.getenv("EDGE_DRIVER_PATH") or "indefinido"))


def main() -> None:
    cfg = parse_args()
    log_environment()
    scraper = PNCPScraper(cfg)
    try:
        scraper.abrir_site()
        scraper.realizar_busca()
        scraper.coletar_resultados()
        scraper.salvar()
    finally:
        scraper.fechar()


if __name__ == "__main__":
    main()
# Fim do script
# Desenvolvido como exemplo de automação PNCP
# Consulte a documentação oficial do Selenium
# para maiores informações sobre APIs disponíveis.
# Bom uso!

# Este é o final real do arquivo

