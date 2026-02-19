from playwright.sync_api import sync_playwright
import pandas as pd
from pathlib import Path
import time

class emissao_guias:
    def __init__(self):
        self.p = sync_playwright().start()
        self.browser = self.p.chromium.launch(headless=False)
        self.page = self.browser.new_page()
        self.sitram_url = "http://www2.sefaz.ce.gov.br/sitram-internet/masterDetailLancamento.do?method=prepareSearch&quiosque=n"
        url_file = Path(__file__).parent / "SITRAM.html" #path to file to don't open page all the time
        self.page.goto(str(url_file)) # change it to sitram_url afterwards

# Add período fato gerador e inscrição estadual 
    
    # Período fato gerador
    # #dataInicialFatoGerador
    # #dataFinalFatoGerador
    def add_date(self):
        self.page.locator("#dataInicialFatoGerador").fill("01/01/2024") # o que vai mudar é o mês e o ano
        self.page.locator("#dataFinalFatoGerador").fill("31/01/2024") # idem
    
    # I.E.
    # #contribuinteCredenciado
    def add_ie(self):
        excel_path = Path(__file__).parent / "Ceara_filiais.xlsx"
        df = pd.read_excel(excel_path)
        ie_number = df['I..E.']
        self.page.locator("#contribuinteCredenciado").fill(ie_number[0]) # o que vai mudar é o número da I.E. de cada filial, então tem que ser um loop para preencher cada uma delas
        time.sleep(3) # pode remover isso aqui


# Pesquisar
    # #pesquisar
    

# Salvar o relatório em PDF - considerando que esse arquivo ta dentro de Apuração ICMS > Ceará
    # Nome do PDF vai ser o mesmo mês e ano que for colocado no período fato gerador
    def save_pdf(self):
        base_path = Path(__file__).parent
        pdf_path = base_path / "save-here" / "file2.pdf"
        self.page.pdf(path=str(pdf_path))

# Alterar o nome do relatório salvo para "ICMS ANTECIPADO E ST (n inscrição estadual) - REF mmYYYY.pdf"
    
# 

if __name__ == "__main__":
    p = emissao_guias()
    p.add_date()
    p.add_ie()