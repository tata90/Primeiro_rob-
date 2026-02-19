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
        self.page.goto(self.sitram_url) 

# Add período fato gerador e inscrição estadual 
    
    # Período fato gerador
    # #dataInicialFatoGerador
    # #dataFinalFatoGerador
    def add_date(self):
        self.page.locator("#dataInicialFatoGerador").clear()
        self.page.locator("#dataInicialFatoGerador").type("01/01/2024", delay=100) # quem vai colocar a data é o usuário
        self.page.wait_for_timeout(1000)
        self.page.locator("#dataFinalFatoGerador").type("31/01/2024", delay=100) # quem vai colocar a data é o usuário

# Pesquisar
    # #pesquisar
    def search(self):
        self.page.locator("#pesquisar").click()
    
# I.E.
    # #contribuinteCredenciado
    def add_ie(self):
        excel_path = Path(__file__).parent / "Ceara_filiais.xlsx"
        df = pd.read_excel(excel_path)

        for ie_number in df["I..E."]:
            
            self.add_date()
            self.page.wait_for_timeout(2000)

            self.page.locator("#contribuinteCredenciado").fill(ie_number)
            self.page.wait_for_timeout(1000)

            self.search()
            self.page.wait_for_timeout(1000)
            
            try:
                self.page.locator("button:has-text('OK')").wait_for(state="visible", timeout=3000)
                print(f"Error processing I.E. {ie_number}")
                self.page.locator("button:has-text('OK')").click()
                self.page.wait_for_timeout(1000)
            except:
                self.save_pdf(ie_number, "01", "2024") #de novo, mes e ano vem do usuário - resolve isso
                pass
            

# Salvar o relatório em PDF - considerando que esse arquivo ta dentro de Apuração ICMS > Ceará
    # Nome do PDF vai ser o mesmo mês e ano que for colocado no período fato gerador
    def save_pdf(self, ie_number, month, year):
        base_path = Path(__file__).parent
        pdf_name = f"ICMS ANTECIPADO E ST ({ie_number}) - REF {month}{year}.pdf"
        pdf_path = base_path / "save-here" / pdf_name
        self.page.pdf(path=str(pdf_path))

# Alterar o nome do relatório salvo para "ICMS ANTECIPADO E ST (n inscrição estadual) - REF mmYYYY.pdf"
    
# Fechar o browser    
    def close(self):
        self.browser.close()

if __name__ == "__main__":
    p = emissao_guias()
    p.add_ie()

    p.close()