from playwright.sync_api import sync_playwright


class emissao_guias:
    def __init__(self):
        self.p = sync_playwright().start()
        self.browser = self.p.chromium.launch(headless=False)
        self.page = self.browser.new_page()
        self.sitram_url = "http://www2.sefaz.ce.gov.br/sitram-internet/masterDetailLancamento.do?method=prepareSearch&quiosque=n"
    
    def open_site(self):
        self.page.goto(self.sitram_url)


# Add período fato gerador e inscrição estadual 
    # #dataInicialFatoGerador
    # #dataFinalFatoGerador


# Pesquisar
    # #pesquisar


# Ctrl + p para salvar o relatório em PDF
    def save_pdf(self):
        self.page.keyboard.down("Control")
        self.page.keyboard.press("p")
        self.page.keyboard.up("Control")

# Alterar o nome do relatório salvo para "ICMS ANTECIPADO E ST (n inscrição estadual) - REF mmYYYY.pdf"

# 