from utils import msg

class Web():
    def __init__(self):
        self.totvs_logged = False

    def open(self):
        msg("Abrindo o navegador")

        from webdriver_manager.chrome import ChromeDriverManager
        from selenium import webdriver

        self.driver = webdriver.Chrome(ChromeDriverManager().install())
    
    def close(self):
        msg("Fechando o navegador")

        self.driver.quit()
    
    def access_totvs(self):
        msg("Acessando TOTVS")

        self.driver.get("https://totvs.grendene.com.br/josso/signon/login.do")
    
    def login_totvs(self, username="rep_trendy", password="Comunidade15", domain="gra_sid"):
        msg("Fazendo login TOTVS")
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.common.keys import Keys

        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "txtUsername")))

        self.driver.find_element(By.ID, "txtUsername").click()

        self.driver.find_element(By.ID, "txtUsername").send_keys(username)

        self.driver.find_element(By.ID, "chkDomain").click()

        self.driver.find_element(By.ID, "txtDomain").send_keys(domain)

        self.driver.find_element(By.ID, "txtPassword").click()

        self.driver.find_element(By.ID, "txtPassword").send_keys(password)

        self.driver.find_element(By.ID, "txtPassword").send_keys(Keys.ENTER)

        self.totvs_logged = True

if __name__ == '__main__':
    web = Web()
    web.open()
    web.access_totvs()
    web.login_totvs()
