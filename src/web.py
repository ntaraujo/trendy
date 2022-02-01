from utils import msg

class Web():
    def __init__(self):
        self.totvs_logged = False
        self.vars = {}

    def open(self):
        msg("Abrindo o navegador")

        from webdriver_manager.chrome import ChromeDriverManager
        from selenium import webdriver

        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.driver.set_window_size(1280, 773)
    
    def close(self):
        msg("Fechando o navegador")

        self.driver.quit()
  
    def wait_for_window(self, timeout = 2):
        from time import sleep
        sleep(round(timeout / 1000))
        wh_now = self.driver.window_handles
        wh_then = self.vars["window_handles"]
        if len(wh_now) > len(wh_then):
            return set(wh_now).difference(set(wh_then)).pop()
    
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
    
    def consulta_pedidos(self):
        msg('Acessando a Consulta de "Pedidos do Cliente - WEB"')

        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions

        WebDriverWait(self.driver, 60).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-selector-light:nth-child(3)")))
        self.driver.find_element(By.CSS_SELECTOR, ".btn-selector-light:nth-child(3)").click()
        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".ng-scope:nth-child(18) > .col-lg-5")))
        self.driver.find_element(By.CSS_SELECTOR, ".ng-scope:nth-child(18) > .col-lg-5").click()
        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-primary")))
        self.vars["window_handles"] = self.driver.window_handles
        self.driver.find_element(By.CSS_SELECTOR, ".btn-primary").click()
        self.vars["win720"] = self.wait_for_window(5000)
        self.vars["root"] = self.driver.current_window_handle
        self.driver.switch_to.window(self.vars["win720"])
        self.driver.switch_to.frame(1)
    
    def pedidos_clientes_centralizador(self, cod_cliente, prev_emb, implatacacao_ini="16022000"):
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions

        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.NAME, "w_cod_cliente")))
        self.driver.find_element(By.NAME, "w_cod_cliente").clear()
        self.driver.find_element(By.NAME, "w_cod_cliente").send_keys(cod_cliente)
        self.driver.find_element(By.NAME, "w_cod_rep").clear()
        self.driver.find_element(By.NAME, "w_cod_rep").send_keys("74800")
        self.driver.find_element(By.NAME, "w_dt_implantacao_ini").clear()
        self.driver.find_element(By.NAME, "w_dt_implantacao_ini").send_keys(implatacacao_ini)
        self.driver.find_element(By.NAME, "w_prev_emb").clear()
        self.driver.find_element(By.NAME, "w_prev_emb").send_keys(prev_emb)
        self.driver.find_element(By.NAME, "w_status").click()
        dropdown = self.driver.find_element(By.NAME, "w_status")
        dropdown.find_element(By.XPATH, "//option[. = 'Todos']").click()
        self.driver.find_element(By.NAME, "I11").click()

if __name__ == '__main__':
    web = Web()
    web.open()
    web.access_totvs()
    web.login_totvs()
    web.consulta_pedidos()
    web.pedidos_clientes_centralizador("1031462", "03032021")
