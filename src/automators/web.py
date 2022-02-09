if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from utils import msg, retry, compiled


class Web:
    def __init__(self):
        self.driver = None
        self.totvs_logged = False
        self.opened = False
        self.vars = {}

    def open(self):
        msg("Abrindo o navegador")

        from webdriver_manager.chrome import ChromeDriverManager
        from selenium import webdriver

        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--allow-running-insecure-content')
        options.add_argument('--allow-insecure-localhost')
        options.add_argument('--unsafely-treat-insecure-origin-as-secure')

        if compiled():
            options.add_argument("--headless")

        self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
        self.driver.set_window_size(1280, 773)

        self.opened = True

    def close(self):
        msg("Fechando o navegador")

        self.driver.quit()

    def wait_for_window(self, timeout=2):
        msg("Aguardando nova janela")

        from time import sleep
        sleep(round(timeout / 1000))
        wh_now = self.driver.window_handles
        wh_then = self.vars["window_handles"]
        if len(wh_now) > len(wh_then):
            return set(wh_now).difference(set(wh_then)).pop()

    def wait_disappear(self, by, what):
        from selenium.webdriver.support import expected_conditions
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchWindowException

        try:
            WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((by, what)))
        except TimeoutException:
            return
        try:
            WebDriverWait(self.driver, 120).until(expected_conditions.staleness_of(self.driver.find_element(by, what)))
        except (NoSuchElementException, NoSuchWindowException):
            pass

    def totvs_access(self):
        msg("Acessando o TOTVS")

        self.driver.get("https://totvs.grendene.com.br/josso/signon/login.do")

    @retry
    def totvs_login(self, password):
        msg("Fazendo login")

        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.common.keys import Keys

        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "txtUsername")))
        self.driver.find_element(By.ID, "txtUsername").click()
        self.driver.find_element(By.ID, "txtUsername").send_keys("rep_trendy")
        self.driver.find_element(By.ID, "chkDomain").click()
        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "txtDomain")))
        self.driver.find_element(By.ID, "txtDomain").send_keys("gra_sid")
        self.driver.find_element(By.ID, "txtPassword").click()
        self.driver.find_element(By.ID, "txtPassword").send_keys(password)
        self.driver.find_element(By.ID, "txtPassword").send_keys(Keys.ENTER)

        self.wait_disappear(By.ID, "loading-screen")

        self.totvs_logged = True

    @retry
    def totvs_fav_pedidos(self):
        msg('Acessando a consulta de "Pedidos do Cliente - WEB"')

        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions

        WebDriverWait(self.driver, 90).until(
            expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-selector-light:nth-child(3)")))
        self.driver.find_element(By.CSS_SELECTOR, ".btn-selector-light:nth-child(3)").click()
        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".ng-scope:nth-child(18) > .col-lg-5")))
        self.driver.find_element(By.CSS_SELECTOR, ".ng-scope:nth-child(18) > .col-lg-5").click()
        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-primary")))
        self.vars["window_handles"] = self.driver.window_handles
        self.driver.find_element(By.CSS_SELECTOR, ".btn-primary").click()
        self.vars["win720"] = self.wait_for_window(5000)
        self.vars["root"] = self.driver.current_window_handle
        self.driver.switch_to.window(self.vars["win720"])
        # TODO wait for frame
        # WebDriverWait(self.driver, 30).until(expected_conditions.frame_to_be_available_and_switch_to_it(1))
        self.driver.switch_to.frame(1)

    @retry
    def totvs_fav_pedidos_fill(self, cod_cliente, prev_emb, implatacacao_ini):
        msg("Preenchendo os dados necessários")

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
        WebDriverWait(dropdown, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH, "//option[. = 'Todos']")))
        dropdown.find_element(By.XPATH, "//option[. = 'Todos']").click()
        self.driver.find_element(By.NAME, "I11").click()

        self.wait_disappear(By.ID, "janelaTudo")

    @retry
    def totvs_fav_pedidos_table(self):
        msg("Coletando uma tabela")

        from lxml.etree import HTML
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions

        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH, "/html/body/form/table[3]/tbody")))

        parsed_table = \
            HTML(self.driver.find_element_by_xpath("/html/body/form/table[3]/tbody").get_attribute('innerHTML'))[0]
        table = []
        for line in parsed_table:
            if len(line) != 17:
                raise Exception(f"totvs_fav_pedidos has not the expected cols number at all lines")
            table.append([col.text or col[0].text for col in line])

        return table

    @retry
    def totvs_fav_pedidos_next_page(self):
        msg("Checando se há outra página")

        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import NoSuchElementException

        try:
            element = self.driver.find_element(By.CSS_SELECTOR, "td:nth-child(2) > a:nth-child(1) > img")
        except NoSuchElementException:
            return False

        if element.is_enabled() and element.get_attribute(
                'src') == "https://totvs-webspeed.grendene.com.br/ems20web/wimages/ii-nex.gif":

            element.click()
            self.wait_disappear(By.ID, "janelaTudo")

            return True
        else:
            return False

    def totvs_fav_pedidos_complete_table(self):
        msg("Preparando para coletar todas as tabelas da consulta")

        table = self.totvs_fav_pedidos_table()
        while self.totvs_fav_pedidos_next_page():
            table += self.totvs_fav_pedidos_table()[1:]
        return table


if __name__ == '__main__':
    web = Web()
    web.open()
    web.totvs_access()
    web.totvs_login()
    web.totvs_fav_pedidos()
    web.totvs_fav_pedidos_fill("1000595", "03012022", "16022000")
    print(web.totvs_fav_pedidos_complete_table())
