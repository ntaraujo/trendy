if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from utils import msg, retry, compiled
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchWindowException
from time import sleep
from lxml.etree import HTML



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

        if compiled:
            options.add_argument("--headless")

        self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
        self.driver.set_window_size(1280, 773)

        self.opened = True

    def close(self):
        msg("Fechando o navegador")

        self.driver.quit()
    
    def print(self, basename):
        from Screenshot import Screenshot_Clipping
        from utils import data_dir_path
        from os import path

        ss = Screenshot_Clipping.Screenshot()
        ss.full_Screenshot(self.driver, save_path=data_dir_path , image_name=basename)
        return path.join(data_dir_path, basename)
    
    def prepare_for_new_window(self):
        self.vars["window_handles"] = self.driver.window_handles
        return self.driver.current_window_handle

    def get_new_window(self):
        msg("Aguardando nova janela")

        wh_then = self.vars["window_handles"]
        for _ in range(10):
            sleep(1)
            wh_now = self.driver.window_handles
            if len(wh_now) > len(wh_then):
                return set(wh_now).difference(set(wh_then)).pop()

    def wait_disappear(self, by, what):
        try:
            WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((by, what)))
        except TimeoutException:
            return
        try:
            WebDriverWait(self.driver, 30).until(expected_conditions.staleness_of(self.driver.find_element(by, what)))
        except (NoSuchElementException, NoSuchWindowException):
            pass

    def totvs_access(self):
        msg("Acessando o TOTVS")

        self.driver.get("https://totvs.grendene.com.br/josso/signon/login.do")

    @retry
    def totvs_login(self, password):
        msg("Fazendo login")

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
    def totvs_fav_program_access(self, type_col, program_line, in_title=None):
        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, f".btn-selector-light:nth-child({type_col})")))
        self.driver.find_element(By.CSS_SELECTOR, f".btn-selector-light:nth-child({type_col})").click()
        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, f".ng-scope:nth-child({program_line}) > .col-lg-5")))
        program = self.driver.find_element(By.CSS_SELECTOR, f".ng-scope:nth-child({program_line}) > .col-lg-5")
        if in_title:
            assert in_title in program.text
        msg(f'Acessando o favorito "{program.text}"')
        program.click()
        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-primary")))
        self.prepare_for_new_window()
        self.driver.find_element(By.CSS_SELECTOR, ".btn-primary").click()
        new_window = self.get_new_window()
        self.driver.switch_to.window(new_window)
    
    @retry
    def totvs_fav_clientes_va_para(self, cod_emitente):
        
        WebDriverWait(self.driver, 30).until(expected_conditions.presence_of_element_located((By.NAME, "Fr_panel")))
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "Fr_panel"))
        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr/td[5]/a")))
        main_window = self.prepare_for_new_window()
        self.driver.find_element(By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr/td[5]/a").click()

        new_window = self.get_new_window()
        self.driver.switch_to.window(new_window)
        self.driver.find_element(By.NAME, "w_cod_emitente").send_keys(cod_emitente)
        self.driver.find_element(By.NAME, "w_cod_emitente").send_keys(Keys.ENTER)
        
        self.driver.switch_to.window(main_window)
    
    @retry
    def totvs_fav_clientes_documentos(self, cod_emitente):
        WebDriverWait(self.driver, 30).until(expected_conditions.presence_of_element_located((By.NAME, "Fr_work")))
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "Fr_work"))
        WebDriverWait(self.driver, 30).until(expected_conditions.presence_of_element_located((By.XPATH, f'/html/body/form/div[1]/center/table/tbody/tr[2]/td/div[2]/center/table/tbody/tr[1]/td/input[@value="{cod_emitente}"]')))
        main_window = self.prepare_for_new_window()
        self.driver.find_element(By.XPATH, "/html/body/form/div[1]/center/table/tbody/tr[2]/td/div[1]/center/table/tbody/tr/th[4]/a").click()
        new_window = self.get_new_window()

        self.driver.switch_to.window(new_window)

        return main_window
    
    @retry
    def totvs_fav_clientes_filtro(self):
        WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.NAME, "bt_param")))
        main_window = self.prepare_for_new_window()
        self.driver.find_element(By.NAME, "bt_param").click()
        new_window = self.get_new_window()

        self.driver.switch_to.window(new_window)
        WebDriverWait(self.driver, 30).until(expected_conditions.presence_of_element_located((By.NAME, "w_dt_venc_ini")))
        el = self.driver.find_element(By.NAME, "w_dt_venc_ini")
        el.clear()
        el.send_keys("01/01/1996")
        el = self.driver.find_element(By.NAME, "w_dt_venc_fin")
        el.clear()
        el.send_keys("31/12/9999")
        self.driver.find_element(By.NAME, "button1").click()
        self.driver.switch_to.window(main_window)

    @retry
    def totvs_fav_pedidos_fill(self, cod_cliente, prev_emb, implatacacao_ini):
        msg("Preenchendo os dados necessários")

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
    def totvs_table(self, tbody_xpath, expected_cols):
        msg("Coletando uma tabela")

        WebDriverWait(self.driver, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH, tbody_xpath)))

        parsed_table = \
            HTML(self.driver.find_element_by_xpath(tbody_xpath).get_attribute('innerHTML'))[0]
        table = []
        for line in parsed_table:
            if len(line) != expected_cols:
                raise Exception(f"Table has not the expected cols number at all lines")
            table.append([self.totvs_table_helper(col) for col in line])

        return table
    
    @staticmethod
    def totvs_table_helper(col):
        r = (col.text or '').strip()
        while not r:
            if len(col) < 1:
                break
            col = col[0]
            r = (col.text or '').strip()
        return r

    @retry
    def totvs_next_page(self, img_css_selector):
        msg("Checando se há outra página")

        try:
            element = self.driver.find_element(By.CSS_SELECTOR, img_css_selector)
        except NoSuchElementException:
            return False

        if element.is_enabled() and element.get_attribute(
                'src') == "https://totvs-webspeed.grendene.com.br/ems20web/wimages/ii-nex.gif":

            element.click()
            self.wait_disappear(By.ID, "janelaTudo")

            return True
        else:
            return False
    
    def totvs_complete_table(self, tbody_xpath, expected_cols, img_css_selector):
        msg("Preparando para coletar todas as tabelas da consulta")

        table = self.totvs_table(tbody_xpath, expected_cols)
        while self.totvs_next_page(img_css_selector):
            partial_table = self.totvs_table(tbody_xpath, expected_cols)
            if len(partial_table) > 1:
                table += partial_table[1:]
        return table

    def totvs_fav_pedidos_complete_table(self):
        return self.totvs_complete_table("/html/body/form/table[3]/tbody", 17, "td:nth-child(2) > a:nth-child(1) > img")
    
    def totvs_fav_clientes_complete_table(self):
        return self.totvs_complete_table("/html/body/form/table[1]/tbody", 19, "body > form > table:nth-child(5) > tbody > tr > td > a > img")


if __name__ == '__main__':
    web = Web()
    web.open()
    web.totvs_access()
    web.totvs_login()
    web.totvs_fav_pedidos()
    web.totvs_fav_pedidos_fill("1000595", "03012022", "16022000")
    print(web.totvs_fav_pedidos_complete_table())
