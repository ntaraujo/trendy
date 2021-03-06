# Generated by Selenium IDE
import pytest
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

class TestConsultapedidos3():
  def setup_method(self, method):
    self.driver = webdriver.Chrome()
    self.vars = {}
  
  def teardown_method(self, method):
    self.driver.quit()
  
  def wait_for_window(self, timeout = 2):
    time.sleep(round(timeout / 1000))
    wh_now = self.driver.window_handles
    wh_then = self.vars["window_handles"]
    if len(wh_now) > len(wh_then):
      return set(wh_now).difference(set(wh_then)).pop()
  
  def test_consultapedidos3(self):
    self.driver.get("https://totvs.grendene.com.br/josso/signon/login.do")
    self.driver.set_window_size(1280, 773)
    self.driver.find_element(By.ID, "txtPassword").send_keys("SENHA_TOTVS")
    self.driver.find_element(By.ID, "boxDomain").click()
    self.driver.find_element(By.ID, "chkDomain").click()
    self.driver.find_element(By.ID, "txtDomain").send_keys("gra_sid")
    self.driver.find_element(By.ID, "txtDomain").send_keys(Keys.ENTER)
    WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-selector-light:nth-child(3)")))
    self.driver.find_element(By.CSS_SELECTOR, ".btn-selector-light:nth-child(3)").click()
    element = self.driver.find_element(By.CSS_SELECTOR, ".table-responsive-light")
    actions = ActionChains(self.driver)
    actions.move_to_element(element).click_and_hold().perform()
    element = self.driver.find_element(By.CSS_SELECTOR, ".table-responsive-light")
    actions = ActionChains(self.driver)
    actions.move_to_element(element).perform()
    element = self.driver.find_element(By.CSS_SELECTOR, ".table-responsive-light")
    actions = ActionChains(self.driver)
    actions.move_to_element(element).release().perform()
    WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".ng-scope:nth-child(18) > .col-lg-5")))
    self.driver.find_element(By.CSS_SELECTOR, ".ng-scope:nth-child(18) > .col-lg-5").click()
    WebDriverWait(self.driver, 30).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".btn-primary")))
    self.vars["window_handles"] = self.driver.window_handles
    self.driver.find_element(By.CSS_SELECTOR, ".btn-primary").click()
    self.vars["win720"] = self.wait_for_window(2000)
    self.vars["root"] = self.driver.current_window_handle
    self.driver.switch_to.window(self.vars["win720"])
    self.driver.close()
    self.driver.switch_to.window(self.vars["root"])
    self.driver.close()
  
