from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

LOGISTICA_ALMACENES = "http://bpmred/portal/servlet/controller?viewName=LogisticaAlmacenes%2Fapplications&from=1"
SMES_DEVOLUCION = "http://bpmred/portal/servlet/controller?viewName=SMEGestionDeSolicitud%2FAplicaciones&from=1"

def login(logistica):
    #nos logueamos a la pagina que necesitamos y retornamos el driver para usar.
    
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
    
    driver.get(LOGISTICA_ALMACENES if logistica else SMES_DEVOLUCION)

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username"))).send_keys()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "password"))).send_keys()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "submitbutton"))).click()

    return driver


def logistica_almacenes(driver):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='%s']" % "Logística de almacenes"))).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='%s']" % "Aplicaciones"))).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='%s']" % "Gestión de almacenes"))).click()


def smes(driver):
    smes = driver.find_element(by='xpath', value="//*[text()='Gestionar SME de Desinstalación y devolución de Materiales']")
    smes.click()
    mostrar = Select(driver.find_element(by='xpath', value="/html/body/div[2]/div/form/div[3]/div/div/div[3]/select"))
    mostrar.select_by_value("100")
    ultima_pagina = driver.find_element(by='xpath', value="//*[text()='9']")
    ultima_pagina.click()
