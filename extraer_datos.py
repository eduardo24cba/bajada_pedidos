from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def extraer_datos(driver, lista_pedidos, opcion_tabla):
    ots = list(lista_pedidos.keys())
    pep = ""
    for item in ots:

        driver.implicitly_wait(2)
       
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[text()='%s']" % item))
            ).click()

        materiales = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, opcion_tabla))
        ).find_element(By.TAG_NAME, value="tbody").find_elements(By.TAG_NAME, value="tr")
        
        lista_materiales = []
        datos = []
        
        for material in materiales:
            columnas = material.find_elements(By.TAG_NAME, value="td")
            for td in columnas[:6]:
                lista_materiales.append(td.text)
        
        driver.implicitly_wait(2)

        #agregamos el pep
        pep = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "sectionContainer"))
        )[0].find_element(By.TAG_NAME, value="table").find_element(By.TAG_NAME, value="tbody"
        ).find_elements(By.TAG_NAME, value="tr")[-1].find_elements(By.TAG_NAME, value="td")[-1].text

        regresar = "visualizarOtVolver" if opcion_tabla == "otDm" else "distribuidorCancelar"
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, regresar))
            ).click()

         
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[text()='%s']" % lista_pedidos.get(item)[0]))
            ).click()
        
        for t in driver.find_elements(By.CLASS_NAME, value="sectionContainer")[:-1]:
            datos.append(t.text)

        
        driver.implicitly_wait(2)

        lista_pedidos[item].append(datos)
        lista_pedidos[item].append(lista_materiales)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "visualizadorVolver"))
            ).click()

        datos.append("Pep:\n"+pep)
