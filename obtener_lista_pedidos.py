from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def obtener_tabla(driver, opcion):
    index_ot = 0
    index_pedido = 0
    bandeja_pedidos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, opcion)))

    try:
        ths = bandeja_pedidos.find_element(By.TAG_NAME, value="table").find_element(
            By.TAG_NAME, value="thead").find_element(By.TAG_NAME, value="tr").find_elements(By.TAG_NAME, value="th")
    except:
        #no hay pedidos pendientes o despachos pendientes
        return {}
    
    for index in range(len(ths)):
        if ths[index].text.lower() == "ot":
            index_ot = index
        if ths[index].text.lower() == "pedido":
            index_pedido = index

    trs = bandeja_pedidos.find_element(By.TAG_NAME, value="table").find_element(By.TAG_NAME, value="tbody").find_elements(By.TAG_NAME, value="tr")

    lista_pedidos = {}

    for item in trs:
        lista_pedidos[item.find_elements(By.TAG_NAME, value="td")[index_ot].text] = [item.find_elements(By.TAG_NAME, value="td")[index_pedido].text]

    return lista_pedidos
