from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl import Workbook

from login import login, logistica_almacenes

from obtener_lista_pedidos import obtener_tabla

from extraer_datos import extraer_datos

from crear_excel import crear_excel

import os

import time

from google_drive import login_drive, upload_file
    
titulos = ["Despacho", "OT", "Pedido", "Solicitante", "Destino", "Provincia", "Urgente", "Centro sap",
           "Almacén sap", "Consolidar", "Almacén consolidación", "Fecha necesidad", "Retira de almacén"]

columnas = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N", "O", "P", "Q", "R", "S"]

titulo_ot = ["Código Material", "Código SAP", "Descripción Material", "Es Solución", "Componentes", "Cantidad"]

####Login####
driver = login(logistica=True)

#logistica almacenes
logistica_almacenes(driver)

#obtenemos la lista de los pedidos
#extraemos los datos de cada pedido y lo agregamos a la lista que ya teniamos, en este caso lista_pedidos
pedidos_pendientes = {}
despachos_preparados = {}

pedidos_pendientes = obtener_tabla(driver, "OTPendientes")

extraer_datos(driver, pedidos_pendientes, "distribuidor")

despachos_preparados = obtener_tabla(driver, "DespachosPreparados")

extraer_datos(driver, despachos_preparados, "otDm")

#unimos los dos diccionarios en uno y creamos el excel
lista_pedidos = {**pedidos_pendientes, **despachos_preparados}

if lista_pedidos:
    file_name = crear_excel(lista_pedidos)
    service = login_drive()
    upload_file(file_name, service)
    
print("terminado")













    
