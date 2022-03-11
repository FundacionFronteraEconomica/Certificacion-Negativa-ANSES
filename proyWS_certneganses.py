from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
import time
import pandas as pd
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument("--headless")

driver = webdriver.Chrome(executable_path=r'chromedriver.exe', options=options)

driver.get("https://servicioswww.anses.gob.ar/censite/index.aspx")
time.sleep(3)

archivo = "sacar certificaciones.xlsx"
wb = load_workbook(archivo)
hoja = wb.get_sheet_names()
nombre = wb.get_sheet_by_name('Hoja1')
wb.close()

for i in range(1598,1599):
    pre, doc, fin = nombre[f'F{i}:H{i}'][0]
    time.sleep(1)
    driver.find_element_by_id("txtCuitPre").send_keys(str(pre.value))
    time.sleep(1)
    driver.find_element_by_id("txtCuitDoc").send_keys(str(doc.value))
    time.sleep(1)
    driver.find_element_by_id("txtCuitDV").send_keys(str(fin.value))
    time.sleep(1)
    driver.find_element_by_id("btnVerificar").click()
    time.sleep(1)
    t1 = driver.find_element_by_id('lblNombre')
    nom = t1.text
    t2 = driver.find_element_by_id('lblCuil')
    cuil = t2.text
    t3 = driver.find_element_by_xpath("//td[@class='fa fa-check fa-5x']")
    cert = t3.text
    t4 = driver.find_element_by_id("ANTECEDENTES")
    obs = t4.text
    d1 = [[nom,cuil,cert,obs]]
    for d in d1:
        print(d1)
    time.sleep(1)
    driver.find_element_by_xpath('//input[@value="SELECCIONAR OTRO PERIODO"]').click()

for i in range(1620,1783):
    pre, doc, fin = nombre[f'F{i}:H{i}'][0]
    time.sleep(1)
    driver.find_element_by_id("txtCuitPre").send_keys(str(pre.value))
    time.sleep(1)
    driver.find_element_by_id("txtCuitDoc").send_keys(str(doc.value))
    time.sleep(1)
    driver.find_element_by_id("txtCuitDV").send_keys(str(fin.value))
    time.sleep(1)
    driver.find_element_by_id("btnVerificar").click()
    time.sleep(1)
    t1 = driver.find_element_by_id('lblNombre')
    nom = t1.text
    t2 = driver.find_element_by_id('lblCuil')
    cuil = t2.text
    t3 = driver.find_element_by_xpath("//td[@class='fa fa-check fa-5x']")
    cert = t3.text
    t4 = driver.find_element_by_id("ANTECEDENTES")
    obs = t4.text
    d2 = [[nom,cuil,cert,obs]]
    for d in d2:
        d1.extend(d2)
    time.sleep(1)
    driver.find_element_by_xpath('//input[@value="SELECCIONAR OTRO PERIODO"]').click()

d1

df = pd.DataFrame(d1)
df

df.to_excel("certificaciones2.xlsx")

