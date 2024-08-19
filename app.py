from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from time import sleep

navegador = webdriver.Chrome()

tabela_produtos = pd.read_excel("buscas.xlsx")

navegador.get("https://www.google.com/")
produto = "iphone 12 64gb"

navegador.find_element(By.NAME,'q').send_keys(produto, Keys.ENTER)

sleep(4)
elementos = navegador.find_elements(By.CLASS_NAME, 'YmvwI')
for item in elementos:
    if 'Shopping' in item.text:
        item.click()
        break

sleep(10)