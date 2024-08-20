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

lista_resultados = navegador.find_elements(By.CLASS_NAME, 'KZmu8e')

for resultado in lista_resultados:
    preco = resultado.find_element(By.CLASS_NAME, 'T14wmb').text
    nome = resultado.find_element(By.CLASS_NAME, 'sh-np__product-title').text
    link = resultado.find_element(By.CLASS_NAME, 'shntl').get_attribute('href')
    print(f'O preço: {preco}\nO nome: {nome} \n O link: {link}')


sleep(5)