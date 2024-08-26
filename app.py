from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import re
from time import sleep

navegador = webdriver.Chrome()
tabela_produtos = pd.read_excel("buscas.xlsx")

# Funções 
def verificar_termos_banidos(lista_termos_banidos, nome):
# Verifica se contém termos banidos
    tem_termos_banidos = False
    for termo in lista_termos_banidos:
        if termo in nome:
            tem_termos_banidos = True
            break
    return tem_termos_banidos

def verificar_tem_todos_termos_produtos(lista_termos_nomes_produto, nome):
    # Verifica se contém todos os termos do produto
    tem_todos_termos_produtos = True
    for termo in lista_termos_nomes_produto:
        if termo.lower() not in nome:
            tem_todos_termos_produtos = False
            break
    return tem_todos_termos_produtos

# Definição de Função de busca do Google Shopping
def busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nomes_produto = produto.split(' ')
    lista_ofertas = []

    navegador.get("https://www.google.com/")
    navegador.find_element(By.NAME, 'q').send_keys(produto, Keys.ENTER)

    sleep(4)
    elementos = navegador.find_elements(By.CLASS_NAME, 'YmvwI')

    for item in elementos:
        if 'Shopping' in item.text:
            item.click()
            break

    sleep(4)  # Tempo extra para garantir que a aba Shopping carregue

    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'KZmu8e')

    for resultado in lista_resultados:
        try:
            nome = resultado.find_element(By.CLASS_NAME, 'sh-np__product-title').text
        except:
            continue

        nome = nome.lower()

        # Verifica se contém termos banidos
        tem_termos_banidos = False
        for termo in lista_termos_banidos:
            if termo in nome:
                tem_termos_banidos = True
                break

        # Verifica se contém todos os termos do produto
        tem_todos_termos_produtos = True
        for termo in lista_termos_nomes_produto:
            if termo.lower() not in nome:
                tem_todos_termos_produtos = False
                break

        if not tem_termos_banidos and tem_todos_termos_produtos:
            try:
                # Use o seletor correto para o preço
                preco_bruto = resultado.find_element(By.CLASS_NAME, 'T14wmb').text
                # Verifica se o valor contém "R$" para garantir que é um preço
                if "R$" not in preco_bruto:
                    print(f"Elemento não é um preço válido: {preco_bruto}")
                    continue
                
                # Limpando o preço
                preco_limpo = re.sub(r'[^\d,]', '', preco_bruto)
                preco_limpo = preco_limpo.replace(',', '.')
                
                # Verifica se o preço tem formato válido
                if re.match(r'^\d+(\.\d{2})?$', preco_limpo):
                    preco = float(preco_limpo)
                else:
                    print(f"Formato de preço inválido: {preco_bruto}")
                    continue
                
            except Exception as e:
                print(f"Erro ao processar o preço: {e}")
                continue

            if preco_minimo <= preco <= preco_maximo:
                try:
                    link = resultado.find_element(By.TAG_NAME, 'a').get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
                except:
                    print("Link não encontrado")
                    continue

    return lista_ofertas

def busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nomes_produto = produto.split(' ')
    lista_ofertas = []

    navegador.get("https://www.buscape.com.br/")
    navegador.find_element(By.CLASS_NAME, 'AutoCompleteStyle_input__WAC2Y').send_keys(produto, Keys.ENTER)
    
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh')

    for resultado in lista_resultados:
        try:
            nome = resultado.find_element(By.CLASS_NAME, 'ProductCard_ProductCard_Name__U_mUQ').text
        except:
            continue
        nome = nome.lower()

        # analisar se ele não tem nenhum termo banido
        tem_termos_banidos = verificar_termos_banidos(lista_termos_banidos, nome)

        # analisar se ele tem TODOS os termos do nome produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produtos(lista_termos_nomes_produto, nome)

        # seleciona só os termos que tem_termos_banidos = False e ao meso tempo tem_todos_termos_produtos = False
        if not tem_termos_banidos and tem_todos_termos_produtos:
            try:
                # Aqui usamos um seletor mais específico com ambas as classes
                preco_bruto = resultado.find_element(By.CSS_SELECTOR, 'p.Text_Text__ARJdp.Text_MobileHeadingS__HEz7L').text
                
                # Verifica se o valor contém "R$" para garantir que é um preço
                if "R$" not in preco_bruto:
                    print(f"Elemento não é um preço válido: {preco_bruto}")
                    continue
                
                # Limpando o preço
                preco_limpo = re.sub(r'[^\d,]', '', preco_bruto)
                preco_limpo = preco_limpo.replace(',', '.')
                
                # Verifica se o preço tem formato válido
                if re.match(r'^\d+(\.\d{2})?$', preco_limpo):
                    preco = float(preco_limpo)
                else:
                    print(f"Formato de preço inválido: {preco_bruto}")
                    continue
                
            except Exception as e:
                print(f"Erro ao processar o preço: {e}")
                continue

            if preco_minimo <= preco <= preco_maximo:
                try:
                    link = resultado.get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
                except:
                    print("Link não encontrado")
                    continue

    return lista_ofertas

    sleep(5)

# Executa a aplicação
tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    # pesquisar pelo produto
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
    

    lista_ofertas_buscape = busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None

# Mandando a tabela para o Excel
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

# Enviando o E-mail
import win32com.client as win32

if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.to = 'seu-email@gmail.com'
    mail.Subject = 'Produto(s) encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,<p/>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada<p/>
    {tabela_ofertas.to_html(index=False)}
    <p>Att.,<p/>
    """
    mail.Send()
    navegador.quit()

sleep(3)
