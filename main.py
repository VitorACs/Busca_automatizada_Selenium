#importar as bibliotecas
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pandas as pd
import time

# criar o navegador
servico = Service(ChromeDriverManager().install())
nav = webdriver.Chrome(service=servico)

# importar base de dados 
produtos_busca_df = pd.read_excel('buscas.xlsx')

def buscar_google (nav, produto, termos_banidos, preco_min, preco_max):
    #entrar no google 
    nav.get('https://www.google.com.br/')
    
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_produtos = produto.split(' ')

    nav.find_element(By.XPATH,'//*[@id="APjFqb"]').send_keys(produto,Keys.ENTER)

    # entrar no shopping 
    time.sleep(5)
    elementos = nav.find_elements(By.CLASS_NAME,'O3S9Rb')
    for elemento in elementos:
        if 'Shopping' in elemento.text:
            elemento.click()
            break
            
    # colocar o preço min e max
    nav.find_element(By.NAME,'lower').send_keys(str(preco_min))
    nav.find_element(By.NAME,'upper').send_keys(str(preco_max),Keys.ENTER)

    time.sleep(3)
    # pegar os resultados
    lista_resultado = nav.find_elements(By.CLASS_NAME,'i0X6df')
    
    lista_oferta1 = []
    for resultado in lista_resultado:
        nome  = resultado.find_element(By.CLASS_NAME,'tAxDx').text
        
        # analisar se ele não tem nenhum temo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra.lower() in nome.lower():
                tem_termos_banidos = True

        # analisar se tem todos os termos do produto
        tem_termos_produtos = True
        for palavra in lista_termos_produtos:
            if palavra.lower() not in nome.lower():
                tem_termos_produtos = False

        # selecionar os termos onde BANIDO = FALSE E PRODUTOS = TRUE
        if tem_termos_banidos == False and tem_termos_produtos == True:
            try:
                preco = resultado.find_element(By.CLASS_NAME,'a8Pemb').text
                preco = preco.replace('R$','').replace(' ','').replace('.','').replace(',','.')
                preco = float(preco)

                # usar modo alternativo para pegar o link 
                elemento_referencia = resultado.find_element(By.CLASS_NAME,'bONr3b')
                elemento_pai = elemento_referencia.find_element(By.XPATH,'..')
                link  = elemento_pai.get_attribute('href')
                lista_oferta1.append((nome,preco,link))
            except:
                continue
                
    return lista_oferta1

def buscar_buscape (nav, produto, termos_banidos, preco_min, preco_max):
    nav.get('https://www.buscape.com.br/')
    
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_produtos = produto.split(' ')
    
    time.sleep(5)
    nav.find_element(By.XPATH,'//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto,Keys.ENTER)
    
    time.sleep(3)
    lista_resultados = nav.find_elements(By.CLASS_NAME,'SearchCard_ProductCard_Inner__7JhKb')
    lista_oferta2=[]
    for resultado in lista_resultados:

        nome = resultado.find_element(By.CLASS_NAME,'SearchCard_ProductCard_NameWrapper__Gv0x_').text

        # analisar se ele não tem nenhum temo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra.lower() in nome.lower():
                tem_termos_banidos = True

        # analisar se tem todos os termos do produto
        tem_termos_produtos = True
        for palavra in lista_termos_produtos:
            if palavra.lower() not in nome.lower():
                tem_termos_produtos = False

        # selecionar os termos onde BANIDO = FALSE E PRODUTOS = TRUE
        if tem_termos_banidos == False and tem_termos_produtos == True:
            try:
                preco = resultado.find_element(By.CLASS_NAME,'Text_MobileHeadingS__Zxam2').text
                preco = preco.replace('R$','').replace(' ','').replace('.','').replace(',','.')
                preco = float(preco)

                if preco_min <= preco <= preco_max:
                    link = resultado.get_attribute('href')
                    lista_oferta2.append((nome,preco,link))
            except:
                continue

    return lista_oferta2   

dataframe = pd.DataFrame()
for i,produto in enumerate(produtos_busca_df['Nome']):
    
    produto = str(produto)
    termos_banidos = produtos_busca_df.loc[i,'Termos banidos']

    preco_min = produtos_busca_df.loc[i,'Preço mínimo']
    preco_max = produtos_busca_df.loc[i,'Preço máximo']
    
    lista_buscape = buscar_buscape (nav, produto, termos_banidos, preco_min, preco_max)
    if lista_buscape:
        tabela_buscape_shopping = pd.DataFrame(lista_buscape, columns=['produto', 'preco', 'link'])
        dataframe = dataframe.append(tabela_buscape_shopping)
    else:
        tabela_buscape_shopping = None
        
    lista_google = buscar_google (nav, produto, termos_banidos, preco_min, preco_max)
    if lista_buscape:
        tabela_google_shopping = pd.DataFrame(lista_google, columns=['produto', 'preco', 'link'])
        dataframe = dataframe.append(tabela_google_shopping)
    else:
        tabela_google_shopping = None

tabela_ofertas = dataframe.reset_index(drop=True)
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)