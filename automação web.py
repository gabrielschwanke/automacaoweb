#Objetivo do projeto é buscar os preços dos produtos smartphone apple iphone xr 64gb e placa de video rtx 3060 no google shopping e buscapé e enviar para o e-mail
#criando um navegador 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.by import By
import time
nav = webdriver.Chrome()
#importar e visualizar a base de dados
tabela_produtos = pd.read_excel('buscas.xlsx')
display(tabela_produtos)

def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    #entrar no google 
    nav.get('https://www.google.com/')

    #tratamento do nome do produto, para ficar tudo em minúsculo 
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()

    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_produto = produto.split(' ')
    
    #pesquisar o nome do produto
    nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto)
    nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
    #entrar na aba shopping
    
    elementos = nav.find_elements(By.CLASS_NAME,'hdtb-mitem')
    for elemento in elementos:
        if 'Shopping' in elemento.text:
            elemento.click()
            break
    #pegar o preço de apenas um produto
    
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')
    lista_ofertas = [] 
    for resultado in lista_resultados:

        nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text
        nome = nome.lower()
        
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
                
        #verificar se o nome tem todos os termos do nome do produto
        tem_termo_produto = True
        for palavra in lista_termos_produto:
            if palavra not in nome:
                tem_termo_produto = False

        if not tem_termos_banidos and tem_termo_produto:
            #verificando o nome
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace('R$','').replace(' ','').replace('.','').replace(',','.')
            preco = float(preco)

        #verificando se o preco ta dentro do mínimo e máximo

        preco_maximo = float(preco_maximo)
        preco_minimo = float(preco_minimo)
        if preco_minimo <= preco <= preco_maximo:

            elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
            elemento_pai = elemento_link.find_element(By.XPATH,'..')
            link = elemento_pai.get_attribute('href')
            lista_ofertas.append((nome, preco, link))
            
    return lista_ofertas

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    #tratar os valores da função
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_produto = produto.split(' ')
    
    #entrar no buscape 
    nav.get('https://www.buscape.com.br/')
    #pesquisar pelo produto no buscape
    nav.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(produto, Keys.ENTER)
    
    time.sleep(5)
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')
         
    lista_ofertas = []
    for resultado in lista_resultados:
        try:
            preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            nome = resultado.get_attribute('title')
            nome = nome.lower()
            link = resultado.get_attribute('href')
            
            #verificação do nome - se o nome tem algum termo banido
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True

            #verificar se o nome tem todos os termos do nome do produto
            tem_termo_produto = True
            for palavra in lista_termos_produto:
                if palavra not in nome:
                    tem_termo_produto = False

            if not tem_termos_banidos and tem_termo_produto:
                preco = preco.replace('R$','').replace(' ','').replace('.','').replace(',','.')
                preco = float(preco)
                if preco_minimo <= preco <= preco_maximo:
                    lista_ofertas.append((nome, preco, link))
                    print(preco, nome, link)
        except:
            pass
        
    return lista_ofertas

tabela_ofertas = pd.DataFrame()#criando tabela vazia
for linha in tabela_produtos.index:
    
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']
    #transformar listas de tuplas em planilha
    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto','preco','link'])
        tabela_ofertas = tabela_ofertas.append(tabela_google_shopping)
    else:
        tabela_google_shopping = None
#print(lista_ofertas_google_shopping)
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto','preco','link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)
    else:
        tabela_buscape = None
    #display(tabela_google_shopping)
    #display(tabela_buscape)
display(tabela_ofertas)    
#print(lista_ofertas_buscape)
#exportar pro excel
#arrumando os indices da tabela
tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

#enviar por email o resultado da tabela
import win32com.client as win32
#verificando se existe alguma oferta dentro da minha tabela_ofertas
if len(tabela_ofertas.index) > 0:
    #enviar o e-mail
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.to = 'gabriel-pereira-995@hotmail.com'
    mail.subject = 'Produto(s) encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f'''
    <p>Prezados,</p>
    <p>Encontramos alguns produtos na faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att.,</p>
    '''
    mail.Send()
nav.quit()#para fechar a janela
    

