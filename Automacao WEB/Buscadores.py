from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import time

class Buscadores:

    # Atualizar se necessário o XPATH, CLASS_NAME, etc.
    def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
        # Entrar no google
        nav.get("https://shopping.google.com/")
        
        # Tratar os valores que vieram da tabela
        produto = produto.lower()
        termos_banidos = termos_banidos.lower()
        lista_termos_banidos = termos_banidos.split(" ")
        lista_termos_produto = produto.split(" ")
        preco_maximo = float(preco_maximo)
        preco_minimo = float(preco_minimo)
        

        # Pesquisar o nome do produto no google
        nav.find_element(By.ID, 
                        'REsRA').send_keys(produto)
        nav.find_element(By.ID, 
                        'REsRA').send_keys(Keys.ENTER)

        # Carregar a lista de resultados da busca no Google Shopping
        lista_resultados = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')
        
        # Para cada resultado, verificar se atende a todas as condicoes
        lista_ofertas = [] # Lista com as respostas
        for resultado in lista_resultados:
            nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
            nome = nome.lower()

            # Verificar o nome - se contem algum termo banido
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True
            
            # Verificar se o nome contem todos os termos do nome do produto
            tem_todos_termos_produto = True
            for palavra in lista_termos_produto:
                if palavra not in nome:
                    tem_todos_termos_produto = False

            if not tem_termos_banidos and tem_todos_termos_produto: # Condicao apos verificacao do nome
                try:
                    preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                    preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                    preco = float(preco)
                    
                    # Verificar se o preco esta dentro do range de preco
                    if preco_minimo <= preco <= preco_maximo:
                        elemento_link = resultado.find_element(By.CLASS_NAME, 'mnIHsc')
                        elemento_pai = elemento_link.find_elements(By.CLASS_NAME, 'shntl')
                        link = elemento_pai[0].get_attribute('href')
                        lista_ofertas.append((nome, preco, link)) # Caso verdadeiro inclui na lista de oferta
                except:
                    continue
                
        return lista_ofertas
        

    def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
        # Tratar os valores que vieram da tabela
        preco_maximo = float(preco_maximo)
        preco_minimo = float(preco_minimo)
        produto = produto.lower()
        termos_banidos = termos_banidos.lower()
        lista_termos_banidos = termos_banidos.split(" ")
        lista_termos_produto = produto.split(" ")
        
        
        # Entrar no buscape
        nav.get("https://www.buscape.com.br/")
        
        # Pesquisar pelo produto no Buscape
        nav.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)
        
        # Carregar a lista de resultados da busca no Buscape
        time.sleep(6) # Aguardar carregar, dependendo da velocidade da internet
        lista_resultados = nav.find_elements(By.CLASS_NAME, 'Hits_ProductCard__Bonl_')
        
        # Para cada resultado, verificar se atende a todas as condicoes
        lista_ofertas = []
        for resultado in lista_resultados:
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__HEz7L').text
                nome = resultado.find_element(By.CLASS_NAME, 'Text_Text__ARJdp').text
                nome = nome.lower()

                link_class = resultado.find_element(By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh')
                link = link_class.get_attribute('href')
                
                # Verificar o nome - se contem algum termo banido
                tem_termos_banidos = False
                for palavra in lista_termos_banidos:
                    if palavra in nome:
                        tem_termos_banidos = True  
                        
                # Verificar se o nome contem todos os termos do nome do produto
                tem_todos_termos_produto = True
                for palavra in lista_termos_produto:
                    if palavra not in nome:
                        tem_todos_termos_produto = False            
                
                if not tem_termos_banidos and tem_todos_termos_produto:
                    preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                    preco = float(preco)
                    # Verificar se o preco esta dentro do range de preco
                    if preco_minimo <= preco <= preco_maximo:
                        lista_ofertas.append((nome, preco, link)) # Caso verdadeiro inclui na lista de oferta
            except:
                pass
        return lista_ofertas