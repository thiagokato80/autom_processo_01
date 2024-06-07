# Importar bibliotecas necessarias

# Criar o navegador
from selenium import webdriver

# Buscadores
from Buscadores import *

# Bibliotecas necessárias para envio do e-mail pela API do Gmail
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from google_auth_oauthlib.flow import InstalledAppFlow
import pybase64

# Tirei os avisos para rodar mais limpo, contudo pode ser necessário para avaliação futura
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

import pandas as pd

# Abrir o navegador do Chrome
nav = webdriver.Chrome()

# Importar e visualizar produtos a serem buscados na WEB (realizaremos apenas busca no GoogleShopping e Buscape)
tabela_produtos = pd.read_excel("Produtos.xlsx")

tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, "Nome"]
    termos_banidos = tabela_produtos.loc[linha, "Termos banidos"]
    preco_minimo = tabela_produtos.loc[linha, "Preço mínimo"]
    preco_maximo = tabela_produtos.loc[linha, "Preço máximo"]
    
    lista_ofertas_google_shopping = Buscadores.busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_google_shopping)
    else:
        tabela_google_shopping = None
        
    lista_ofertas_buscape = Buscadores.busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)
    else:
        tabela_buscape = None

#print(tabela_ofertas) # Caso queira verificar a tabela de ofertas

# Exportar tabela de ofertas para o Excel
tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)

# Carregue as credenciais do Google API
creds = None
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/gmail.compose']


# Carrega as credeciais - Roda novamente cada vez que o programa é disparado
credentials_file = 'credenciais.json' # Utilize suas próprias credenciais (https://console.cloud.google.com/apis/dashboard)
flow = InstalledAppFlow.from_client_secrets_file(
            credentials_file, SCOPES)
creds = flow.run_local_server(port=0)

# Crie um objeto de serviço para a API do Gmail
service = build('gmail', 'v1', credentials=creds)

# enviar por e-mail o resultado da tabela
# verificando se existe alguma oferta dentro da tabela de ofertas
if len(tabela_ofertas.index) > 0:
    # vou enviar email
    mensagem = MIMEMultipart()
    mensagem['From'] = 'thiagokato@gmail.com'
    mensagem['To'] = 'thiagokato@gmail.com'
    mensagem['Subject'] = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    corpo_html = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att.,</p>
    """
    
    # Adicione a parte HTML do corpo da mensagem
    parte_html = MIMEText(corpo_html, 'html')
    mensagem.attach(parte_html)

    # Converta a mensagem em uma string codificada em Base64
    mensagem_bytes = mensagem.as_bytes()
    mensagem_codificada = pybase64.b64encode(mensagem_bytes).decode('utf-8')

    # Envie o e-mail
    requisicao = {
        'raw': mensagem_codificada
    }
    service.users().messages().send(userId='me', body=requisicao).execute()

nav.quit()  