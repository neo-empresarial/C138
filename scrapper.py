import requests
from bs4 import BeautifulSoup
from lxml import etree
from lxml import html
import pandas as pd

def scrapper():
  #Usando o URL apresentado no documento de calibração
  url = 'http://www.inmetro.gov.br/laboratorios/rbc/detalhe_laboratorio.asp?num_certificado=34&situacao=AT&area=DIMENSIONAL'

  response = requests.get(url)

  html_content = response.content

  #Usando o BeautifulSoup para fazer o parse do HTML
  soup = BeautifulSoup(html_content, 'html.parser')

  #Aqui, juntamos o BeautifulSoup com o lxml para fazer o parse do HTML, já que o lxml aceita XPATH
  html_tree = html.fromstring(str(soup))

  #Aqui, usamos o XPATH para pegar a tabela que queremos
  table_rows = html_tree.xpath('//table[4]/tr')

  #Aqui, criamos uma lista vazia para armazenar os dados da tabela
  rows_data = []

  #Aqui, iteramos sobre as linhas da tabela e pegamos o texto de cada célula
  for row in table_rows:
    cells = row.xpath('.//td|.//th')
    row_data = [cell.text_content().strip() for cell in cells]
    rows_data.append(row_data)

  #Aqui, criamos um DataFrame do Pandas com os dados da tabela
  df = pd.DataFrame(rows_data, columns = None, index = None) 

  df = df.dropna(axis=1, how='all') #Aqui, removemos as colunas que só tem valores nulos

  df = df.dropna(axis=0, how='all') #Aqui, removemos as linhas que só tem valores nulos

  # df = df[~df[1].astype(str).str.startswith('Método')]
  # print(df)

  df = df.replace('Medição de', 'Medir', regex=True)
  df = df.replace('Medição por', 'Medir', regex=True)
  df = df.replace('para Medir', 'de Medir', regex=True)
  df.columns = ['Descrição do serviço', 'Parâmetro, Faixa e Método', 'Capacidade de Medição e Calibração (CMC)']

  for i, row in df.iterrows():
    if pd.isna(df.at[i, 'Descrição do serviço']) or row['Descrição do serviço'] == '':
      df.at[i, 'Descrição do serviço'] = df.at[i-1, 'Descrição do serviço']

  return df

df_web = scrapper()
print(df_web.head(50))