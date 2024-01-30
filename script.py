import pandas as pd 
import openpyxl
import datetime
import nltk
from nltk import tokenize
import requests
from bs4 import BeautifulSoup
from lxml import etree
from lxml import html
import re
import ast

workbook = openpyxl.load_workbook('certificados.xlsx/Trena a laser.xlsx', data_only=True)
sheet = workbook.active

# Get the max row count
max_row = sheet.max_row

# Get the max column count
max_column = sheet.max_column

#Função para pegar a última linha com conteúdo
def get_last_row(sheet):
    for i in range(max_row, 0, -1):
        row_values = [cell.value for cell in sheet[i]]

        if any(row_values):
            return i
        
    return None

last_row = get_last_row(sheet)

#Função para pegar a última coluna com conteúdo
def get_last_column(sheet):	
    for i in range(max_column, 0, -1):
        column_values = [cell.value for cell in sheet[i]]

        if any(column_values):
            return i
        
    return None

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
    df = df.replace('Medição por', 'Medir por', regex=True)
    df = df.replace('para Medir', 'de Medir', regex=True)
    df.columns = ['Descrição do serviço', 'Parâmetro, Faixa e Método', 'Capacidade de Medição e Calibração (CMC)']

    for i, row in df.iterrows():
        if pd.isna(df.at[i, 'Descrição do serviço']) or row['Descrição do serviço'] == '':
            df.at[i, 'Descrição do serviço'] = df.at[i-1, 'Descrição do serviço']

    # df = df.to_excel('web.xlsx', index=False)

    return df

def convert_to_float(value):
    try:
        return float(value.split()[0].replace(',', '.'))
    except:
        return value

def process_string(string):
    result = re.findall(r'\d+', string)

    if result:
        if len(result) != 1:
            return float(result[0]), float(result[1])
        else:
            return 0.0, float(result[0])
    else:
        return None

last_column = get_last_column(sheet)

print('--------------------------------------------------')

centro_found = False
padroes_found = False
capa_data = []

for i in range (1, last_row + 1):
    row_values = []
    for j in range (1, last_column + 1):
        cell_obj = sheet.cell(row = i, column = j)

        if str(cell_obj.value).startswith('CENTRO'):
            centro_found = True
            continue
        if str(cell_obj.value).startswith('Padrões utilizados'):
            padroes_found = True
            continue

        if centro_found and not padroes_found:
            row_values.append(cell_obj.value)
            if len(row_values) == 9:
                # print(f'valor da linha {i}: {row_values}')
                capa_data.append(row_values)

    if padroes_found:
        break

def create_df_capa():
    df_capa = pd.DataFrame(capa_data)
    df_capa = df_capa.dropna(axis=1, how='all')
    df_capa = df_capa.dropna(axis=0, how='all')
    df_capa = df_capa[~df_capa.apply(lambda row: 'Ocultar' in row.values, axis=1)]
    df_capa = df_capa.drop_duplicates().reset_index(drop=True)

    return df_capa
# df_capa = pd.DataFrame(capa_data)
# df_capa = df_capa.dropna(axis=1, how='all')
# df_capa = df_capa.dropna(axis=0, how='all')
# df_capa = df_capa[~df_capa.apply(lambda row: 'Ocultar' in row.values, axis=1)]
# df_capa = df_capa.drop_duplicates().reset_index(drop=True)

df_capa = create_df_capa()

print("Informações da capa:")
print(df_capa)


print('--------------------------------------------------')

#Aqui, vamos tentar extrair os as máquinas utilizadas na medição

padroes_found = False
procedimento_found = False
padroes_data = []

for i in range(1, last_row + 1):
    row_values = []
    for j in range(1, last_column + 1):
        cell_obj = sheet.cell(row = i, column = j)

        if str(cell_obj.value) == 'Padrões utilizados':
            padroes_found = True
            continue

        if str(cell_obj.value).startswith('Procedimento'):
            procedimento_found = True
            continue

        if padroes_found and not procedimento_found:
            row_values.append(cell_obj.value)
            if len(row_values) == 9:
                padroes_data.append(row_values)

    if procedimento_found:
        break

df_padroes = pd.DataFrame(padroes_data)
df_padroes = df_padroes.dropna(axis=1, how='all')

descricao_column = df_padroes[df_padroes.eq('Descrição').any(axis = 1)].stack().index[1][1]

start_row = df_padroes[df_padroes[descricao_column] == 'Descrição'].index[0]
value_below_descricao = df_padroes.loc[start_row + 1][descricao_column]

result_values = [value_below_descricao]
current_row = start_row + 1

while current_row < len(df_padroes) and df_padroes.loc[current_row, descricao_column] is not None:
    result_values.append(df_padroes.loc[current_row, descricao_column])
    current_row += 1

machines_df = pd.DataFrame(result_values, columns=[descricao_column])

machines_df = machines_df.dropna(how='all', inplace=False)
machines_df = machines_df.drop_duplicates().reset_index(drop=True)

machines_df = machines_df[machines_df[descricao_column] != '#N/A']
machines_df.columns = ['Descrição do serviço']

print('Máquinas utilizadas:')
print(machines_df)

print('--------------------------------------------------')

#Nesse bloco de código, fazemos um loop para pegar somente os valores tabelados de medição do certificado

resultados_found = False
observacoes_found = False
data = []

for i in range(1, last_row + 1):
    row_values = []

    for j in range(1, last_column + 1):
        cell_obj = sheet.cell(row = i, column = j)

        if str(cell_obj.value) == 'Resultados':
            resultados_found = True
            continue
    
        if str(cell_obj.value) == 'Observações':
            observacoes_found = True
            continue

        if resultados_found and not observacoes_found:
            row_values.append(cell_obj.value)
            if len(row_values) == 9:
                data.append(row_values)

    if observacoes_found:
        break

df_dados = pd.DataFrame(data)
df_dados = df_dados.dropna(axis=1, how='all')
df_dados = df_dados.dropna(axis=0, how='all')
df_dados = df_dados[~df_dados.apply(lambda row: 'Ocultar' in row.values, axis=1)]
df_dados = df_dados.drop_duplicates().reset_index(drop=True)

#Vamos tentar separar as diferentes tabelas que temos dentro do df_dados

new_table_indices = df_dados[df_dados.apply(lambda row: any(cell and str(cell).startswith('Valor') for cell in row.values), axis=1)].index

# Extract tables
tables = []

for i in range(len(new_table_indices)):
    start_idx = new_table_indices[i]
    end_idx = new_table_indices[i+1] if i+1 < len(new_table_indices) else len(df_dados)
    table = df_dados.iloc[start_idx:end_idx, :].reset_index(drop=True)
    tables.append(table)

print('Dados da medição:')
print(df_dados)
print(tables)

#O índice dentro do [] indica qual tabela específica vamos printar (em ordem de cima para baixo no documento)
#Ou seja, tables separa as tabelas que temos dentro do df_dados

print('--------------------------------------------------')

df_web = scrapper()
print('df_web:')
print(df_web)

print('--------------------------------------------------')

# Quando damos um merge() nos dataframes, conseguimos um resultado semelhante ao de um PROCV
# Ou seja, aqui estamos fazendo um PROCV entre site e máquina, logo temos a regra da máquina utilizada para cada serviço

df_merge = pd.merge(machines_df, df_web, on='Descrição do serviço', how='left')
df_merge = df_merge.drop_duplicates().reset_index(drop=True)

print('PROCV entre df_web e machines_df:')
print(df_merge)

print('--------------------------------------------------')

df_capa_merge = df_capa.copy()
df_web_merge = df_web.copy()

df_capa_merge[df_capa_merge.columns[1]] = df_capa_merge[df_capa_merge.columns[1]].str.lower()
df_web_merge[df_web_merge.columns[0]] = df_web_merge[df_web_merge.columns[0]].str.lower()

df_service = pd.merge(df_capa_merge, df_web_merge, left_on=df_capa_merge.iloc[:, 1], right_on=df_web_merge.iloc[:, 0], how='inner')
df_service = df_service.drop_duplicates().reset_index(drop=True)


print('--------------------------------------------------')

print('PROCV entre df_capa e df_web:')
print(df_service)

print('--------------------------------------------------')

#Nesse ponto do código, temos todas as informações que precisamos para fazer os cálculos

#Vamos transformar a coluna de medições em números, extrair o maior valor e usá-lo para saber qual regra usar

print(tables)

first_column = tables[0].iloc[:, 0]
    
for i, table in enumerate(tables):
    tables[i].iloc[:, 0] = table.iloc[:, 0].apply(convert_to_float)

first_column.columns = ['Resultados']

print(first_column)

last_value = first_column.iloc[-1]
print(last_value)

#Temos o maior valor (sempre o mais inferior da tabela), agora vamos usar ele para saber qual regra usar

print('--------------------------------------------------')

# Vamos tentar resolver uma questão sobre trabalhos feitos em campo ou em laboratório
# Imaginamos que, para trabalhos feitos em campo, existe alguma célula com um texto específico

print('Separando o df_web em dois data frames diferentes:')

df_web_split = df_web.copy()

indices = df_web_split[df_web_split['Descrição do serviço'] == 'INSTRUMENTOS E GABARITOS DE MEDIÇÃO DE ÂNGULO'].index[1]

df_web_lab = df_web_split.iloc[:indices - 2]
df_web_field = df_web_split.iloc[indices - 2:]

print('df para medições em laboratório:')
print(df_web_lab)

print('--------------------------------------------------')

print('df para medições em campo:')
print(df_web_field)

print('--------------------------------------------------')

#Vamos descobrir se o certificado é de medição em campo ou em laboratório

if 'LOCAL DA CALIBRAÇÃO' in df_capa.iloc[:, 0].astype(str).values:
    print('Certificado de medição em campo')
    working_df = df_web_field
else:
    print('Certificado de medição em laboratório')
    working_df = df_web_lab

print('Dataframe correto para a medição apresentada no certificado:')
print(working_df)

#Assim, working_df vai sempre armazenar o dataframe correto para utilizarmos no merge()

print('--------------------------------------------------')

# Vamos partir para o merge() e assim ter as informações de erro

working_df[working_df.columns[0]] = working_df[working_df.columns[0]].str.lower()

print(working_df)
print(df_capa_merge)

df_merge_service = pd.merge(df_capa_merge, working_df, left_on=df_capa_merge.iloc[:, 1], right_on=working_df.iloc[:, 0], how='inner')

print(df_merge_service)

df_merge_service = df_merge_service.drop_duplicates().reset_index(drop=True)
df_merge_service.iloc[:, -1] = df_merge_service.iloc[:, -1].str.replace('*', '')
df_merge_service = df_merge_service.dropna(axis=0, how='all')
df_merge_service = df_merge_service.dropna(axis=1, how='all')

print('PROCV entre df_capa e df_web:')
print(df_merge_service)

print('--------------------------------------------------')

#Aqui, vamos pegar o valor de erro e o valor de incerteza

# parameters_values = df_merge_service.iloc[:, -1].values

# non_empty_paramenters_values = [value for value in parameters_values if pd.notna(value) and value is not None and value != '']

# print('Valores de erro e incerteza:')
# print(non_empty_paramenters_values)

#Temos os valores de erro e incerteza, sendo eles strings

# float_parameters_values = [float(value.split()[0].replace(',', '.')) for value in non_empty_paramenters_values]
# print('Valores de erro e incerteza em float:')
# print(float_parameters_values)

print('--------------------------------------------------')

#Vamos para a parte difícil: extrair o intervalo numérico a partir da string

# range_strings = df_merge_service.iloc[:, 4].values
# print(range_strings)

df_merge_service['Intervalo'] = df_merge_service.iloc[:, 4].apply(process_string)
df_merge_service = df_merge_service.dropna(axis=0, how='any')
df_merge_service = df_merge_service.dropna(axis=1, how='any')  

print('Intervalos numéricos:')
print(df_merge_service)

print('--------------------------------------------------')

#Agora que temos todas as peças, vamos ao que importa: os cálculos
#Queremos pegar cada valor da first_column e ver em qual intervalo ele se encaixa
#Depois, vamos pegar o valor de erro e incerteza correspondente

#PONTO DE ERRO

def process_cmc_information(cmc_value):
    cmc_value = str(cmc_value).strip()
    
    # Check if it's a distance (type 1)
    distance_match = re.match(r'([\d.,]+)\s*([µm]+)', cmc_value)
    if distance_match:
        print('Caso 1 utilizado')
        value = float(distance_match.group(1).replace(',', '.'))
        unit = distance_match.group(2)
        return value, unit

    # Check if it's an equation (type 2)
    equation_match = re.match(r'\[([\s\S]+)\]', cmc_value)
    if equation_match:
        print('Caso 2 utilizado')
        return equation_match.group(1)

    # Check if it's an angle (type 3)
    angle_match = re.match(r'\s*(\d+)\s*\'\'\s*', cmc_value)
    if angle_match:
        print('Caso 3 utilizado')
        return float(angle_match.group(1))

    # Check if it's a percentage (type 4)
    percentage_match = re.match(r'([\d.,]+)%', cmc_value)
    if percentage_match:
        print('Caso 4 utilizado')
        return float(percentage_match.group(1).replace(',', '.')) / 100

    # Default case: return the original value
    return cmc_value


df_merge_service['Capacidade de Medição e Calibração (CMC)'] = df_merge_service['Capacidade de Medição e Calibração (CMC)'].apply(process_cmc_information)

print(df_merge_service)
#LEMBRETE -> Revisar a função abaixo (ainda não está funcionando)
def get_error_and_uncertainty(valor, intervalos):
    for intervalo in intervalos:
        if intervalo[0] <= valor <= intervalo[1]:
            return intervalo[0]
    return None


# first_column['Valor Correspondente'] = first_column.apply(lambda x: get_error_and_uncertainty(float(x) if x is not None else None, df_merge_service['Intervalo']))

single_row = df_merge_service.iloc[0]

print('Primeira coluna (?)')
print(single_row)

# Iterate through each table in the 'tables' list
for i, row in df_merge_service.iterrows():
    # Iterate through each table in the 'tables' list
    for j, table in enumerate(tables):
        # Get the numerical range for the current row and table
        intervalo = tuple(row['Intervalo'])
        
        # Create a new column with True if the value is within the range, False otherwise
        table[f'Within_Range_{i + 1}'] = pd.to_numeric(table.iloc[:, 0], errors='coerce').between(*intervalo)

# Display the modified tables
for j, table in enumerate(tables):
    print('Tabelas enumeradas:')
    print(f"Table {j + 1}:\n{table}\n")

print('--------------------------------------------------')

#Agora, vamos coorelacionar df_merge_service com as tabelas

for i, table in enumerate(tables):
    # Create a new column to store the selected values
    table['Selected_Value'] = ""

    # Iterate through the rows of the table
    for index, row in table.iterrows():
        selected_value = ""

        # Iterate through the columns and find the first True value in boolean columns
        for col in table.columns:
            if isinstance(col, str) and col.startswith('Within_Range') and row[col]:
                # Extract the range index from the column name
                range_index = int(col.split('_')[-1]) - 1

                # Check if the range_index is within the valid range of rows in df
                if 0 <= range_index < len(df_merge_service):
                    # Get the value from the 6th column of df based on the range_index
                    selected_value = df_merge_service.iloc[range_index, -2]  # Assuming the 6th column index is 5
                    break  # Exit the loop once a match is found

        # Update the 'Selected_Value' column with the corresponding value from df
        table.at[index, 'Selected_Value'] = selected_value

# Display the modified tables
for i, table in enumerate(tables):
    print(f"Table {i + 1}:\n{table}\n")

print('--------------------------------------------------')

print('Tabelas com os valores selecionados:')
print(tables)

print('--------------------------------------------------')

#Aqui, vamos pegar os valores de erro e incerteza e colocar em uma coluna separada

for i, table in enumerate(tables):
    u_column_index = table.columns[table.iloc[0].astype(str).str.startswith('U')].tolist()
    if u_column_index:
        u_column_index = u_column_index[0]
        print(u_column_index)
        break
    
has_meters = '[m]' in table.iloc[:, u_column_index].values
has_mm = '[mm]' in table.iloc[:, u_column_index].values
has_µm = '[µm]' in table.iloc[:, u_column_index].values


for i in range(len(tables)):
    table = tables[i]

    if 'Selected_Value' in table.columns:
        table[['CMC_Value', 'CMC_Unit']] = table['Selected_Value'].apply(pd.Series)

        table = table.drop('Selected_Value', axis=1)

        tables[i] = table
    else:
        print('Não tem Selected_Value')

print(tables)

print ('--------------------------------------------------')

#Agora que temos os valores de CMC e as unidades separadas, podemos trabalhar com as colunas

def convert_to_meters(row):
    value, unit = row['CMC_Value'], row['CMC_Unit']
    if has_meters and unit == 'mm':
        return value / 1000  
    elif has_meters and unit == 'µm':
        return value / 1000000  
    else:
        return value

def convert_to_mm(row):
    value, unit = row['CMC_Value'], row['CMC_Unit']
    if has_mm and unit == 'm':
        return value * 1000  
    elif has_mm and unit == 'µm':
        return value / 1000  
    else:
        return value

def convert_to_µm(row):
    value, unit = row['CMC_Value'], row['CMC_Unit']
    if has_µm and unit == 'm':
        return value * 1000000  
    elif has_µm and unit == 'mm':
        return value / 1000  
    else:
        return value
    
for i in range(len(tables)):
    table = tables[i]
    if 'CMC_Value' in table.columns and 'CMC_Unit' in table.columns and has_meters:
        table['CMC_Value'] = table.apply(convert_to_meters, axis=1)
        tables[i] = table
    elif 'CMC_Value' in table.columns and 'CMC_Unit' in table.columns and has_mm:
        table['CMC_Value'] = table.apply(convert_to_mm, axis=1)
        tables[i] = table
    elif 'CMC_Value' in table.columns and 'CMC_Unit' in table.columns and has_µm:
        table['CMC_Value'] = table.apply(convert_to_µm, axis=1)
        tables[i] = table
    else:
        print('Não tem CMC_Value ou CMC_Unit')

for i, table in enumerate(tables):
    tables[i].iloc[:, u_column_index] = table.iloc[:, u_column_index].apply(convert_to_float)



print(tables)

print('--------------------------------------------------')

#Vamos comparar a nossa coluna CMC_Value com os valores de U

new_column_name = 'U'
for i in range(len(tables)):
    tables[i].columns.values[u_column_index] = new_column_name

for i in range(len(tables)):
    table = tables[i]
    table['CMC_Value'] = table['CMC_Value'].replace('', None)
    table['CMC_Verification'] = None
    table['Range_Verification'] = None

print(tables)

# Compare the specified columns for each DataFr

error_ocurred = False

for i, table in enumerate(tables):
    for index, row in table.iterrows():
        u_column_value = row['U']
        # print(type(u_column_value))
        cmc_value = row['CMC_Value']
        # print(type(cmc_value))

        if pd.notna(cmc_value) and pd.notna(u_column_value):
            if u_column_value >= cmc_value:
                print('ok')
                table.at[index, 'CMC_Verification'] = True
            elif u_column_value  < cmc_value:
                print('error')
                table.at[index, 'CMC_Verification'] = False
        elif pd.isna(cmc_value) and type(u_column_value) != str and pd.notna(u_column_value):
            table.at[index, 'Range_Verification'] = True

print(tables)

print('--------------------------------------------------')


