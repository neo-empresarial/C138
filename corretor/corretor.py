import nltk

nltk.download('punkt')

with open('corretor/treinamento.txt', 'r', encoding='utf-8') as f:
    treinamento = f.read()

# print(treinamento[:500])

# print(len(treinamento))

texto_exemplo = 'Olá, tudo bem?'

tokens = texto_exemplo.split()

palavras_separadas = nltk.tokenize.word_tokenize(texto_exemplo)

def separa_palavras(lista_tokens): ## função para separar somente palavras, tirando caracteres especiais
    lista_palavras = []
    for token in lista_tokens:
        if token.isalpha():
            lista_palavras.append(token)
    return lista_palavras

palavras_separadas = separa_palavras(palavras_separadas)
# print(palavras_separadas)

lista_tokens = nltk.tokenize.word_tokenize(treinamento)

lista_palavras = separa_palavras(lista_tokens) ## lista_palavras armazena todas as palavras do treinamento tokenizadas, sem caracteres especiais

# print(len(lista_palavras))

def normalizacao(lista_palavras): ## função para normalizar as palavras, deixando todas em minúsculo
    lista_normalizada = []
    for palavra in lista_palavras:
        lista_normalizada.append(palavra.lower())
    return lista_normalizada

lista_normalizada = normalizacao(lista_palavras) ## lista_normalizada armazena todas as palavras do treinamento tokenizadas, sem caracteres especiais e em minúsculo
print(lista_normalizada[:10])

print(len(set(lista_normalizada))) ## set() retorna somente os valores únicos da lista, logo esse print retorna o total de palavras únicas do treinamento

