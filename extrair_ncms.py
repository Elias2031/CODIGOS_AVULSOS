import requests
import re
from bs4 import BeautifulSoup

url = "https://www.legisweb.com.br/legislacao/?id=123549#ini-legis"  # Substitua pela URL da página
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Encontrar todos os elementos que contenham o formato desejado
elements = soup.find_all(string=re.compile(r'\d{4}\.\d{2}\.\d{2}'))
ncms = []

for element in elements:
    print(element)
    # Verificar se o elemento é uma string e se corresponde ao padrão
    matches = re.findall(r'\d{4}\.\d{2}\.\d{2}', element)
    if matches:
        # print(matches)
        ncms.extend(matches)  # Adiciona as correspondências à lista

# Remover duplicatas convertendo a lista em um conjunto
ncms = list(set(ncms))

ncms.sort()

# Exibir os resultados
# print(ncms)

# Salvar os resultados em um arquivo
with open('NCMS_EXTRAIDOS.TXT', 'a') as arquivo:
    for ncm in ncms:
        arquivo.write(ncm + '\n')
