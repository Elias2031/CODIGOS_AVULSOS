import pandas as pd
import re
import unidecode

def remove_accents_and_clean(text):
    return unidecode.unidecode(text)

def jaccard_similarity(str1, str2):
    # Remove acentos, caracteres especiais e converte para lowercase
    str1 = remove_accents_and_clean(str(str1).lower().replace('%',''))
    str2 = remove_accents_and_clean(str(str2).lower().replace('%',''))
    
    # Criar conjuntos de caracteres (ou bigramas/trigramas se necessário)
    set1 = set(str1)
    set2 = set(str2)
    
    # Calcula a interseção e a união
    intersection = set1.intersection(set2)
    union = set1.union(set2)
    
    # Calcula a similaridade de Jaccard
    similaridade = len(intersection) / len(union)
    
    print('NOME COLUNA =>', str1, 'NOME RUBRICA =>', str2, 'SIMILARIDADE =>', similaridade)
    
    return similaridade

# Carregar o arquivo "folha de ponto"
arquivo_folha = 'Planilha ((Fechamento da Folha)) LINK - 11.2024 - CONTABILIDADE.xlsm (7).xlsx'
dataframe_folha_ponto = pd.read_excel(arquivo_folha)
dataframe_folha_ponto.columns = dataframe_folha_ponto.columns.str.lower().str.replace(' ', '')

# Carregar o arquivo "funcionários"
arquivo_funcionarios = 'Trabalhadores.XLS'
dataframe_arquivo_funcionarios = pd.read_excel(arquivo_funcionarios, engine='xlrd', usecols=[0, 2])
dataframe_arquivo_funcionarios.columns = dataframe_arquivo_funcionarios.columns.str.lower().str.replace(' ', '')

# Carregar o arquivo "rubricas"
arquivo_rubricas = 'Rubricas - TESTE LINK.XLS'
dataframe_arquivo_rubricas = pd.read_excel(arquivo_rubricas, engine='xlrd', usecols=[2, 3])
dataframe_arquivo_rubricas.columns = dataframe_arquivo_rubricas.columns.str.lower().str.replace(' ', '')

# Abrir arquivo para escrita
with open("resultado.txt", "w") as arquivo_txt:
    # Escrever cabeçalho no TXT
    arquivo_txt.write("|46591297000108|LINK TESTE|\n")
    
    # Verificar similaridade entre nomes e dados
    for index_folha, row_folha in dataframe_folha_ponto.iterrows():
        maior_similaridade = 0
        nome_folha_ponto = row_folha['nome']  # Ajuste o nome da coluna para corresponder ao arquivo
        
        # Procurar o funcionário mais semelhante
        for index_func, row_funcionarios in dataframe_arquivo_funcionarios.iterrows():
            nome_arquivo_funcionarios = row_funcionarios.iloc[1]  # Nome do funcionário

            similaridade = jaccard_similarity(nome_folha_ponto, nome_arquivo_funcionarios)
            
            if similaridade > maior_similaridade:
                maior_similaridade = similaridade
                matricula_nome_mais_similar = str(row_funcionarios.iloc[0]).zfill(6)
                linha_mais_similar = row_folha

        if maior_similaridade >= 0.9:
            # Iterar pelas colunas da linha mais semelhante
            for nome_coluna, dado in linha_mais_similar.items():
                maior_similaridade_rubrica = 0
                rubrica_escolhida = None
                
                # Procurar rubrica mais semelhante
                for index_rubricas, row_rubricas in dataframe_arquivo_rubricas.iterrows():
                    nome_rubrica = row_rubricas.iloc[1]
                    similaridade_rubrica = jaccard_similarity(nome_coluna, nome_rubrica)
                    
                    if similaridade_rubrica > maior_similaridade_rubrica:
                        maior_similaridade_rubrica = similaridade_rubrica
                        rubrica_escolhida = nome_rubrica
                        rubrica = row_rubricas.iloc[0]

                # Verifica se a similaridade da rubrica é maior que 0.75
                if maior_similaridade_rubrica >= 0.75:
                    # Verifica se o dado está no formato hh:mm:ss
                    dado_str = str(dado)
                    if re.match(r'^\d{2}:\d{2}:\d{2}$', dado_str):  # Verifica se está no formato hh:mm:ss
                        dado_formatado = dado_str[:5].replace(':', '.')  # Transforma hh:mm:ss em hh.mm
                    else:
                        dado_formatado = dado_str.replace(':', '.')  # Substitui apenas os ":" se existirem
                    
                    # Escrever no arquivo apenas se ambas as similaridades forem satisfatórias
                    arquivo_txt.write(f"|{matricula_nome_mais_similar}|{int(rubrica)}|{dado_formatado}|\n")


print("Processamento concluído! Dados salvos em 'resultado.txt'.")