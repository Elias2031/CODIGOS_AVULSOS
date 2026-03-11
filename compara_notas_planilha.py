import pandas as pd
import numpy as np
import os
import xml.etree.ElementTree as ET
import zipfile

# Função para calcular a similaridade de Jaccard
def jaccard_similarity(str1, str2):
    set1 = set(str1.lower().split())
    set2 = set(str(str2).lower().split())
    intersection = set1.intersection(set2)
    union = set1.union(set2)
    return len(intersection) / len(union)

diretorio_atual = os.getcwd()
arquivos_na_pasta = os.listdir(diretorio_atual)
excel = pd.read_excel('tabela de itens vendidos em 2023.xlsx')

# Verifica se a coluna 'CHAVE NF DE COMPRA' existe, caso contrário, cria
if 'CHAVE NF DE COMPRA' not in excel.columns:
    excel['CHAVE NF DE COMPRA'] = np.nan

arquivos_zip = [arquivo for arquivo in arquivos_na_pasta if arquivo.endswith('.zip')]

# List to keep track of items not found
itens_nao_encontrados = []

# Iterar sobre as linhas do DataFrame
for indice, ncm_tabela in enumerate(excel['ncm'].values):
    cest_tabela = excel.iloc[indice]['CEST']
    nome_item_tabela = excel.iloc[indice]['nome']
    
    if not pd.isna(cest_tabela):
        cest_tabela = round(float(cest_tabela))
        ncm_tabela = round(ncm_tabela)

        if len(str(cest_tabela)) < 7:
            cest_tabela = '0' + str(cest_tabela)

        if not arquivos_zip:
            print("Nenhum arquivo ZIP encontrado na pasta.")
        else:
            melhor_similaridade = 0
            melhor_item = None

            # LOOP PARA CADA ARQUIVO ZIP ENCONTRADO NA PASTA
            for nome_arquivo_zip in arquivos_zip:
                with zipfile.ZipFile(nome_arquivo_zip, 'r') as arquivo_zip:
                    lista_de_arquivos = arquivo_zip.namelist()

                    # LOOP PARA CADA ARQUIVO XML ENCONTRADO NO ARQUIVO ZIP 
                    for arquivo_nome in lista_de_arquivos:
                        if not arquivo_nome.startswith('CANCELADO'):
                            with arquivo_zip.open(arquivo_nome) as arquivo_xml:
                                tree = ET.parse(arquivo_xml)
                                root = tree.getroot()
                                items = root.findall('.//{http://www.portalfiscal.inf.br/nfe}det')
                                chave_nota = root.find('.//{http://www.portalfiscal.inf.br/nfe}chNFe').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}chNFe') is not None else 0
                                
                                for item in items:
                                    nome_item_nota = item.find('.//{http://www.portalfiscal.inf.br/nfe}xProd').text if item.find('.//{http://www.portalfiscal.inf.br/nfe}xProd') is not None else 0
                                    ncm_item_nota = item.find('.//{http://www.portalfiscal.inf.br/nfe}NCM').text if item.find('.//{http://www.portalfiscal.inf.br/nfe}NCM') is not None else 0
                                    
                                    
                                    
                                    if str(ncm_item_nota).replace(' ', '') == str(ncm_tabela).replace(' ', ''):
                                        emit = root.find('.//{http://www.portalfiscal.inf.br/nfe}emit')    
                                        icms = item.find('.//{http://www.portalfiscal.inf.br/nfe}ICMS')
                                            
                                        # Verificação segura de CSOSN e CST
                                        csosn = icms.find('.//{http://www.portalfiscal.inf.br/nfe}CSOSN') if icms is not None else None
                                        cst = icms.find('.//{http://www.portalfiscal.inf.br/nfe}CST') if icms is not None else None
                                            
                                        if csosn is not None:
                                            situacao_tributaria = [csosn.text, 'CSOSN']
                                        elif cst is not None:
                                            situacao_tributaria = [cst.text, 'CST']
                                        else:
                                            situacao_tributaria = ['Desconhecido', 'Desconhecido']
                                        
                                        melhor_item = {
                                                'Chave Nota': chave_nota,
                                                'NCM Nota': ncm_item_nota,
                                                'Nome Item Nota': nome_item_nota,
                                                'Modelo Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}mod').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}mod') is not None else 0,
                                                'Série Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}serie').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}serie') is not None else 0,
                                                'Número Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF') is not None else 0,
                                                'Valor Total Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}vNF').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}vNF') is not None else 0,
                                                'Data Emissão': root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi') is not None else 0,
                                                'Situação': root.find('.//{http://www.portalfiscal.inf.br/nfe}cStat').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}cStat') is not None else 0,
                                                'CEST do Item': item.find('.//{http://www.portalfiscal.inf.br/nfe}CEST').text if item.find('.//{http://www.portalfiscal.inf.br/nfe}CEST') is not None else 0,
                                                'Numero do Item' : item.get('nItem'),
                                                'Cfop do item' : item.find('.//{http://www.portalfiscal.inf.br/nfe}CFOP').text if item.find('.//{http://www.portalfiscal.inf.br/nfe}CFOP') is not None else 0,
                                                'Nome do participante' : emit.find('.//{http://www.portalfiscal.inf.br/nfe}xNome').text,
                                                'UF do participante' : emit.find('.//{http://www.portalfiscal.inf.br/nfe}UF').text,
                                                'nome_excel' : nome_item_tabela
                                            }
                                        
                                        similaridade = jaccard_similarity(nome_item_tabela, nome_item_nota)
                                            
                                        if similaridade > melhor_similaridade:
                                            melhor_similaridade = similaridade
                                            melhor_item = {
                                                'Chave Nota': chave_nota,
                                                'NCM Nota': ncm_item_nota,
                                                'Nome Item Nota': nome_item_nota,
                                                'Modelo Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}mod').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}mod') is not None else 0,
                                                'Série Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}serie').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}serie') is not None else 0,
                                                'Número Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF') is not None else 0,
                                                'Valor Total Nota': root.find('.//{http://www.portalfiscal.inf.br/nfe}vNF').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}vNF') is not None else 0,
                                                'Data Emissão': root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi') is not None else 0,
                                                'Situação': root.find('.//{http://www.portalfiscal.inf.br/nfe}cStat').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}cStat') is not None else 0,
                                                'CEST do Item': item.find('.//{http://www.portalfiscal.inf.br/nfe}CEST').text if item.find('.//{http://www.portalfiscal.inf.br/nfe}CEST') is not None else 0,
                                                'Numero do Item' : item.get('nItem'),
                                                'Cfop do item' : item.find('.//{http://www.portalfiscal.inf.br/nfe}CFOP').text if item.find('.//{http://www.portalfiscal.inf.br/nfe}CFOP') is not None else 0,
                                                'Nome do participante' : emit.find('.//{http://www.portalfiscal.inf.br/nfe}xNome').text,
                                                'UF do participante' : emit.find('.//{http://www.portalfiscal.inf.br/nfe}UF').text
                                            }

            if melhor_item:
                excel.at[indice, 'CHAVE NF DE COMPRA'] = melhor_item['Chave Nota']
                excel.at[indice, 'Nº ITEM'] = melhor_item['Numero do Item']
                excel.at[indice, 'CFOP NA COMPRA'] = melhor_item['Cfop do item']
                excel.at[indice, 'CSOSN/CST'] = f'{situacao_tributaria[1]} = {situacao_tributaria[0]}'
                excel.at[indice, 'FORNECEDOR'] = f"{melhor_item['Nome do participante']} - {melhor_item['UF do participante']}"
                print(situacao_tributaria)
            else:
                itens_nao_encontrados.append(nome_item_tabela)

# Salva o DataFrame atualizado de volta ao arquivo Excel
excel.to_excel('tabela de itens vendidos em 2023_atualizada.xlsx', index=False)

# Exibe os itens que não foram encontrados
print("Itens não encontrados:")
for item in itens_nao_encontrados:
    print(item)
