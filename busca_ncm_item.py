import xml.etree.ElementTree as ET
import pandas as pd
import datetime
import zipfile
import os 
import re 

global razao_social
razao_social = ''

diretorio_atual = os.getcwd()
arquivos_na_pasta_notas_fiscais = os.listdir(os.path.join(diretorio_atual, 'notas_fiscais'))
arquivos_zip = [arquivo for arquivo in arquivos_na_pasta_notas_fiscais if arquivo.endswith('.zip')]

def extrair_formatar_data_nf(root):
    # EXTRAIR E FORMATAR DATA DA NOTA
    data_emissao_nota = root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi')
    data_emissao_nota = data_emissao_nota.text if data_emissao_nota is not None else '0'    
    if data_emissao_nota != '0':
        # Converter a string para um objeto datetime
        data_objeto = datetime.datetime.strptime(data_emissao_nota[:10], "%Y-%m-%d")
        
        # Formatar para exibição no Excel
        data_emissao_formatada_nota_para_excel = data_objeto.strftime("%d/%m/%Y")
        
        # Extrair mês e ano
        mes = data_objeto.month
        mes = str(mes).zfill(2)
        ano = data_objeto.year
        
    else:
        data_emissao_formatada_nota_para_excel = '0'
        mes = '0'
        ano = '0'
    mes_ano = f'{ano}-{mes}'    
    dados_datas = {
        'data_emissao_formatada_nota_para_excel' : data_emissao_formatada_nota_para_excel,
        'mes_ano' : mes_ano
    }    
        
    return dados_datas

def extrair_valores_notas(root, caminho, valor_padrao):
    """Extrai um valor de um elemento XML, retornando um padrão se não encontrado."""
    elemento = root.find(f'.//{{http://www.portalfiscal.inf.br/nfe}}{caminho}')
    return elemento.text if elemento is not None else valor_padrao

def jaccard_similarity(str1, str2):
    # Convert to lowercase and remove punctuation
    str1 = re.sub(r'[^\w\s]', '', str1.lower())
    str2 = re.sub(r'[^\w\s]', '', str2.lower())
    
    set1 = set(str1.split())
    set2 = set(str2.split())
    
    intersection = set1.intersection(set2)
    union = set1.union(set2)
    
    return len(intersection) / len(union)


def comecar_apurar():
    global dicionario_csts_corretos 
    dicionario_csts_corretos = {}
    notas_processadas = set()
    contador_notas_dest = 0

    # Processamento de notas de ENTRADA (quando o zip possui 'dest' no nome)
    for nome_arquivo_zip in arquivos_zip:
        if 'dest' in nome_arquivo_zip.lower() or 'Destinatário' in nome_arquivo_zip.lower():
            
            with zipfile.ZipFile(os.path.join('notas_fiscais', nome_arquivo_zip), 'r') as arquivo_zip:
                lista_de_arquivos = [f for f in arquivo_zip.namelist() if not f.startswith('CANCELADO')]
                for arquivo_nome in lista_de_arquivos:
                    contador_notas_dest +=1
                    with arquivo_zip.open(arquivo_nome) as arquivo_xml:
                        tree = ET.parse(arquivo_xml)
                        root = tree.getroot()
                        chave_nota = extrair_valores_notas(root, 'chNFe', '')
                        finalidade_nfe = extrair_valores_notas(root, 'finNFe', '1')
                        # Verifica se a nota já foi processada
                        
                        # Verifica se a nota é uma devolução
                        if finalidade_nfe != '1':
                            # print(f"Nota fiscal {chave_nota} é uma NOTA FISCAL DE DEVOLUÇÃO. Ignorando...")
                            continue
                        
                        if chave_nota in notas_processadas:
                            # print(f"Nota {chave_nota} já foi processada. Pulando arquivo {arquivo_nome}.")
                            continue
                        else:
                            notas_processadas.add(chave_nota)
                        
                        # Para notas de entrada, extraímos o CNPJ do destinatário
                        dados_datas = extrair_formatar_data_nf(root)
                        mes_ano = dados_datas['mes_ano']
                        
                        items = root.findall('.//{http://www.portalfiscal.inf.br/nfe}det')
                        for item in items:
                            nome_item_nota = extrair_valores_notas(item, 'xProd', '')
                            ncm_item_nota = extrair_valores_notas(item, 'NCM', 0)
                            
                            nome_para_comparar = 'SOLUCAO DESENGRAXANTE 0,900 ML ROYAL'
                            
                            similaridade = jaccard_similarity(nome_item_nota, nome_para_comparar)
                            
                            if ncm_item_nota == '29012900' or similaridade >= 0.2:
                                
                                teste = extrair_valores_notas(item, 'CSOSN', 0)
                                if teste == 0:
                                    pis_teste_obj = item.find('.//{http://www.portalfiscal.inf.br/nfe}PIS')
                                    pis_teste = extrair_valores_notas(pis_teste_obj, 'CST', 0)
                                    # if pis_teste != 0:
                                    print(mes_ano, chave_nota, pis_teste)

comecar_apurar()