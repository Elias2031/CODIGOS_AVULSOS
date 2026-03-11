import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

def extrair_informacoes_nota(root):
    nota = {}
    nota['modelo'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}mod').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}mod') is not None else 0
    nota['serie'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}serie').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}serie') is not None else 0
    nota['numero'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF') is not None else 0
    nota['chave'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}chNFe').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}chNFe') is not None else 0
    nota['valor_total'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}vNF').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}vNF') is not None else 0
    nota['valor_desc'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}vDesc').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}vDesc') is not None else 0
    nota['valor_icms'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}vICMS').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}vICMS') is not None else 0
    nota['data_emissao'] = root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi').text if root.find('.//{http://www.portalfiscal.inf.br/nfe}dhEmi') is not None else 0
    nota['data_emissao_formatada'] = datetime.strptime(nota['data_emissao'][:10], "%Y-%m-%d").strftime("%d%m%Y")
    return nota

def verificar_chave_em_txt(chave, arquivo_txt):
    with open(arquivo_txt, 'r',) as chaves_txt:
        for linha in chaves_txt:
            if chave in linha:
                return True
    return False

diretorio_atual = os.getcwd()
arquivos_na_pasta = os.listdir(diretorio_atual)
arquivos_zip = [arquivo for arquivo in arquivos_na_pasta if arquivo.endswith('.zip')]
arquivo_excel = [arquivo for arquivo in arquivos_na_pasta if arquivo.endswith('.xlsx')]
arquivo_txt = [arquivo for arquivo in arquivos_na_pasta if arquivo.endswith('.txt')]

# LOOP PARA CADA ARQUIVO ZIP ENCONTRADO NA PASTA
for nome_arquivo_zip in arquivos_zip:
    with zipfile.ZipFile(nome_arquivo_zip, 'r') as arquivo_zip:
        lista_de_arquivos = arquivo_zip.namelist()
          
        # LOOP PARA CADA ARQUIVO XML ENCONTRADO NO ARQUIVO ZIP 
        for arquivo_nome in lista_de_arquivos:
            if not arquivo_nome.startswith('CANCELADO'):

                # ABRE O ARQUIVO XML E EXTRAI VALORES DAS VARIAVEIS REFERENTE A NOTA FISCAL
                with arquivo_zip.open(arquivo_nome) as arquivo_xml:
                    tree = ET.parse(arquivo_xml)
                    root = tree.getroot()
                    nota_fiscal = extrair_informacoes_nota(root)

                    if not verificar_chave_em_txt(nota_fiscal['chave'], arquivo_txt[0]):
                        print("Chave não encontrada no arquivo TXT:", nota_fiscal['chave'])
                    else: 
                        print("Chave encontrada no arquivo TXT:", nota_fiscal['chave'])