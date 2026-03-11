# Importa a biblioteca pandas com o alias pd
import pandas as pd
import os
from decimal import Decimal

def soma_c190(nome_arquivo):
    def processar_linhas_c190(linhas):
        novas_linhas = []
        c190_acumulado = None
        contador_linhas_excluidas = 0
        dentro_c100 = False  # Flag para identificar quando estamos dentro de um bloco C100

        for linha in linhas:
            campos = linha.split('|')

            if campos[1] == 'C100':
                # Quando encontrar um novo C100, salva o que acumulou até agora
                if c190_acumulado:
                    novas_linhas.append('|'.join(c190_acumulado))
                    c190_acumulado = None
                novas_linhas.append(linha)
                dentro_c100 = True

            elif campos[1] == 'C190' and dentro_c100:
                if campos[3] == '5405':
                    if c190_acumulado is None:
                        c190_acumulado = campos.copy()
                    else:
                        # Acumula os valores dos campos 4 e 5
                        c190_acumulado[4] = str(Decimal(c190_acumulado[4].replace(',', '.')) + Decimal(campos[4].replace(',', '.'))).replace('.', ',')
                        c190_acumulado[5] = str(Decimal(c190_acumulado[5].replace(',', '.')) + Decimal(campos[5].replace(',', '.'))).replace('.', ',')

                        contador_linhas_excluidas += 1  # Só conta exclusões a partir do segundo
                else:
                    novas_linhas.append(linha)

            else:
                # Se mudar para outro bloco fora de C100
                if c190_acumulado:
                    novas_linhas.append('|'.join(c190_acumulado))
                    c190_acumulado = None
                novas_linhas.append(linha)
                dentro_c100 = False

        # No final, se sobrou algum C190 acumulado, adiciona
        if c190_acumulado:
            novas_linhas.append('|'.join(c190_acumulado))

        return novas_linhas, contador_linhas_excluidas

    # Linhas de entrada
    linhas = []

    with open(nome_arquivo, 'r', encoding='utf-8') as file:
        for linha in file :
            linhas.append(linha.strip())

    # Processa as linhas C190 e substitui as linhas originais
    linhas_processadas, contador_linhas_excluidas = processar_linhas_c190(linhas)

    with open(nome_arquivo, 'w', encoding='utf-8') as file:
        for linha in linhas_processadas:
            if linha.startswith('|C990|'):
                # Extrai a linha C990
                c990 = [linha.strip().split('|')]
                # Altera a linha C990 de acordo com o contador
                c990[0][2] = str(int(c990[0][2]) - contador_linhas_excluidas)
                # Escreve a linha C990 modificada no arquivo modificado
                file.write('|'.join(c990[0]) + '\n')
            elif linha.startswith('|9999|'):
                # Extrai a linha 9999 referente ao total de linhas no arquivo
                c9999 = [linha.strip().split('|')]
                # Altera a linha 9999 de acordo com o contador
                c9999[0][2] = str(int(c9999[0][2]) - contador_linhas_excluidas)
                # Escreve a linha 9999 modificada no arquivo modificado
                file.write('|'.join(c9999[0]) + '\n')
            elif linha.startswith('|9900|C190|'):
                # Extrai a linha C9900 referente aos arquivos C190
                c9900 = [linha.strip().split('|')]
                # Altera a qtd de linhas C190 de acordo com o contador
                c9900[0][3] = str(int(c9900[0][3]) - contador_linhas_excluidas)
                # Escreve a linha c9900 modificada no arquivo modificado
                file.write('|'.join(c9900[0]) + '\n')
            else : file.write(linha + '\n')

arquivos = os.listdir()

arquivos_txt = [arquivo for arquivo in arquivos if arquivo.endswith('.txt')]

# Define os nomes dos arquivos de entrada e saída
arquivo_sped = arquivos_txt[0]
novo_nome = "MODIFICADO_" + arquivo_sped
arquivo_sped_modificado = novo_nome

# Solicita ao usuário o valor a ser modificado
valor_a_modificar = float(input('Qual valor deseja adicionar ao 5405?'))

# Inicializa listas e variáveis
nota = []
escreve= True
array_notas = []
dentro_da_nota = False
contador = 0
conta_5405 = 0

with open(arquivo_sped, 'r', encoding='utf-8') as file:
    # Abre o arquivo modificado para escrita
    with open(arquivo_sped_modificado, 'w', encoding='utf-8') as file_modificado:
        # Itera sobre cada linha do arquivo original
        for line in file:
            # Verifica se a linha começa com '|C990|', indicando o final do bloco C
            if line.startswith('|C990|'):
                # Extrai a linha C990
                c990 = [line.strip().split('|')]
                # Altera a linha C990 de acordo com o contador
                c990[0][2] = str(int(c990[0][2]) + contador)
                # Escreve a linha C990 modificada no arquivo modificado
                file_modificado.write('|'.join(c990[0]) + '\n')
                # Adiciona a nota atual à lista de notas
                array_notas.append(nota)
                # Define que já não está mais dentro de uma nota
                dentro_da_nota = False
                # Não escreve a linha original no arquivo modificado
                escreve=False

            if line.startswith('|9900|C190|'):
                # Extrai a linha C9900 referente aos arquivos C190
                c9900 = [line.strip().split('|')]
                # Altera a qtd de linhas C190 de acordo com o contador
                c9900[0][3] = str(int(c9900[0][3]) + contador)
                # Escreve a linha c9900 modificada no arquivo modificado
                file_modificado.write('|'.join(c9900[0]) + '\n')
                # Não escreve a linha original no arquivo modificado
                escreve=False

            if line.startswith('|9999|'):
                # Extrai a linha 9999 referente ao total de linhas no arquivo
                c9999 = [line.strip().split('|')]
                # Altera a linha 9999 de acordo com o contador
                c9999[0][2] = str(int(c9999[0][2]) + contador)
                # Escreve a linha 9999 modificada no arquivo modificado
                file_modificado.write('|'.join(c9999[0]) + '\n')
                # Não escreve a linha original no arquivo no arquivo modificado
                escreve=False

            # Verifica se a linha começa com '|C100|', indicando o início de uma nota fiscal
            if line.startswith('|C100|'):
                # Se já estiver dentro de uma nota, adiciona a nota anterior à lista
                if dentro_da_nota:
                    array_notas.append(nota)
                # Inicializa a nota com a linha atual
                nota = [line.strip().split('|')]
                # Define que estamos dentro da nota
                dentro_da_nota = True
                conta_5405 = 0

            # Se estiver dentro de uma nota, adiciona a linha à nota atual
            elif dentro_da_nota:
                nota.append(line.strip().split('|'))

            # Verifica se a linha começa com '|C190|', indicando a apuração de um documento fiscal
            if line.startswith('|C190|'):
                # Extrai a linha
                linha = line.strip().split('|')
                # Extrai o valor da nota
                valor_cfop = float(linha[5].replace(',', '.'))

                # Verifica condições para modificar o valor
                if nota[0][2] == '1':
                    if linha[3] == '5102' or linha[3] == '6102' or linha[3] == '5101':
                        if valor_a_modificar > 0 and valor_a_modificar >= valor_cfop:
                            # Modifica valores na linha C190
                            linha[3] = '5405'
                            linha[2] = '500'
                            valor_a_modificar = valor_a_modificar - valor_cfop

                        elif valor_a_modificar > 0 and valor_a_modificar < valor_cfop:
                            # Retira o valor_a_modificar da linha atual
                            linha[5] = str(round(float(linha[5].replace(',','.')) - valor_a_modificar, 2)).replace('.', ',')

                            # Cria uma nova linha com o valor_a_modificar
                            nova_linha_valor_modificar = '|C190|500|5405|0,00|' + str(round(float(valor_a_modificar), 2)).replace('.',',') + '|0,00|0,00|0,00|0,00|0,00|0,00||\n'
                            file_modificado.write(nova_linha_valor_modificar)
                            valor_a_modificar = 0
                            contador += 1
                            # # # Conclui as alterações quando o valor desejado é atingido
                        if valor_a_modificar <= 0:
                            print('Alterações concluídas com sucesso')

                # Escreve a linha modificada ou não no arquivo modificado
                nova_linha = '|'.join(linha) + '\n'
                file_modificado.write(nova_linha)
            else:
                if escreve == True:
                    # Se não for uma linha C190, escreve a linha original no arquivo modificado
                    file_modificado.write(line)
                escreve= True
soma_c190(novo_nome)