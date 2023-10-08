import xmltodict
import os
import json
import pandas as pd

def pegar_arquivos(arquivo, valores):
    with open(f'nfs/{arquivo}', 'rb') as arquivo_xml:
        dicionario_arquivo = xmltodict.parse(arquivo_xml)

        try:
            if 'NFe' in dicionario_arquivo:
                informacoes_nfe = dicionario_arquivo['NFe']['infNFe']
            else: informacoes_nfe = dicionario_arquivo['nfeProc']['NFe']['infNFe']
            numero_nota = informacoes_nfe['@Id']
            empresa_emissora = informacoes_nfe['emit']['xNome']
            nome_cliente = informacoes_nfe['dest']['xNome']
            endereco = []
            for i in informacoes_nfe['dest']['enderDest']:
                endereco.append((informacoes_nfe['dest']['enderDest'][i]))
            if 'vol' in informacoes_nfe['transp']:
                peso = informacoes_nfe['transp']['vol']['pesoB']
            else:
                peso = 'Não Informado'

            valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])

        except Exception as e:
            print(e)
            print(json.dumps(dicionario_arquivo, indent=4))   #Mostrando em forma de xml

arquivos = os.listdir("nfs")

colunas = ['Numero_nota', 'Empresa_emissora', 'Cliente', 'Endereço', 'Peso']
valores = []


for arquivo in arquivos:
    pegar_arquivos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('Notas_Fiscais.xlsx', index=False)