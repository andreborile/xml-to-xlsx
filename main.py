import xmltodict
import os
import pandas as pd

def pegar_info(nome_arquivo, valores):
    with open(f'xml_to_excel/nfs/{nome_arquivo}', 'rb') as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        numero_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf["emit"]["xNome"]
        nome_cliente =infos_nf["dest"]["xNome"]
        endereco = infos_nf["dest"]["enderDest"]
        if "vol" in infos_nf["transp"]:
            peso = infos_nf["transp"]["vol"]["pesoB"]
        else: 
            peso = "não informado"
        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])

lista_arquivos = os.listdir("xml_to_excel/nfs")

colunas = ['numero_nota', 'empresa_emissora', 'nome_cliente', 'endereco', 'peso']
valores = []

for arquivo in lista_arquivos:
    pegar_info(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFiscais.xlsx', index=False)