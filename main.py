import xmltodict
import os
import pandas as pd

def pegar_infos(nome_arquivo, valores):
    endereco = {}
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)


        infos_nf = dic_arquivo["NFe"]['infNFe'] if "NFe" in dic_arquivo else dic_arquivo['nfeProc']["NFe"]['infNFe']
        numero_nota             = infos_nf["@Id"]
        empresa_emissora        = infos_nf['emit'].get('xNome',"")
        nome_cliente            = infos_nf["dest"].get('xNome',"")
        endereco['lgr']         = infos_nf["dest"]["enderDest"].get("xLgr","")
        endereco['nro']         = infos_nf["dest"]["enderDest"].get("nro","")
        endereco['Complemento'] = infos_nf["dest"]["enderDest"].get("xCpl","")
        endereco['Bairro']      = infos_nf["dest"]["enderDest"].get("xBairro","")
        endereco['UF']          = infos_nf["dest"]["enderDest"].get("UF","")
        endereco['CEP']         = infos_nf["dest"]["enderDest"].get("CEP","")
        endereco['pais']        = infos_nf["dest"]["enderDest"].get("xPais","")
        endereco['contato']     = infos_nf["dest"]["enderDest"].get("fone","")
        peso = infos_nf["transp"]["vol"]["pesoB"] if "vol" in infos_nf["transp"] else "Não informado"

        valores.append(
            [numero_nota,
             empresa_emissora,
             nome_cliente,
             endereco['lgr'],
             endereco['nro'],
             endereco['Complemento'],
             endereco['Bairro'],
             endereco['UF'],
             endereco['CEP'],
             endereco['pais'],
             endereco['contato'],
             peso])

colunas = ["numero_nota",
           "empresa_emissora",
           "nome_cliente",
           "lgr",
           "nro",
           "cpl",
           "Bairro",
           "UF",
           "CEP",
           "Pais",
           "Contato",
           "peso"]
valores = []
lista_arquivos = os.listdir("nfs")
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)
print("Arquivo Gerado")