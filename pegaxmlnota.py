import os
import xmltodict
import pandas as pd



def pega_nota(nome_arquivo, valores):
    print(f"{nome_arquivo} acessada com sucesso")

    with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
        dic_nota = xmltodict.parse(arquivo_xml)
        

    try:
    
        if "NFe" in dic_nota:
            infos_nf = dic_nota["NFe"]["infNFe"]
        else:
            infos_nf = dic_nota["nfeProc"]["NFe"]["infNFe"]

        
        numero_nota = infos_nf["ide"]["nNF"]
        cnpj = infos_nf["emit"]["CNPJ"]

    
        itens = infos_nf["det"]
        if isinstance(itens, dict):
            itens = [itens] 

        for item in itens:
            prod = item["prod"]

            descricao = prod.get("xProd", "")
            total_por_item = prod.get("vProd", "")
            quantidade = prod.get("qCom", "")
            valor_unitario = prod.get("vUnCom", "")
            valor_total_nf = infos_nf["total"]["ICMSTot"]["vNF"]

            valores.append([
                numero_nota,
                cnpj,
                descricao,
                total_por_item,
                quantidade,
                valor_unitario,
                valor_total_nf
            ])

    except Exception as e:
        print(f"Erro encontrado em: {e}, no arquivo {nome_arquivo}")



lista_notas = os.listdir("nfs")

colunas = [
    'numero_nota',
    'cnpj',
    'descricao',
    'total_por_item',
    'quantidade',
    'valor_unitario',
    'valor_total_nf'
]

valores = []

for nota in lista_notas:
    pega_nota(nota, valores)


tabela = pd.DataFrame(columns=colunas, data=valores)


tabela.to_excel('NotasFiscais.xlsx', index=False)
print("Planilha 'NotasFiscais.xlsx' criada com sucesso!")

