import re
import pdfplumber
import pandas as pd
import os


diretorio_raiz = input(r"Digite o caminho da pasta ou 'sair': ").strip().lower()

dados_excel = []

while diretorio_raiz != 'sair':

    for dirpath, dirnames, filenames in os.walk(diretorio_raiz):
        for filename in filenames:

            if not filename.lower().endswith(".pdf"):
                continue

            caminho_completo = os.path.join(dirpath, filename)
            print(f"Processando arquivo: {caminho_completo}")

            with pdfplumber.open(caminho_completo) as pdf:
                texto = "\n".join(page.extract_text() or "" for page in pdf.pages)

            numeroNF = re.search(r"RECEBEDOR\s.*?([\d\.,]{6,})\s*", texto)
            data_emissao = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", texto)
            #tipoNF = re.search(r"\b([01])\s*-\s*(ENTRADA|SA[IÍ]DA)", texto, re.IGNORECASE)
            tipoNF ="SPED"

            razao_social_emitente = re.search(
                r"Identificação do emitente DANFE\s+(.*?)\s+DOCUMENTO AUXILIAR",
                texto,
                re.DOTALL
            )

            razao_social_destinatario = re.search(

                r"DESTINATARIO/REMETENTE.*?\nNOME/RAZÃO SOCIAL.*?\n([A-Z0-9 .&/-]+?)\s+\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}",
                texto,
                re.DOTALL
            )
            
           

            padrao_totalNF = re.search(
                r"TOTAL\s*DOS\s*PRODUTOS.*?([\d\.,]{6,})\s*",
                texto,
                re.IGNORECASE | re.DOTALL
            )

            cnpjs = re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
            cnpj_emit = cnpjs[0] if len(cnpjs) > 0 else None
            cnpj_dest = cnpjs[1] if len(cnpjs) > 1 else None

            regex_item = re.compile(
                r"(\d{6})\s+"
                r"([A-Z0-9 \/.-]+?)\s+"
                r"(\d+,\d{4})\s+"
                r"([\d\.,]+)\s+"
                r"([\d\.,]+)"
            )

            for cod, desc, qtde, unit, total in regex_item.findall(texto):
                dados_excel.append({
                    "arquivo": filename,
                    "numero_nf": numeroNF.group(1) if numeroNF else None,
                    "data_emissao": data_emissao.group(1) if data_emissao else None,
                    "tipo_nf": tipoNF,
                    "cnpj_emitente": cnpj_emit,
                    "razão social emitente": razao_social_emitente.group(1) if razao_social_emitente else None,
                    "cnpj_destinatario": cnpj_dest,
                    "razão social destinatario": razao_social_destinatario.group(1) if razao_social_destinatario else None,
                    "valor_total_nf": padrao_totalNF.group(1) if padrao_totalNF else None,
                    "descricao": desc.strip(),
                    "qtde": qtde,
                    "valor_unit": unit,
                    "total_item": total
                })

    # salva somente ao final
    if dados_excel:
        df = pd.DataFrame(dados_excel)
        df.to_excel("DANFE.xlsx", index=False)
        print("Arquivo resultado.xlsx gerado com sucesso!")

    diretorio_raiz = input("Digite outro caminho ou 'sair': ").strip().lower()

print("Encerrando...")
