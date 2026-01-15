import re
import pdfplumber
import pandas as pd
import os

diretorio_raiz = input(r"Digite o caminho da pasta ou 'sair': ").strip()

dados_excel = []

while diretorio_raiz.lower() != 'sair':

    for dirpath, dirnames, filenames in os.walk(diretorio_raiz):
        for filename in filenames:

            if not filename.lower().endswith(".pdf"):
                continue

            caminho_completo = os.path.join(dirpath, filename)
            print(f"Processando arquivo: {caminho_completo}")

            with pdfplumber.open(caminho_completo) as pdf:
                texto = "\n".join(page.extract_text() or "" for page in pdf.pages)

            RazaoSocial =  re.findall(r'(?i)nome\s*/?\s*nome\s+empresarial\s+(.+)', texto)
            RazaoSocial_emit = RazaoSocial[0] if len(RazaoSocial) > 0 else None
            RazaoSocial_dest = RazaoSocial[1] if len(RazaoSocial) > 1 else None
            


            numeroNF =  re.search(r'Número RPS\s+(\d{1,7})\s+\1',texto)
            if numeroNF == None:
                numeroNF =  re.search(
                r"Numero\s+da\s+NFS-e:\s*([\d\.]+)",
                texto,
                re.IGNORECASE
            )
                
            numero_nf_final = None

            if numeroNF:
                numero_nf_final = numeroNF.group(1)
            else:
                # tenta pegar do nome do arquivo: data_NUMERO_CNPJ.pdf
                nome_match = re.search(r"\d{2}-\d{2}-\d{4}_(\d+)_", filename)
                if nome_match:
                    numero_nf_final = nome_match.group(1)


            data_emissao = re.search(
                r"Data\s+de\s+Emiss[aã]o\s+(\d{2}/\d{2}/\d{4})",
                texto,
                re.IGNORECASE
            )

           
            tipoNF = "NFs"

          
            padrao_totalNF = re.search(
                r"VALOR\s*DO\s*SERVIÇO\s*=\s*R\$\s*([\d\.,]+)",
                texto,
                re.IGNORECASE
            )

         
            cnpjs = re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
            cnpj_emit = cnpjs[0] if len(cnpjs) > 0 else None
            cnpj_dest = cnpjs[1] if len(cnpjs) > 1 else None

           
            discriminacao = re.search(
                r"Discrimina[cç][aã]o\s+dos\s+Servi[cç]os\s*(.*?)(?:C[oó]digo|VALOR)",
                texto,
                re.IGNORECASE | re.DOTALL
            )

            descricao = (
                discriminacao.group(1).strip()
                if discriminacao else None
            )

            dados_excel.append({
                "arquivo": filename,
                "numero_nf": numero_nf_final,
                "data_emissao": data_emissao.group(1) if data_emissao else None,
                "tipo_nf": tipoNF,
                "cnpj_emitente": cnpj_emit,
                'razao social emitente':RazaoSocial_emit,
                "cnpj_destinatario": cnpj_dest,
                'razao social destinario': RazaoSocial_dest,
                "valor_total_nf": padrao_totalNF.group(1) if padrao_totalNF else None,
                "descricao": descricao
            
            })

    if dados_excel:
        df = pd.DataFrame(dados_excel)
        df.to_excel("NFS.xlsx", index=False)
        print("Arquivo resultado.xlsx gerado com sucesso!")

    diretorio_raiz = input("Digite outro caminho ou 'sair': ").strip()

print("Encerrando...")
