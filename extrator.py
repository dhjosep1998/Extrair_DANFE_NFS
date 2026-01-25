import re
import os
import pdfplumber
import pandas as pd


def detectar_tipo_nf(texto: str):
    t = texto.upper()

    if "DANFE" in t or "DOCUMENTO AUXILIAR DA NOTA FISCAL" in t:
        return "DANFE"

    if "NFS-E" in t or "DANFSE" in t or "DOCUMENTO AUXILIAR DA NFS-E" in t:
        return "NFS"

    return None



def extrair_numero_nfs(texto, filename):
    padroes = [
        r"Numero\s+da\s+NFS-e[:\s]*([\d\.]+)",
        r"NFS-e\s*N[ºo]?\s*[:\-]?\s*([\d\.]+)",
        r"Número\s+RPS\s+(\d{1,7})\s+\1",
        r"RPS[:\s]+(\d+)"
    ]
    for p in padroes:
        m = re.search(p, texto, re.IGNORECASE)
        if m:
            return m.group(1)

    # fallback: tenta pegar do nome do arquivo
    m = re.search(r"\d{2}-\d{2}-\d{4}_(\d+)_", filename)
    return m.group(1) if m else None


def processar_notas(pasta, progresso_callback=None):
    dados = []

    pdfs = [
        os.path.join(dp, f)
        for dp, _, fs in os.walk(pasta)
        for f in fs if f.lower().endswith(".pdf")
    ]

    total = len(pdfs)
    if total == 0:
        return False

    for idx, caminho in enumerate(pdfs, start=1):
        arquivo = os.path.basename(caminho)

        with pdfplumber.open(caminho) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)

        tipo = detectar_tipo_nf(texto)

        cnpjs = re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
        cnpj_emit = cnpjs[0] if len(cnpjs) > 0 else None
        cnpj_dest = cnpjs[1] if len(cnpjs) > 1 else None

        # ================= DANFE =================
        if tipo == "DANFE":
            numero_nf = re.search(r"RECEBEDOR\s.*?([\d\.,]{6,})", texto)
            data = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", texto)

            razao_emit = re.search(
                r"Identificação do emitente DANFE\s+(.*?)\s+DOCUMENTO AUXILIAR",
                texto,
                re.DOTALL
            )

            razao_dest = re.search(
                r"DESTINATARIO/REMETENTE.*?\nNOME/RAZÃO SOCIAL.*?\n([A-Z0-9 .&/-]+?)\s+\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}",
                texto,
                re.DOTALL
            )

            total_nf = re.search(
                r"TOTAL\s*DOS\s*PRODUTOS.*?([\d\.,]+)",
                texto,
                re.IGNORECASE | re.DOTALL
            )

            regex_item = re.compile(
                r"(\d{6})\s+"
                r"([A-Z0-9 \/.-]+?)\s+"
                r"(\d+,\d{4})\s+"
                r"([\d\.,]+)\s+"
                r"([\d\.,]+)"
            )

            itens = regex_item.findall(texto)
            if not itens:
                itens = [(None, None, None, None, None)]

            for _, desc, qtde, unit, total_item in itens:
                dados.append({
                    "arquivo": arquivo,
                    "tipo_nf": "DANFE",
                    "numero_nf": numero_nf.group(1) if numero_nf else None,
                    "data_emissao": data.group(1) if data else None,
                    "cnpj_emitente": cnpj_emit,
                    "razao_social_emitente": razao_emit.group(1).strip() if razao_emit else None,
                    "cnpj_destinatario": cnpj_dest,
                    "razao_social_destinatario": razao_dest.group(1).strip() if razao_dest else None,
                    "descricao": desc,
                    "qtde": qtde,
                    "valor_unit": unit,
                    "total_item": total_item,
                    "valor_total_nf": total_nf.group(1) if total_nf else None
                })

        # ================= NFS =================
        elif tipo == "NFS":
            numero_nf = re.search(
                r"N[uú]mero\s+da\s+NFS-e\s+(\d+)",
                texto,
                re.IGNORECASE
            )

            data = re.search(
                r"Data\s+e\s+Hora\s+da\s+emiss[aã]o\s+da\s+NFS-e\s+(\d{2}/\d{2}/\d{4})",
                texto,
                re.IGNORECASE
            )

            valor_total = re.search(
                r"Valor\s+do\s+Servi[cç]o\s+R\$\s*([\d\.,]+)",
                texto,
                re.IGNORECASE
            )

            descricao = re.search(
                r"Descri[cç][aã]o\s+do\s+Servi[cç]o\s+(.*?)(?:TRIBUTA[CÇ][AÃ]O|VALOR\s+DO\s+SERVIÇO)",
                texto,
                re.IGNORECASE | re.DOTALL
            )

            emitente = re.search(
                r"EMITENTE\s+DA\s+NFS-e.*?Nome\s*/\s*Nome\s+Empresarial\s+(.*?)\n",
                texto,
                re.IGNORECASE | re.DOTALL
            )

            tomador = re.search(
                r"TOMADOR\s+DO\s+SERVIÇO.*?Nome\s*/\s*Nome\s+Empresarial\s+(.*?)\n",
                texto,
                re.IGNORECASE | re.DOTALL
            )

            dados.append({
                "arquivo": arquivo,
                "tipo_nf": "NFS",
                "numero_nf": numero_nf.group(1) if numero_nf else None,
                "data_emissao": data.group(1) if data else None,
                "cnpj_emitente": cnpj_emit,
                "razao_social_emitente": emitente.group(1).strip() if emitente else None,
                "cnpj_destinatario": cnpj_dest,
                "razao_social_destinatario": tomador.group(1).strip() if tomador else None,
                "descricao": descricao.group(1).strip() if descricao else None,
                "qtde": None,
                "valor_unit": None,
                "total_item": None,
                "valor_total_nf": valor_total.group(1) if valor_total else None
            })

        if progresso_callback:
            progresso_callback(idx / total)

    pd.DataFrame(dados).to_excel("NOTAS_FISCAIS.xlsx", index=False)
    return True
