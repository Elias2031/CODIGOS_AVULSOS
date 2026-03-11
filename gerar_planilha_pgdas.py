import re
from pathlib import Path

import pdfplumber
import pandas as pd


# =====================
# UTILITÁRIOS
# =====================

def br_to_float(s: str) -> float:
    if not s:
        return 0.0
    s = s.replace("R$", "").strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def extrair_texto_pdf(pdf_path: Path) -> str:
    texto = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto.append(page.extract_text() or "")
    return "\n".join(texto)


def extrair_periodo(texto: str) -> str:
    m = re.search(
        r"Per[ií]odo de Apura[cç][aã]o:\s*([0-9]{2}/[0-9]{2}/[0-9]{4}\s*a\s*[0-9]{2}/[0-9]{2}/[0-9]{4})",
        texto
    )
    return m.group(1) if m else ""


def extrair_rpa_caixa(texto: str) -> float:
    m = re.search(
        r"Receita Bruta do PA\s*\(RPA\)\s*-\s*Caixa\s*([0-9\.\,]+)",
        texto
    )
    return br_to_float(m.group(1)) if m else 0.0


# =====================
# EXTRAÇÃO POR ATIVIDADE (COM TODOS OS IMPOSTOS)
# =====================

IMPOSTOS_COLS = ["irpj", "csll", "cofins", "pis_pasep", "inss_cpp", "icms", "ipi", "iss", "total"]

def extrair_blocos_atividade(texto: str) -> list[dict]:
    texto = re.sub(r"\r", "", texto)

    pattern = (
        r"Valor do D[eé]bito por Tributo para a Atividade\s*\(R\$\)\s*:\s*"
        r"(?P<desc>.*?)"
        r"Receita Bruta Informada:\s*R\$\s*(?P<receita>[0-9\.\,]+)\s*"
        r"IRPJ\s+CSLL\s+COFINS\s+PIS/Pasep\s+INSS/CPP\s+ICMS\s+IPI\s+ISS\s+Total\s*"
        r"(?P<valores>[0-9\.\,\s]+)"
    )

    blocos = []

    for m in re.finditer(pattern, texto, flags=re.DOTALL | re.IGNORECASE):
        desc = re.sub(r"\s+", " ", m.group("desc")).strip()
        receita = br_to_float(m.group("receita"))

        # captura apenas tokens numéricos (ex.: "3.469,33")
        tokens = re.sub(r"\s+", " ", m.group("valores")).strip().split(" ")
        tokens = [t for t in tokens if re.match(r"^[0-9\.\,]+$", t)]

        # precisamos de 9 valores (IRPJ..Total)
        valores = [0.0] * 9
        if len(tokens) >= 9:
            for i in range(9):
                valores[i] = br_to_float(tokens[i])

        bloco = {
            "descricao_atividade": desc,
            "receita_atividade": receita,
        }

        # injeta cada imposto no dict
        for col, val in zip(IMPOSTOS_COLS, valores):
            bloco[col] = val

        blocos.append(bloco)

    return blocos


def classificar_por_icms_iss(icms: float, iss: float) -> str:
    if icms > 0 and iss > 0:
        return "SERVIÇO + VENDA"
    if icms > 0:
        return "SÓ VENDA"
    if iss > 0:
        return "SÓ SERVIÇO"
    return "INCONSISTENTE"


# =====================
# PROCESSAMENTO EM LOTE
# =====================

def processar_pasta_pgdas(pasta_pgdas: str, saida_xlsx="analise_pgdas_consolidada.xlsx"):
    pasta = Path(pasta_pgdas)

    linhas_detalhadas = []
    linhas_resumo = []

    for pdf in pasta.glob("*.pdf"):
        print(f"Processando: {pdf.name}")

        texto = extrair_texto_pdf(pdf)
        periodo = extrair_periodo(texto)
        rpa_caixa = extrair_rpa_caixa(texto)

        blocos = extrair_blocos_atividade(texto)

        if not blocos:
            linhas_resumo.append({
                "arquivo": pdf.name,
                "periodo": periodo,
                "rpa_caixa": rpa_caixa,
                "soma_receita": 0.0,
                **{f"soma_{c}": 0.0 for c in IMPOSTOS_COLS},
                "classificacao_periodo": "INCONSISTENTE",
            })
            continue

        soma_receita = 0.0
        soma_impostos = {c: 0.0 for c in IMPOSTOS_COLS}

        for b in blocos:
            b["arquivo"] = pdf.name
            b["periodo"] = periodo
            b["rpa_caixa"] = rpa_caixa
            b["classificacao"] = classificar_por_icms_iss(b["icms"], b["iss"])

            soma_receita += b["receita_atividade"]
            for c in IMPOSTOS_COLS:
                soma_impostos[c] += b[c]

            linhas_detalhadas.append(b)

        linhas_resumo.append({
            "arquivo": pdf.name,
            "periodo": periodo,
            "rpa_caixa": rpa_caixa,
            "soma_receita": soma_receita,
            **{f"soma_{c}": soma_impostos[c] for c in IMPOSTOS_COLS},
            "classificacao_periodo": classificar_por_icms_iss(soma_impostos["icms"], soma_impostos["iss"]),
        })

    df_detalhado = pd.DataFrame(linhas_detalhadas)
    df_resumo = pd.DataFrame(linhas_resumo)

    # (opcional) organiza colunas numa ordem boa
    cols_detalhe = [
        "arquivo", "periodo", "rpa_caixa",
        "descricao_atividade", "receita_atividade",
        *IMPOSTOS_COLS,
        "classificacao",
    ]
    cols_resumo = [
        "arquivo", "periodo", "rpa_caixa",
        "soma_receita",
        *[f"soma_{c}" for c in IMPOSTOS_COLS],
        "classificacao_periodo",
    ]

    df_detalhado = df_detalhado.reindex(columns=[c for c in cols_detalhe if c in df_detalhado.columns])
    df_resumo = df_resumo.reindex(columns=[c for c in cols_resumo if c in df_resumo.columns])

    with pd.ExcelWriter(saida_xlsx, engine="openpyxl") as writer:
        df_resumo.to_excel(writer, index=False, sheet_name="Resumo por Período")
        df_detalhado.to_excel(writer, index=False, sheet_name="Detalhado por Atividade")

    print(f"\n✔ Arquivo final gerado: {saida_xlsx}")


# =====================
# EXECUÇÃO
# =====================

if __name__ == "__main__":
    processar_pasta_pgdas(
        pasta_pgdas="pgdas",
        saida_xlsx="analise_pgdas_consolidada.xlsx"
    )
