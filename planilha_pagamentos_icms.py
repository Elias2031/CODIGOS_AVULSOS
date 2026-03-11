import re
from pathlib import Path

import pdfplumber
import pandas as pd


PASTA_PDFS = Path("PAGAMENTOS DE ICMS")
OUTPUT_XLSX = Path("pagamentos_consolidados_ICMS_JA.xlsx")


def br_to_float(valor: str) -> float:
    if not valor:
        return 0.0
    valor = valor.replace(".", "").replace(",", ".")
    return float(valor)


def extrair_pagamentos(pdf_path: Path) -> list[dict]:
    registros = []
    tipo_tributo_atual = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue

            linhas = texto.splitlines()

            for linha in linhas:
                # Tipo do tributo
                m_tipo = re.search(r"TIPO DO TRIBUTO\s*:\s*(.+)", linha)
                if m_tipo:
                    tipo_tributo_atual = m_tipo.group(1).strip()
                    continue

                # Linha de pagamento
                if re.match(r"\d{2}/\d{2}/\d{4}", linha):
                    partes = re.split(r"\s+", linha)

                    try:
                        registros.append({
                            "arquivo_origem": pdf_path.name,
                            "tipo_tributo": tipo_tributo_atual,
                            "data_pagamento": partes[0],
                            "periodo": partes[1],
                            "data_vencimento": partes[2],
                            "codigo_receita": partes[3],
                            "valor_principal": br_to_float(partes[4]),
                            "valor_pago": br_to_float(partes[5]),
                            "banco_agencia": partes[6],
                            "lote": partes[7],
                            "numero_documento": partes[9],
                            "situacao": partes[10],
                        })
                    except Exception:
                        pass

    return registros


def main():
    todos_registros = []

    for pdf in PASTA_PDFS.glob("*.pdf"):
        print(f"[INFO] Processando {pdf.name}")
        dados_pdf = extrair_pagamentos(pdf)
        todos_registros.extend(dados_pdf)

    df = pd.DataFrame(todos_registros)

    if df.empty:
        print("[ERRO] Nenhum dado encontrado.")
        return

    # Conversão de datas
    for col in ["data_pagamento", "data_vencimento"]:
        df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")

    df = df.sort_values(["tipo_tributo", "data_pagamento"])

    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"[OK] Planilha gerada: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
