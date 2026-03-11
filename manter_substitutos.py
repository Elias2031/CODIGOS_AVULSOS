from __future__ import annotations

import re
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, List


# ---------- REGEX DOS PADRÕES ----------

PADRAO_SPED = re.compile(
    r"^(?P<cnpj>\d{14})-"
    r"\d+-"
    r"(?P<dt_ini>\d{8})-"
    r"(?P<dt_fim>\d{8})-"
    r".+-SPED-EFD"
    r"(?:\..+)?$"
)

PADRAO_PISCOFINS = re.compile(
    r"^PISCOFINS_"
    r"(?P<dt_ini>\d{8})_"
    r"(?P<dt_fim>\d{8})_"
    r"(?P<cnpj>\d{14})_"
    r".+?"
    r"(?:\..+)?$"
)

EXTENSOES_ACEITAS = {".txt", ".sped", ".efd", ".zip", ".csv", ""}


@dataclass(frozen=True)
class ChaveGrupo:
    cnpj: str
    dt_ini: str
    dt_fim: str


# ---------- FUNÇÕES ----------

def coletar_arquivos(root: Path) -> List[Path]:
    return [p for p in root.rglob("*") if p.is_file()]


def extrair_chave(p: Path) -> ChaveGrupo | None:
    nome = p.name

    m = PADRAO_SPED.match(nome)
    if not m:
        m = PADRAO_PISCOFINS.match(nome)

    if not m:
        return None

    if p.suffix.lower() not in EXTENSOES_ACEITAS:
        return None

    return ChaveGrupo(
        cnpj=m.group("cnpj"),
        dt_ini=m.group("dt_ini"),
        dt_fim=m.group("dt_fim"),
    )


def escolher_mais_novo(arquivos: List[Path]) -> Path:
    return max(
        arquivos,
        key=lambda p: (p.stat().st_mtime, p.stat().st_size, p.name),
    )


def main(pasta_raiz: str, dry_run: bool = True) -> None:
    root = Path(pasta_raiz).resolve()

    grupos: Dict[ChaveGrupo, List[Path]] = {}

    for p in coletar_arquivos(root):
        chave = extrair_chave(p)
        if chave:
            grupos.setdefault(chave, []).append(p)

    total_excluir = 0

    for chave, arquivos in grupos.items():
        if len(arquivos) <= 1:
            continue

        manter = escolher_mais_novo(arquivos)
        excluir = [p for p in arquivos if p != manter]

        print(f"\n[GRUPO]")
        print(f" CNPJ    : {chave.cnpj}")
        print(f" PERÍODO : {chave.dt_ini} → {chave.dt_fim}")
        print(f" MANTER  : {manter.name}")

        for p in excluir:
            print(f" EXCLUIR : {p.name}")
            total_excluir += 1
            if not dry_run:
                p.unlink()

    print("\n=== RESUMO ===")
    print(f"Grupos processados : {len(grupos)}")
    print(f"Arquivos excluídos : {total_excluir}")
    print(f"Modo              : {'DRY-RUN' if dry_run else 'EXECUÇÃO REAL'}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        raise SystemExit("Uso: python dedup.py <pasta_raiz> [--run]")

    pasta = sys.argv[1]
    executar = "--run" in sys.argv
    main(pasta, dry_run=not executar)
