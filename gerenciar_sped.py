#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerencia arquivos SPED com nome:
CNPJ-IE-YYYYMMDD-YYYYMMDD-SEQUENCIAL-HASH-SPED-EFD[.ext]
Mantém o arquivo vencedor e move os demais do mesmo período p/ 'excluir/'.

Desempate:
1) maior SEQUENCIAL
2) mtime (mais recente)
3) tamanho (maior)
4) nome (ordem alfabética) - fallback

Uso:
  python gerenciar_sped.py --dir "C:\\SPED"            # dry-run
  python gerenciar_sped.py --dir "C:\\SPED" --apply    # aplica
"""

import re
import sys
import argparse
from pathlib import Path
import shutil
from collections import defaultdict

FILENAME_RE = re.compile(
    r"""^
    (?P<cnpj>\d{14})
    -
    (?P<ie>[A-Za-z0-9]{5,20})
    -
    (?P<dt_ini>\d{8})
    -
    (?P<dt_fim>\d{8})
    -
    (?P<seq>\d+)
    -
    (?P<hash>[A-Fa-f0-9]{40})
    -
    (?P<sistema>[A-Za-z0-9]+)
    -
    (?P<modulo>[A-Za-z0-9]+)
    (?:\.[A-Za-z0-9]+)?$
    """,
    re.VERBOSE
)

def parse_args():
    ap = argparse.ArgumentParser(description="Manter retificador por período e mover demais para 'excluir/'.")
    ap.add_argument("--dir", required=True, help="Pasta com os arquivos.")
    ap.add_argument("--apply", action="store_true", help="Executa as movimentações (sem isso é dry-run).")
    ap.add_argument("--excluir", default="excluir", help="Nome da subpasta para arquivos removidos.")
    return ap.parse_args()

def key_period(info):
    return (
        info["cnpj"],
        info["ie"],
        info["dt_ini"],
        info["dt_fim"],
        info["sistema"].upper(),
        info["modulo"].upper(),
    )

def winner_sort_key(p, seq):
    # desempate: maior seq, mtime mais novo, maior tamanho, nome
    stat = p.stat()
    return (seq, stat.st_mtime, stat.st_size, p.name)

def main():
    args = parse_args()
    base = Path(args.dir).expanduser().resolve()
    if not base.is_dir():
        print(f"[ERRO] Pasta não encontrada: {base}", file=sys.stderr)
        sys.exit(1)

    grupos = defaultdict(list)

    for p in base.iterdir():
        if not p.is_file():
            continue
        m = FILENAME_RE.match(p.name)
        if not m:
            continue
        info = m.groupdict()
        seq = int(info["seq"])
        grupos[key_period(info)].append({"path": p, "seq": seq, "name": p.name})

    if not grupos:
        print("[INFO] Nenhum arquivo no padrão esperado foi encontrado.")
        return

    destino_excluir = base / args.excluir
    move_list = []

    for chave, itens in grupos.items():
        cnpj, ie, dt_ini, dt_fim, sistema, modulo = chave
        # Ordena pelo critério de vencedor (desc nos dois primeiros)
        itens_sorted = sorted(
            itens,
            key=lambda x: winner_sort_key(x["path"], x["seq"]),
            reverse=True
        )
        vencedor = itens_sorted[0]
        perdedores = itens_sorted[1:]

        print(f"\n[GRUPO] {cnpj}-{ie}-{dt_ini}-{dt_fim}-{sistema}-{modulo}")
        print(f"  [+] Mantém: {vencedor['name']} (seq={vencedor['seq']})")

        if perdedores:
            for perd in perdedores:
                print(f"  [-] Excluir: {perd['name']} (seq={perd['seq']})")
                move_list.append(perd["path"])
        else:
            print("  [=] Só um arquivo no período; nada a excluir.")

    if not move_list:
        print("\n[OK] Não há duplicados por período. Nada a mover.")
        return

    if not args.apply:
        print("\n[DRY-RUN] Arquivos que seriam movidos para 'excluir/':")
        for p in move_list:
            print("   ->", p.name)
        print("\nUse --apply para aplicar as mudanças.")
        return

    destino_excluir.mkdir(exist_ok=True)
    moved = 0
    for p in move_list:
        destino = destino_excluir / p.name
        if destino.exists():
            # evita overwrite criando sufixo
            stem, suffix = destino.stem, destino.suffix
            i = 1
            while (destino_excluir / f"{stem}__dup{i}{suffix}").exists():
                i += 1
            destino = destino_excluir / f"{stem}__dup{i}{suffix}"
        shutil.move(str(p), str(destino))
        moved += 1
        print(f"[MOVE] {p.name} -> {destino.relative_to(base)}")

    print(f"\n[OK] Movidos {moved} arquivo(s) para '{args.excluir}/'.")

if __name__ == "__main__":
    main()
