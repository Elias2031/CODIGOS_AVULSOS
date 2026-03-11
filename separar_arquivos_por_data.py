from __future__ import annotations

import shutil
import zipfile
from pathlib import Path
from datetime import datetime, date
import xml.etree.ElementTree as ET

CORTE = date(2023, 12, 31)  # 31/12/2023


def strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def parse_dt(text: str) -> datetime | None:
    if not text:
        return None
    s = text.strip()

    # ISO (com ou sem timezone): 2023-12-31T10:20:30-03:00
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass

    # Só data: 2023-12-31
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d")
    except Exception:
        pass

    # Data compacta: 20231231
    if len(s) == 8 and s.isdigit():
        try:
            return datetime.strptime(s, "%Y%m%d")
        except Exception:
            pass

    return None


def find_first(root: ET.Element, local_tag: str) -> str | None:
    """Busca o primeiro elemento com tag local (ignorando namespace) e retorna o texto."""
    local_tag = local_tag.lower()
    for el in root.iter():
        if strip_ns(el.tag).lower() == local_tag:
            txt = (el.text or "").strip()
            return txt or None
    return None


def extract_emission_date_from_xml(xml_bytes: bytes) -> date | None:
    """
    Extrai a data de emissão para:
    - NF-e / NFC-e: <ide><dhEmi> ou <ide><dEmi>
    - CT-e: <ide><dhEmi> (também pode aparecer como dEmi em alguns casos)
    Estratégia: tenta tags mais relevantes; se falhar, tenta heurística ISO em qualquer texto.
    """
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return None

    # Prioridade (mais comum e confiável)
    for tag in ("dhEmi", "dEmi"):
        val = find_first(root, tag)
        dt = parse_dt(val or "")
        if dt:
            return dt.date()

    # Alguns XMLs podem ter outros nomes (fallback leve)
    for tag in ("dhSaiEnt", "dSaiEnt"):
        val = find_first(root, tag)
        dt = parse_dt(val or "")
        if dt:
            return dt.date()

    # Heurística: procura qualquer texto com "YYYY-MM-DD"
    for el in root.iter():
        txt = (el.text or "").strip()
        if len(txt) >= 10 and txt[4:5] == "-" and txt[7:8] == "-":
            dt = parse_dt(txt)
            if dt:
                return dt.date()

    return None


def classify_zip_by_xml_dates(zip_path: Path) -> tuple[str, date | None]:
    """
    Lê XMLs do ZIP e retorna:
    - categoria: "ANTES", "DEPOIS", "INDEFINIDO"
    - data escolhida: menor data encontrada (mais conservador para corte)
    """
    dates: list[date] = []

    with zipfile.ZipFile(zip_path, "r") as zf:
        for name in zf.namelist():
            if not name.lower().endswith(".xml"):
                continue
            try:
                xml_bytes = zf.read(name)
            except Exception:
                continue
            d = extract_emission_date_from_xml(xml_bytes)
            if d:
                dates.append(d)

    if not dates:
        return "INDEFINIDO", None

    chosen = min(dates)  # conservador: se tiver mistura, cai no mais antigo
    if chosen < CORTE:
        return "ANTES", chosen
    return "DEPOIS", chosen  # inclui 31/12/2023 e após


def move_safe(src: Path, dst_dir: Path) -> Path:
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    if dst.exists():
        dst = dst_dir / f"{src.stem}__DUP{src.suffix}"
    shutil.move(str(src), str(dst))
    return dst


def separar_zips(pasta: str | Path) -> None:
    pasta = Path(pasta)
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não existe: {pasta}")

    out_antes = pasta / "ANTES_31-12-2023"
    out_depois = pasta / "DEPOIS_31-12-2023"
    out_indef = pasta / "DATA_NAO_ENCONTRADA"
    out_erro = pasta / "ZIP_INVALIDO_OU_ERRO"

    zips = sorted(pasta.glob("*.zip"))
    if not zips:
        print(f"[INFO] Nenhum .zip encontrado em: {pasta}")
        return

    for zp in zips:
        try:
            cat, d = classify_zip_by_xml_dates(zp)

            if cat == "ANTES":
                dst = move_safe(zp, out_antes)
                print(f"[OK] {zp.name} -> {dst.parent.name} | dhEmi/dEmi(min)={d} (< {CORTE})")
            elif cat == "DEPOIS":
                dst = move_safe(zp, out_depois)
                print(f"[OK] {zp.name} -> {dst.parent.name} | dhEmi/dEmi(min)={d} (>= {CORTE})")
            else:
                dst = move_safe(zp, out_indef)
                print(f"[WARN] {zp.name} -> {dst.parent.name} | sem dhEmi/dEmi")

        except zipfile.BadZipFile:
            dst = move_safe(zp, out_erro)
            print(f"[ERRO] ZIP inválido: {zp.name} -> {dst.parent.name}")
        except Exception as e:
            try:
                dst = move_safe(zp, out_erro)
                print(f"[ERRO] Falha: {zp.name} -> {dst.parent.name} | {e}")
            except Exception:
                print(f"[ERRO] Falha: {zp.name} | {e}")


if __name__ == "__main__":
    # Ajuste para sua pasta
    PASTA_ZIPS = r"C:\Users\elias.junior\Desktop\automacao-backup\NOVA_FASE\AUTOMACAO_RECUPERACAO_LUCRO_REAL\NOTAS_FISCAIS"
    separar_zips(PASTA_ZIPS)
