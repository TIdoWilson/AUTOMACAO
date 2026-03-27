from __future__ import annotations

import argparse
import re
import shutil
import sys
import tkinter as tk
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from tkinter import filedialog
import xml.etree.ElementTree as ET


def parse_decimal(value: str) -> Decimal:
    raw = (value or "").strip()
    if not raw:
        return Decimal("0")

    normalized = raw.replace(".", "").replace(",", ".") if "," in raw else raw
    try:
        return Decimal(normalized)
    except InvalidOperation:
        return Decimal("0")


def format_decimal_br(value: Decimal) -> str:
    rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{rounded:.2f}".replace(".", ",")


def ask_folder() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askdirectory(title="Selecione a pasta com os XML")
    root.destroy()
    if not selected:
        return None
    return Path(selected)


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def namespace_of(tag: str) -> str:
    if tag.startswith("{") and "}" in tag:
        return tag[1:].split("}", 1)[0]
    return ""


def qname(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}" if namespace else tag


def first_descendant(element: ET.Element | None, tag_name: str) -> ET.Element | None:
    if element is None:
        return None
    for child in element.iter():
        if local_name(child.tag) == tag_name:
            return child
    return None


def direct_child(element: ET.Element | None, tag_name: str) -> ET.Element | None:
    if element is None:
        return None
    for child in list(element):
        if local_name(child.tag) == tag_name:
            return child
    return None


def direct_child_index(element: ET.Element, tag_name: str) -> int:
    for index, child in enumerate(list(element)):
        if local_name(child.tag) == tag_name:
            return index
    return len(list(element))


def register_namespaces(xml_path: Path) -> None:
    namespaces: dict[str, str] = {}
    for _, node in ET.iterparse(xml_path, events=("start-ns",)):
        prefix, uri = node
        if prefix not in namespaces:
            namespaces[prefix] = uri

    for prefix, uri in namespaces.items():
        if re.fullmatch(r"ns\d+", prefix or ""):
            continue
        ET.register_namespace(prefix, uri)


def ensure_child(parent: ET.Element, namespace: str, tag_name: str, insert_at: int | None = None) -> ET.Element:
    child = direct_child(parent, tag_name)
    if child is not None:
        return child

    child = ET.Element(qname(namespace, tag_name))
    if insert_at is None or insert_at >= len(list(parent)):
        parent.append(child)
    else:
        parent.insert(insert_at, child)
    return child


def update_xml(xml_path: Path) -> tuple[bool, str]:
    register_namespaces(xml_path)

    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = namespace_of(root.tag)

    inf_dps = first_descendant(root, "infDPS")
    valores = direct_child(inf_dps, "valores")
    v_serv = first_descendant(valores, "vServ")
    bruto = parse_decimal(v_serv.text if v_serv is not None else "")
    novo_csll = bruto * Decimal("0.01")
    novo_texto = format_decimal_br(novo_csll)

    if valores is None:
        return False, "tag valores nao encontrada"

    trib = ensure_child(valores, ns, "trib")
    tot_trib_index = direct_child_index(trib, "totTrib")
    trib_fed = direct_child(trib, "tribFed")
    if trib_fed is None:
        trib_fed = ET.Element(qname(ns, "tribFed"))
        trib.insert(tot_trib_index, trib_fed)

    ret_cp_index = direct_child_index(trib_fed, "vRetCP")
    ret_irrf_index = direct_child_index(trib_fed, "vRetIRRF")
    insert_at = max(ret_cp_index, ret_irrf_index)
    if insert_at < len(list(trib_fed)):
        insert_at += 1

    v_ret_csll = ensure_child(trib_fed, ns, "vRetCSLL", insert_at=insert_at)
    v_ret_csll.text = novo_texto

    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
    return True, novo_texto


def backup_xmls(folder: Path, backup_folder: Path) -> int:
    backup_folder.mkdir(exist_ok=True)
    copied = 0
    for xml_file in folder.glob("*.xml"):
        target = backup_folder / xml_file.name
        if not target.exists():
            shutil.copy2(xml_file, target)
            copied += 1
    return copied


def process_folder(folder: Path) -> int:
    if not folder.exists() or not folder.is_dir():
        print("Pasta invalida.")
        return 1

    xml_files = sorted(folder.glob("*.xml"))
    if not xml_files:
        print("Nenhum XML encontrado na pasta selecionada.")
        return 1

    backup_folder = folder / "originais"
    copied = backup_xmls(folder, backup_folder)

    updated = 0
    for xml_file in xml_files:
        ok, result = update_xml(xml_file)
        if not ok:
            print(f"[ERRO] {xml_file.name}: {result}")
            continue
        updated += 1
        print(f"[OK] {xml_file.name}: vRetCSLL = {result}")

    print()
    print(f"Pasta processada: {folder}")
    print(f"Backup em: {backup_folder}")
    print(f"XMLs com backup criado nesta execucao: {copied}")
    print(f"XMLs atualizados: {updated}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Ajusta a tag vRetCSLL para 1% do valor bruto da nota em formato brasileiro."
    )
    parser.add_argument(
        "--pasta",
        help="Pasta com os XMLs. Se omitida, o script abre um seletor de pasta.",
    )
    args = parser.parse_args()

    selected_folder = Path(args.pasta) if args.pasta else ask_folder()
    if selected_folder is None:
        print("Selecao cancelada pelo usuario.")
        return 1

    return process_folder(selected_folder)


if __name__ == "__main__":
    sys.exit(main())
