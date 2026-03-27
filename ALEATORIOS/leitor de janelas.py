# -*- coding: utf-8 -*-
"""
Inspeciona janelas no Windows e exporta informações para TXT e CSV.
"""

import sys
import re
import csv
from typing import List, Optional, Iterable, Tuple
from pywinauto import Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.findwindows import ElementNotFoundError

FOCUS_TYPES = {
    "Button", "Edit", "ComboBox", "CheckBox", "RadioButton", "Hyperlink",
    "MenuItem", "ListItem", "TabItem", "TreeItem", "DataItem", "Pane",
    "Table", "List", "Tree", "Tab", "Window"
}

def short(text: str, maxlen: int = 80) -> str:
    text = (text or "").strip()
    return (text[: maxlen - 1] + "…") if len(text) > maxlen else text

def window_listing() -> List[UIAWrapper]:
    desktop = Desktop(backend="uia")
    return [w for w in desktop.windows() if w.is_visible()]

def describe_window(w: UIAWrapper) -> Tuple[str, int, str]:
    title = w.window_text().strip() or w.element_info.name or ""
    pid = w.process_id()
    cls = ""
    try:
        cls = w.class_name()
    except Exception:
        cls = ""
    return title, pid, cls

def selector_hint(ctrl: BaseWrapper) -> str:
    ei = ctrl.element_info
    parts = []
    if getattr(ei, "automation_id", None):
        parts.append(f"automation_id='{ei.automation_id}'")
    name = getattr(ei, "name", None) or ctrl.window_text()
    if name:
        safe_name = name.replace("'", "\\'")
        parts.append(f"name='{safe_name}'")
    ct = getattr(ei, "control_type", None)
    if ct:
        parts.append(f"control_type='{ct}'")
    if parts:
        return f".child_window({', '.join(parts)})"
    try:
        cls = ctrl.class_name()
        if cls:
            return f".child_window(class_name='{cls}')"
    except Exception:
        pass
    return "(sem dica de seletor)"

def rect_to_str(ctrl: BaseWrapper) -> str:
    try:
        r = ctrl.rectangle()
        return f"({r.left},{r.top})-({r.right},{r.bottom}) {r.width()}x{r.height()}"
    except Exception:
        return ""

def dump_controls(root: UIAWrapper,
                  txtfile: str,
                  csvfile: str,
                  types_filter: Optional[Iterable[str]] = None,
                  max_items: int = 1000) -> None:
    if types_filter is None:
        types_filter = FOCUS_TYPES

    with open(txtfile, "w", encoding="utf-8") as f_txt, \
         open(csvfile, "w", newline="", encoding="utf-8") as f_csv:

        writer = csv.writer(f_csv)
        writer.writerow(["ControlType", "Name", "AutomationID", "Class",
                         "Enabled", "Bounds", "SelectorHint"])

        f_txt.write("=== Controles detectados ===\n\n")

        count = 0
        try:
            descendants = root.descendants()
        except ElementNotFoundError:
            f_txt.write("Nenhum controle encontrado.\n")
            return
        except Exception as e:
            f_txt.write(f"Falha ao obter descendentes: {e}\n")
            return

        for c in descendants:
            if count >= max_items:
                f_txt.write(f"\n[limite] Mostrando apenas os primeiros {max_items} controles.\n")
                break
            try:
                ei = c.element_info
                ctype = getattr(ei, "control_type", None)
                if ctype and ctype not in types_filter:
                    continue

                name = (getattr(ei, "name", None) or c.window_text() or "").strip()
                auto_id = getattr(ei, "automation_id", "") or ""
                cls = ""
                try:
                    cls = c.class_name()
                except Exception:
                    pass
                enabled = ""
                try:
                    enabled = "enabled" if c.is_enabled() else "disabled"
                except Exception:
                    pass

                bounds = rect_to_str(c)
                hint = selector_hint(c)

                # TXT
                f_txt.write(f"- [{ctype or '?'}] name='{short(name,70)}'  "
                            f"automation_id='{auto_id}'  class='{cls}'  {enabled}\n")
                f_txt.write(f"  bounds: {bounds}\n")
                f_txt.write(f"  dica seletor: {hint}\n\n")

                # CSV
                writer.writerow([ctype or "?", name, auto_id, cls, enabled, bounds, hint])

                count += 1
            except Exception as e:
                f_txt.write(f"(aviso) falha ao ler um controle: {e}\n")
                continue

def main():
    wins = window_listing()
    if not wins:
        print("Nenhuma janela encontrada.")
        return

    print("=== Janelas Abertas ===")
    for idx, w in enumerate(wins):
        title, pid, cls = describe_window(w)
        print(f"{idx:>3} - {short(title, 60)} (PID={pid}, Classe={cls})")

    sel = input("\nDigite o número da janela (ou Enter para sair): ").strip()
    if sel == "":
        return
    if not re.fullmatch(r"\d+", sel):
        print("Entrada inválida.")
        return
    idx = int(sel)
    if not (0 <= idx < len(wins)):
        print("Número inválido.")
        return

    target = wins[idx]
    title, pid, cls = describe_window(target)

    save_dir = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\ALEATORIOS"
    base = f"janela_{idx}_inspecao"
    txtfile = f"{save_dir}\\{base}.txt"
    csvfile = f"{save_dir}\\{base}.csv"


    with open(txtfile, "w", encoding="utf-8") as f:
        f.write("=== Janela Selecionada ===\n")
        f.write(f"Título : {title or '(sem título)'}\n")
        f.write(f"PID    : {pid}\n")
        f.write(f"Classe : {cls}\n\n")

    dump_controls(target, txtfile, csvfile)

    print(f"\nExport concluído:\n  {txtfile}\n  {csvfile}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nEncerrado pelo usuário.")
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)
