import csv
import ctypes
import re
import statistics
import time
import unicodedata
from collections import defaultdict, deque

import tkinter as tk
from tkinter import filedialog

import cv2
import numpy as np
import pytesseract
import uiautomation as uia
from mss import mss


WINDOW_TITLE = "Valor Folha - Fator R"
FISCAL_TITLE = "Fiscal"
WORKSPACE_TITLE = "Espaco de trabalho"
HEADER_LABELS = [
    "mes/ano",
    "folha e encargos",
    "folha",
    "contribuicao previdenciaria",
    "pgdas",
    "total",
    "log",
]


def canon(text):
    text = (text or "").strip().lower()
    text = "".join(
        ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn"
    )
    return " ".join(text.split())


def nome(ctrl):
    try:
        return (ctrl.Name or "").strip()
    except Exception:
        return ""


def tipo(ctrl):
    try:
        return ctrl.ControlTypeName or ""
    except Exception:
        return ""


def rect_of(ctrl):
    try:
        r = ctrl.BoundingRectangle
        return int(r.left), int(r.top), int(r.right), int(r.bottom)
    except Exception:
        return None


def center_xy(ctrl):
    rc = rect_of(ctrl)
    if not rc:
        return None
    l, t, r, b = rc
    return (l + r) // 2, (t + b) // 2


def iter_descendants(root, max_depth=12):
    q = deque([(root, 0)])
    while q:
        node, depth = q.popleft()
        if depth > max_depth:
            continue
        yield node
        try:
            for ch in node.GetChildren():
                q.append((ch, depth + 1))
        except Exception:
            continue


def find_fiscal_window():
    root = uia.GetRootControl()
    for w in root.GetChildren():
        try:
            if tipo(w) == "WindowControl" and canon(FISCAL_TITLE) in canon(nome(w)):
                return w
        except Exception:
            continue
    return None


def find_workspace_in_fiscal(fiscal_window):
    if not fiscal_window:
        return None
    for node in iter_descendants(fiscal_window, max_depth=4):
        try:
            if tipo(node) == "PaneControl" and canon(WORKSPACE_TITLE) in canon(nome(node)):
                return node
        except Exception:
            continue
    return None


def find_fator_window_inside_fiscal():
    fiscal = find_fiscal_window()
    if not fiscal:
        return None
    workspace = find_workspace_in_fiscal(fiscal)
    if not workspace:
        return None
    for node in workspace.GetChildren():
        try:
            if tipo(node) == "WindowControl" and canon(WINDOW_TITLE) in canon(nome(node)):
                return node
        except Exception:
            continue
    return None


def force_foreground_by_handle(hwnd):
    if not hwnd:
        return False
    try:
        user32 = ctypes.windll.user32
        user32.SetForegroundWindow(hwnd)
        user32.SetActiveWindow(hwnd)
        return True
    except Exception:
        return False


def focus_fiscal_and_fator_window(fator_window):
    fiscal = find_fiscal_window()
    if fiscal:
        force_foreground_by_handle(getattr(fiscal, "NativeWindowHandle", 0))
        try:
            fiscal.SetFocus()
        except Exception:
            pass
        time.sleep(0.12)

    if fator_window:
        force_foreground_by_handle(getattr(fator_window, "NativeWindowHandle", 0))
        try:
            fator_window.SetFocus()
        except Exception:
            pass
        time.sleep(0.12)


def wait_fator_window(timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        # Prioriza localizar a janela MDI dentro do Fiscal.
        w = find_fator_window_inside_fiscal()
        if w:
            return w

        # Fallback para casos em que o ERP exponha como janela global.
        root = uia.GetRootControl()
        for top in root.GetChildren():
            try:
                if tipo(top) == "WindowControl" and canon(WINDOW_TITLE) in canon(nome(top)):
                    return top
            except Exception:
                continue
        time.sleep(0.25)
    return None


def header_controls(window):
    by_label = defaultdict(list)
    for node in iter_descendants(window, max_depth=12):
        if tipo(node) != "TextControl":
            continue
        n = canon(nome(node))
        if n in HEADER_LABELS:
            by_label[n].append(node)

    out = {}
    for label in HEADER_LABELS:
        arr = by_label.get(label, [])
        if not arr:
            continue
        arr_sorted = sorted(arr, key=lambda c: (rect_of(c)[1], rect_of(c)[0]) if rect_of(c) else (10**9, 10**9))
        out[label] = arr_sorted[0]
    return out


def parent_chain(ctrl, stop_at=None):
    chain = []
    cur = ctrl
    while cur:
        chain.append(cur)
        if stop_at is not None and cur == stop_at:
            break
        try:
            cur = cur.GetParentControl()
        except Exception:
            break
    return chain


def contains_point(rect, x, y):
    if not rect:
        return False
    l, t, r, b = rect
    return l <= x <= r and t <= y <= b


def detect_grid_root(window, headers):
    controls = [headers[k] for k in HEADER_LABELS if k in headers]
    if len(controls) < 3:
        return None

    points = []
    for c in controls:
        pt = center_xy(c)
        if pt:
            points.append(pt)
    if len(points) < 3:
        return None

    chain = parent_chain(controls[0], stop_at=window)
    candidates = []
    for anc in chain:
        rc = rect_of(anc)
        if not rc:
            continue
        if all(contains_point(rc, x, y) for x, y in points):
            area = (rc[2] - rc[0]) * (rc[3] - rc[1])
            candidates.append((area, anc))

    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def try_get_clipboard_text():
    root = tk.Tk()
    root.withdraw()
    try:
        txt = root.clipboard_get()
    except Exception:
        txt = ""
    root.destroy()
    return txt


def parse_clipboard_table(txt):
    txt = (txt or "").strip()
    if not txt:
        return [], []

    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    if not lines:
        return [], []

    sep = "\t" if any("\t" in ln for ln in lines[:3]) else ";"
    rows = [ln.split(sep) for ln in lines]
    max_cols = max(len(r) for r in rows)
    if max_cols < 3:
        return [], []

    rows = [r + [""] * (max_cols - len(r)) for r in rows]
    first = [canon(x) for x in rows[0]]
    if any("mes/ano" in c for c in first):
        headers = rows[0]
        data_rows = rows[1:]
        return headers, data_rows

    # Alguns grids copiam sem cabecalho. Detecta por padrao MM/AAAA na primeira coluna.
    if re.fullmatch(r"\d{2}/\d{4}", (rows[0][0] or "").strip()):
        default_headers = ["Mês/Ano", "Folha e Encargos", "Folha", "Contribuição Previdenciária", "PGDAS", "Total", "Log"]
        headers = default_headers[:max_cols]
        return headers, rows

    return [], []


def try_copy_table_from_window(window):
    """
    Tenta copiar a grade por teclado de forma segura (sem loops de TAB).
    """
    before = try_get_clipboard_text()

    def try_copy_now():
        uia.SendKeys("^a")
        time.sleep(0.12)
        uia.SendKeys("^c")
        time.sleep(0.32)
        after = try_get_clipboard_text()
        if not after or after == before:
            return [], []
        return parse_clipboard_table(after)

    focus_fiscal_and_fator_window(window)
    uia.SendKeys("{ESC}")
    time.sleep(0.08)

    # Tentativa imediata no controle já focado
    headers, rows = try_copy_now()
    if headers and rows:
        return headers, rows

    # Segunda tentativa curta (sem navegar foco por TAB para evitar efeitos colaterais)
    focus_fiscal_and_fator_window(window)
    time.sleep(0.15)
    headers, rows = try_copy_now()
    if headers and rows:
        return headers, rows

    return [], []


def group_by_near_y(items, tol=6):
    if not items:
        return []
    items = sorted(items, key=lambda it: it["y"])
    groups = [[items[0]]]
    for it in items[1:]:
        if abs(it["y"] - statistics.median([x["y"] for x in groups[-1]])) <= tol:
            groups[-1].append(it)
        else:
            groups.append([it])
    return groups


def extract_table_by_uia(grid_root, headers):
    ordered = []
    for key in HEADER_LABELS:
        c = headers.get(key)
        if not c:
            continue
        pt = center_xy(c)
        rc = rect_of(c)
        if pt and rc:
            ordered.append({"key": key, "x": pt[0], "header_y": rc[1]})
    ordered.sort(key=lambda item: item["x"])
    if len(ordered) < 3:
        return []

    header_y = min(item["header_y"] for item in ordered)
    candidates = []
    checkboxes = []

    for n in iter_descendants(grid_root, max_depth=12):
        t = tipo(n)
        if t not in ("TextControl", "CheckBoxControl"):
            continue
        rc = rect_of(n)
        if not rc:
            continue
        x = (rc[0] + rc[2]) // 2
        y = (rc[1] + rc[3]) // 2
        if y <= header_y + 8:
            continue

        if t == "TextControl":
            txt = nome(n)
            if not txt:
                continue
            candidates.append({"x": x, "y": y, "text": txt})
        else:
            val = ""
            try:
                tg = n.GetTogglePattern()
                val = "Sim" if tg and tg.ToggleState == 1 else "Nao"
            except Exception:
                val = "Nao"
            checkboxes.append({"x": x, "y": y, "text": val})

    rows = []
    for grp in group_by_near_y(candidates, tol=6):
        row = {h["key"]: "" for h in ordered}
        for cell in grp:
            nearest = min(ordered, key=lambda col: abs(col["x"] - cell["x"]))
            key = nearest["key"]
            if row[key]:
                row[key] = f"{row[key]} | {cell['text']}"
            else:
                row[key] = cell["text"]

        row_center_y = int(statistics.mean([c["y"] for c in grp]))
        row_checks = [c for c in checkboxes if abs(c["y"] - row_center_y) <= 6]
        for chk in row_checks:
            nearest = min(ordered, key=lambda col: abs(col["x"] - chk["x"]))
            key = nearest["key"]
            if key in ("folha", "pgdas") and not row[key]:
                row[key] = chk["text"]

        if row.get("mes/ano"):
            rows.append(row)

    return rows


def find_button(window, label):
    wanted = canon(label)
    for node in iter_descendants(window, max_depth=8):
        if tipo(node) != "ButtonControl":
            continue
        if canon(nome(node)) == wanted:
            return node
    return None


def detect_table_region(window):
    """Detecta o retangulo da grade usando a barra de rolagem como referencia."""
    wrc = rect_of(window)
    if not wrc:
        return None
    wl, wt, wr, wb = wrc

    b_top = find_button(window, "Uma linha acima")
    b_bottom = find_button(window, "Uma linha abaixo")
    if b_top and b_bottom:
        rtc = rect_of(b_top)
        rbc = rect_of(b_bottom)
        if rtc and rbc:
            right = max(wl + 200, rtc[0] - 4)
            top = max(wt + 90, rtc[1] - 30)
            bottom = min(wb - 40, rbc[3] + 2)
            left = wl + 10
            if right - left > 250 and bottom - top > 140:
                return left, top, right, bottom

    # fallback por proporcao da janela
    left = wl + 10
    top = wt + 120
    right = wr - 130
    bottom = wb - 150
    if right - left <= 250 or bottom - top <= 140:
        return None
    return left, top, right, bottom


def capture_region_bgr(region):
    l, t, r, b = region
    monitor = {"left": int(l), "top": int(t), "width": int(r - l), "height": int(b - t)}
    with mss() as sct:
        shot = sct.grab(monitor)
    bgra = np.array(shot, dtype=np.uint8)
    return cv2.cvtColor(bgra, cv2.COLOR_BGRA2BGR)


def ocr_words(image_bgr):
    scaled = cv2.resize(image_bgr, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(scaled, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    thr = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 41, 13
    )
    data = pytesseract.image_to_data(
        thr,
        lang="por+eng",
        config="--oem 3 --psm 11",
        output_type=pytesseract.Output.DICT,
    )
    out = []
    n = len(data.get("text", []))
    for i in range(n):
        txt = (data["text"][i] or "").strip()
        if not txt:
            continue
        conf = -1
        try:
            conf = float(data["conf"][i])
        except Exception:
            pass
        if conf < 0:
            continue
        x = int(data["left"][i])
        y = int(data["top"][i])
        w = int(data["width"][i])
        h = int(data["height"][i])
        out.append({
            "text": txt,
            "canon": canon(txt),
            "x": x + w // 2,
            "y": y + h // 2,
            "left": x,
            "top": y,
            "w": w,
            "h": h,
            "conf": conf,
        })
    return out, thr


def infer_column_centers(words, img_w):
    # centros base (fallback)
    centers = {
        "mes_ano": int(img_w * 0.08),
        "folha_enc": int(img_w * 0.31),
        "folha": int(img_w * 0.39),
        "contrib": int(img_w * 0.58),
        "pgdas": int(img_w * 0.67),
        "total": int(img_w * 0.83),
        "log": int(img_w * 0.96),
    }

    top_words = [w for w in words if w["y"] <= 90]
    if not top_words:
        top_words = [w for w in words if w["y"] <= 90]

    def pick(pred, default_key):
        cands = [w for w in top_words if pred(w["canon"])]
        if cands:
            return int(statistics.median([w["x"] for w in cands]))
        return centers[default_key]

    centers["mes_ano"] = pick(lambda s: ("mes" in s and "ano" in s) or s == "mes/ano", "mes_ano")
    centers["folha_enc"] = pick(lambda s: "encargos" in s, "folha_enc")
    centers["contrib"] = pick(lambda s: s.startswith("contrib"), "contrib")
    centers["pgdas"] = pick(lambda s: "pgdas" in s, "pgdas")
    centers["total"] = pick(lambda s: s == "total", "total")
    centers["log"] = pick(lambda s: s == "log", "log")

    money_tuned = False
    # Ajuste robusto pelas colunas monetárias observadas nas linhas (R$ ...).
    money_x = sorted(w["x"] for w in words if "r$" in w["canon"] and w["y"] > 80)
    if money_x:
        clusters = [[money_x[0]]]
        for x in money_x[1:]:
            if abs(x - clusters[-1][-1]) <= 70:
                clusters[-1].append(x)
            else:
                clusters.append([x])
        med = sorted(int(statistics.median(c)) for c in clusters if len(c) >= 2)
        if len(med) >= 3:
            centers["folha_enc"] = med[0]
            centers["contrib"] = med[1]
            centers["total"] = med[2]
            centers["folha"] = int((centers["folha_enc"] + centers["contrib"]) / 2)
            money_tuned = True

    folhas = sorted([w["x"] for w in top_words if w["canon"] == "folha"])
    if len(folhas) >= 2 and not money_tuned:
        centers["folha_enc"] = folhas[0]
        centers["folha"] = folhas[-1]
    elif len(folhas) == 1 and not money_tuned:
        centers["folha_enc"] = folhas[0]
        centers["folha"] = min(img_w - 40, folhas[0] + int(img_w * 0.08))

    order = ["mes_ano", "folha_enc", "folha", "contrib", "pgdas", "total", "log"]
    for i in range(1, len(order)):
        prev_k = order[i - 1]
        cur_k = order[i]
        if centers[cur_k] <= centers[prev_k]:
            centers[cur_k] = centers[prev_k] + 20
    return centers


def build_column_bounds(centers, img_w):
    order = [
        ("Mês/Ano", "mes_ano"),
        ("Folha e Encargos", "folha_enc"),
        ("Folha", "folha"),
        ("Contribuição Previdenciária", "contrib"),
        ("PGDAS", "pgdas"),
        ("Total", "total"),
        ("Log", "log"),
    ]
    xs = [centers[k] for _, k in order]
    bounds = []
    for i, (label, key) in enumerate(order):
        left = 0 if i == 0 else (xs[i - 1] + xs[i]) // 2
        right = img_w if i == len(order) - 1 else (xs[i] + xs[i + 1]) // 2
        bounds.append((label, key, left, right, xs[i]))
    return bounds


def checkbox_marked(bin_img, x_center, y_center):
    h, w = bin_img.shape[:2]
    x0 = max(0, x_center - 7)
    x1 = min(w, x_center + 7)
    y0 = max(0, y_center - 7)
    y1 = min(h, y_center + 7)
    roi = bin_img[y0:y1, x0:x1]
    if roi.size == 0:
        return False
    dark = np.count_nonzero(roi == 0)
    return dark >= 18


def normalize_month_token(text):
    raw = (text or "").strip()
    if not raw:
        return None
    m = re.search(r"(\d{2})/(\d{4})", raw)
    if m:
        mm = int(m.group(1))
        yyyy = int(m.group(2))
        if 1 <= mm <= 12 and 2000 <= yyyy <= 2100:
            return f"{mm:02d}/{yyyy:04d}"

    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) >= 6:
        dd = digits[-6:]
        mm = int(dd[:2])
        yyyy = int(dd[2:])
        if 1 <= mm <= 12 and 2000 <= yyyy <= 2100:
            return f"{mm:02d}/{yyyy:04d}"
    if len(digits) == 5:
        mm = int(digits[0])
        yyyy = int(digits[1:])
        if 1 <= mm <= 9 and 2000 <= yyyy <= 2100:
            return f"{mm:02d}/{yyyy:04d}"
    return None


def extract_rows_from_ocr(words, col_bounds, bin_img):
    first_col = next((c for c in col_bounds if c[0] == "Mês/Ano"), None)
    if not first_col:
        return []
    _, _, first_left, first_right, _ = first_col

    month_words = []
    for w in words:
        if not (first_left <= w["x"] < first_right):
            continue
        mm = normalize_month_token(w["text"])
        if mm:
            month_words.append({**w, "month": mm})
    if not month_words:
        return []

    month_words.sort(key=lambda w: w["y"])
    rows_y = []
    for w in month_words:
        if not rows_y or abs(w["y"] - rows_y[-1][0]) > 7:
            rows_y.append([w["y"], w["month"]])

    result = []
    for y, mes in rows_y:
        row = {label: "" for label, _, _, _, _ in col_bounds}
        row["Mês/Ano"] = mes

        row_words = [w for w in words if abs(w["y"] - y) <= 8]
        for w in sorted(row_words, key=lambda x: x["x"]):
            for label, key, left, right, _cx in col_bounds:
                if left <= w["x"] < right:
                    if label == "Mês/Ano":
                        break
                    if label in ("Folha", "PGDAS"):
                        break
                    if row[label]:
                        row[label] += " " + w["text"]
                    else:
                        row[label] = w["text"]
                    break

        # marcações das colunas checkbox
        for label in ("Folha", "PGDAS"):
            for lbl, _key, left, right, cx in col_bounds:
                if lbl == label:
                    glyphs = [
                        (w["text"] or "").strip().lower()
                        for w in row_words
                        if left <= w["x"] < right and abs(w["y"] - y) <= 8
                    ]
                    if any(g in ("7", "v", "x", "/", "1", "iv", "vv", "iV".lower(), "V".lower()) for g in glyphs):
                        row[label] = "Sim"
                        break
                    if checkbox_marked(bin_img, cx, y):
                        row[label] = "Sim"
                    else:
                        row[label] = "Não"
                    break

        def pick_money(txt):
            m = re.search(r"R\\$\\s*[-\\d\\.,]+", txt or "")
            if m:
                return m.group(0).replace("]", "").strip()
            return ""

        # Higieniza colunas numéricas
        for label in ("Folha e Encargos", "Contribuição Previdenciária", "Total"):
            row[label] = pick_money(row[label]) or row[label].strip()

        # Alguns OCRs jogam o valor total junto de "Log"
        if (not row["Total"] or row["Total"] == "R$") and row["Log"]:
            money_in_log = pick_money(row["Log"])
            if money_in_log:
                row["Total"] = money_in_log
                row["Log"] = re.sub(r"R\\$\\s*[-\\d\\.,]+", "", row["Log"]).strip()

        if "log" in row["Log"].lower():
            row["Log"] = "Log"

        result.append(row)
    return result


def read_table_by_ocr(window):
    focus_fiscal_and_fator_window(window)
    region = detect_table_region(window)
    if not region:
        raise RuntimeError("Nao consegui detectar a regiao da tabela.")

    img = capture_region_bgr(region)
    words, bin_img = ocr_words(img)
    if not words:
        raise RuntimeError("OCR nao retornou texto da grade.")

    centers = infer_column_centers(words, bin_img.shape[1])
    col_bounds = build_column_bounds(centers, bin_img.shape[1])
    rows = extract_rows_from_ocr(words, col_bounds, bin_img)
    if not rows:
        raise RuntimeError("Nao foi possivel montar linhas da tabela por OCR.")

    headers = [label for label, _k, _l, _r, _x in col_bounds]
    data = [[r.get(h, "") for h in headers] for r in rows]
    return headers, data


def print_rows(headers, rows):
    print("\n[TABELA LIDA]")
    print(" | ".join(headers))
    print("-" * 100)
    for r in rows:
        print(" | ".join(r))
    print(f"\nTotal de linhas: {len(rows)}")


def ask_save_csv(headers, rows):
    root = tk.Tk()
    root.withdraw()
    path = filedialog.asksaveasfilename(
        title="Salvar tabela Fator R como CSV",
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        initialfile="tabela_fator_r.csv",
    )
    root.destroy()
    if not path:
        print("[INFO] Exportacao cancelada pelo usuario.")
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(headers)
        writer.writerows(rows)
    print(f"[OK] CSV salvo em: {path}")


def main():
    print("[INFO] Procurando janela 'Valor Folha - Fator R' dentro do Fiscal...")
    janela = wait_fator_window(timeout=30)
    if not janela:
        raise RuntimeError(
            "Nao encontrei a janela 'Valor Folha - Fator R' dentro do Fiscal. Deixe essa janela aberta e tente novamente."
        )
    try:
        janela.SetFocus()
    except Exception:
        pass

    print("[INFO] Lendo tabela por OCR (sem atalhos de copia)...")
    headers, rows = read_table_by_ocr(janela)

    print_rows(headers, rows)

    print("\n[INFO] Escolha onde salvar o CSV no popup.")
    ask_save_csv(headers, rows)


if __name__ == "__main__":
    main()
