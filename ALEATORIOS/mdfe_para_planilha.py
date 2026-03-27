# mdfe_para_planilha_gui_descargas_por_coluna.py
import os
import glob
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

try:
    import requests
except ImportError:
    requests = None


IBGE_MUNICIPIO_URL = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios/{codigo}"


def _ns_from_root(root: ET.Element) -> str:
    if root.tag.startswith("{") and "}" in root.tag:
        return root.tag.split("}")[0].strip("{")
    return ""


def _safe_find(elem: ET.Element, path: str, ns: dict) -> ET.Element | None:
    try:
        return elem.find(path, ns)
    except Exception:
        return None


def _safe_findall(elem: ET.Element, path: str, ns: dict) -> list[ET.Element]:
    try:
        return elem.findall(path, ns)
    except Exception:
        return []


def _safe_text(elem: ET.Element | None) -> str:
    if elem is None or elem.text is None:
        return ""
    return elem.text.strip()


def uf_por_codigo_ibge(codigo: str, cache: dict[str, str]) -> str:
    """
    Retorna a UF (sigla) a partir do código IBGE do município.
    Usa cache para reduzir chamadas.
    """
    codigo = (codigo or "").strip()
    if not codigo:
        return ""

    if codigo in cache:
        return cache[codigo]

    if requests is None:
        # Sem requests instalado
        return ""

    try:
        url = IBGE_MUNICIPIO_URL.format(codigo=codigo)
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()

        # Estrutura típica: data["microrregiao"]["mesorregiao"]["UF"]["sigla"]
        uf = ""
        try:
            uf = data["microrregiao"]["mesorregiao"]["UF"]["sigla"]
        except Exception:
            # fallback mais defensivo
            uf = (
                data.get("regiao-imediata", {})
                    .get("regiao-intermediaria", {})
                    .get("UF", {})
                    .get("sigla", "")
            )

        cache[codigo] = uf or ""
        return cache[codigo]
    except Exception:
        cache[codigo] = ""
        return ""


def parse_mdfe_xml(xml_path: str, uf_cache: dict[str, str]) -> dict:
    tree = ET.parse(xml_path)
    root = tree.getroot()

    nsuri = _ns_from_root(root)
    ns = {"m": nsuri} if nsuri else {}

    # Chave MDFe
    inf_mdfe = _safe_find(root, ".//m:infMDFe" if ns else ".//infMDFe", ns)
    mdfe_id = inf_mdfe.get("Id", "") if inf_mdfe is not None else ""
    chave_mdfe = mdfe_id[4:] if mdfe_id.startswith("MDFe") else mdfe_id  # 44 dígitos

    # Número MDFe
    nmdf = _safe_text(_safe_find(root, ".//m:ide/m:nMDF" if ns else ".//ide/nMDF", ns))

    # Descargas
    inf_mun_descargas = _safe_findall(root, ".//m:infMunDescarga" if ns else ".//infMunDescarga", ns)

    descargas = []  # lista de strings "Municipio - UF"
    chaves_nfe = []

    for d in inf_mun_descargas:
        xmun = _safe_text(_safe_find(d, "m:xMunDescarga" if ns else "xMunDescarga", ns))
        cmun = _safe_text(_safe_find(d, "m:cMunDescarga" if ns else "cMunDescarga", ns))

        uf = uf_por_codigo_ibge(cmun, uf_cache)
        if xmun:
            if uf:
                descargas.append(f"{xmun} - {uf}")
            else:
                descargas.append(xmun)

        # NF-e vinculadas (chaves)
        inf_nfes = _safe_findall(d, "m:infNFe" if ns else "infNFe", ns)
        for infnfe in inf_nfes:
            ch = _safe_text(_safe_find(infnfe, "m:chNFe" if ns else "chNFe", ns))
            if ch:
                chaves_nfe.append(ch)

    # remove duplicados preservando ordem
    def uniq(seq):
        seen = set()
        out = []
        for x in seq:
            if x not in seen:
                seen.add(x)
                out.append(x)
        return out

    descargas_uniq = uniq(descargas)
    chaves_nfe_uniq = uniq(chaves_nfe)

    return {
        "chave_mdfe": chave_mdfe,
        "numero_mdfe": nmdf,
        "quantidade_nfe_vinculadas": len(chaves_nfe_uniq),
        "chaves_nfe": ";".join(chaves_nfe_uniq),
        "arquivo_xml": os.path.basename(xml_path),
        "descargas_lista": descargas_uniq,  # para depois virar colunas
    }


def ajustar_excel(xlsx_path: str):
    wb = load_workbook(xlsx_path)
    ws = wb.active

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Quebra de linha na coluna de chaves (D no layout abaixo)
    wrap_cols = {"D"}
    for col in ws.columns:
        col_letter = col[0].column_letter
        max_len = 0
        for cell in col:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
            if col_letter in wrap_cols:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 80)

    wb.save(xlsx_path)


def escolher_pasta() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    pasta = filedialog.askdirectory(title="Selecione a pasta com os XML de MDFe")
    root.destroy()
    return pasta or ""


def escolher_saida_xlsx() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    caminho = filedialog.asksaveasfilename(
        title="Salvar planilha",
        defaultextension=".xlsx",
        filetypes=[("Excel (*.xlsx)", "*.xlsx")],
        initialfile="mdfe_saida.xlsx",
    )
    root.destroy()
    return caminho or ""


def main():
    if requests is None:
        messagebox.showwarning(
            "Dependência ausente",
            "O pacote 'requests' não está instalado.\n\n"
            "Instale com:\n"
            "pip install requests\n\n"
            "Sem isso, o script não consegue buscar a UF via IBGE."
        )
        return

    pasta = escolher_pasta()
    if not pasta:
        messagebox.showinfo("Cancelado", "Nenhuma pasta selecionada.")
        return

    arquivos = sorted(glob.glob(os.path.join(pasta, "*.xml")))
    if not arquivos:
        messagebox.showwarning("Atenção", f"Nenhum XML encontrado em:\n{pasta}")
        return

    saida = escolher_saida_xlsx()
    if not saida:
        messagebox.showinfo("Cancelado", "Nenhum arquivo de saída selecionado.")
        return

    uf_cache: dict[str, str] = {}

    registros = []
    erros = []

    # Primeiro: parse para descobrir o máximo de descargas (para criar colunas)
    parsed = []
    max_desc = 0

    for arq in arquivos:
        try:
            item = parse_mdfe_xml(arq, uf_cache)
            parsed.append(item)
            max_desc = max(max_desc, len(item["descargas_lista"]))
        except Exception as e:
            erros.append(f"{os.path.basename(arq)} -> {e}")
            parsed.append({
                "chave_mdfe": "",
                "numero_mdfe": "",
                "quantidade_nfe_vinculadas": 0,
                "chaves_nfe": "",
                "arquivo_xml": os.path.basename(arq),
                "descargas_lista": [],
            })

    # Monta linhas com colunas descarga_1..descarga_N
    for item in parsed:
        row = {
            "chave_mdfe": item["chave_mdfe"],
            "numero_mdfe": item["numero_mdfe"],
            "quantidade_nfe_vinculadas": item["quantidade_nfe_vinculadas"],
            "chaves_nfe": item["chaves_nfe"],
            "arquivo_xml": item["arquivo_xml"],
        }
        for i in range(max_desc):
            col = f"descarga_{i+1}"
            row[col] = item["descargas_lista"][i] if i < len(item["descargas_lista"]) else ""
        registros.append(row)

    # Define ordem de colunas
    cols = ["chave_mdfe", "numero_mdfe", "quantidade_nfe_vinculadas", "chaves_nfe"]
    cols += [f"descarga_{i+1}" for i in range(max_desc)]
    cols += ["arquivo_xml"]

    df = pd.DataFrame(registros, columns=cols)
    df.to_excel(saida, index=False)
    ajustar_excel(saida)

    msg = (
        f"Planilha gerada com sucesso:\n{saida}\n\n"
        f"Arquivos: {len(arquivos)}\n"
        f"Máx. municípios de descarga em um MDFe: {max_desc}"
    )
    if erros:
        msg += f"\n\nErros ao ler: {len(erros)} (detalhes no console)"
        print("Erros:")
        for x in erros:
            print(" -", x)

    messagebox.showinfo("Concluído", msg)


if __name__ == "__main__":
    main()
