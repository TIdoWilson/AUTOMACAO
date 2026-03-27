# -*- coding: utf-8 -*-
"""
Atualiza a tabela de faturamento em um arquivo .doc ou .docx.

Comportamento:
- Duplica a ÚLTIMA página do documento.
- Na nova página:
    * atualiza a tabela (rotaciona meses, soma total),
    * atualiza "Documento enviado para ..." (se informado),
    * atualiza a data em "Francisco Beltrão - PR, ...".
- Cabeçalho, assinatura e linha azul permanecem como no modelo.
"""

import os
import re
import shutil
import tempfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client as win32

# ========================= Constantes Word ==========================

WD_GO_TO_PAGE = 1                  # WdGoToItem.wdGoToPage
WD_GO_TO_LAST = -1                 # WdGoToDirection.wdGoToLast
WD_STORY = 6                       # WdUnits.wdStory
WD_SECTION_BREAK_NEXT_PAGE = 2     # WdBreakType.wdSectionBreakNextPage

WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2

# ========================= Datas ==========================

MESES_EXTENSO = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}

MAPA_MES_ABREV = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}
INV_MES_ABREV = {v: k for k, v in MAPA_MES_ABREV.items()}


def data_hoje_extenso() -> str:
    hoje = datetime.today()
    dia = hoje.day
    mes_nome = MESES_EXTENSO.get(hoje.month, "")
    ano = hoje.year
    return f"{dia:02d} de {mes_nome} de {ano}"


def interpretar_data_usuario(texto: str) -> str:
    """
    Converte entrada do usuário em data por extenso.

    Exemplos aceitos:
        10102025
        10/10/2025
        10-10-25
        10/102025
        1010/2025
    Se tiver letras (ex.: '10 de outubro de 2025'), é retornado como está.
    """
    if texto is None:
        raise ValueError("Data vazia.")

    texto = texto.strip()
    if not texto:
        raise ValueError("Data vazia.")

    # Já está por extenso
    if re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", texto):
        return texto

    nums = re.sub(r"\D", "", texto)
    if len(nums) < 4:
        raise ValueError("Informe pelo menos dia e mês.")

    dia = mes = ano = None

    if len(nums) == 8:
        # ddmmyyyy
        dia = int(nums[:2])
        mes = int(nums[2:4])
        ano = int(nums[4:])
    elif len(nums) == 6:
        # ddmmyy
        dia = int(nums[:2])
        mes = int(nums[2:4])
        ano = int(nums[4:])
        ano = 2000 + ano if ano <= 49 else 1900 + ano
    elif len(nums) == 4:
        # ddmm (ano corrente)
        dia = int(nums[:2])
        mes = int(nums[2:4])
        ano = datetime.today().year
    else:
        dia = int(nums[:2])
        mes = int(nums[2:4])
        ano = int(nums[4:])
        if ano < 100:
            ano = 2000 + ano if ano <= 49 else 1900 + ano

    if not (1 <= dia <= 31):
        raise ValueError("Dia inválido.")
    if not (1 <= mes <= 12):
        raise ValueError("Mês inválido.")

    mes_nome = MESES_EXTENSO.get(mes)
    if not mes_nome:
        raise ValueError("Mês inválido.")

    return f"{dia:02d} de {mes_nome} de {ano}"


# ========================= Valores monetários ==========================

def parse_valor_monetario(texto) -> Decimal:
    """Converte texto (pt-BR ou bruto) para Decimal. Suporta negativos."""
    s = str(texto or "").strip()
    if not s:
        return Decimal("0.00")

    s = s.replace("R$", "").replace(" ", "")

    neg = False
    if s.startswith("-"):
        neg = True
        s = s[1:]
    elif s.startswith("+"):
        s = s[1:]

    if not s:
        return Decimal("0.00")

    s = re.sub(r"[^0-9,\.]", "", s)

    try:
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
            val = Decimal(s)
        elif "," in s:
            s = s.replace(".", "").replace(",", ".")
            val = Decimal(s)
        elif "." in s:
            val = Decimal(s)
        else:
            if not s.isdigit():
                s_digits = re.sub(r"\D", "", s) or "0"
            else:
                s_digits = s
            val = Decimal(s_digits) / Decimal("100")
    except InvalidOperation:
        val = Decimal("0.00")

    if neg:
        val = -val

    return val.quantize(Decimal("0.01"))


def format_valor_brl(valor: Decimal) -> str:
    """Formata Decimal no padrão brasileiro 1.234,56 (com suporte a negativos)."""
    valor = valor.quantize(Decimal("0.01"))
    neg = valor < 0
    if neg:
        valor = -valor

    inteiro = int(valor)
    centavos = int((valor - Decimal(inteiro)) * 100)

    inteiro_str = f"{inteiro:,}".replace(",", ".")
    s = f"{inteiro_str},{centavos:02d}"
    if neg:
        s = "-" + s
    return s


def interpretar_valor_usuario(texto_usuario: str):
    valor_num = parse_valor_monetario(texto_usuario)
    valor_str = format_valor_brl(valor_num)
    return valor_num, valor_str


# ========================= Mês / Ano ==========================

def parse_mes_ano(rotulo: str):
    """
    Tenta interpretar vários formatos de mês/ano e devolve (mes, ano).
    """
    texto = (rotulo or "").strip()
    if not texto:
        raise ValueError("Texto vazio para mês/ano")

    m = re.search(r"([A-Za-z]{3})\s*[/\-\s]\s*(\d{2,4})", texto)
    if m:
        abrev = m.group(1).strip().upper()
        if abrev not in MAPA_MES_ABREV:
            raise ValueError(f"Abreviação de mês desconhecida: '{abrev}'")
        mes = MAPA_MES_ABREV[abrev]
        ano_txt = m.group(2)
        ano = int(ano_txt)
        if len(ano_txt) == 2:
            ano = 2000 + ano if ano <= 49 else 1900 + ano
        return mes, ano

    m = re.search(r"(\d{1,2})\s*[/\-]\s*(\d{2,4})", texto)
    if m:
        mes = int(m.group(1))
        ano_txt = m.group(2)
        ano = int(ano_txt)
        if len(ano_txt) == 2:
            ano = 2000 + ano if ano <= 49 else 1900 + ano
        if not (1 <= mes <= 12):
            raise ValueError(f"Mês fora do intervalo 1-12 em '{texto}'")
        return mes, ano

    m = re.search(r"(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{2,4})", texto)
    if m:
        g1, g2, g3 = m.groups()
        if len(g1) == 4:
            ano_txt = g1
            mes = int(g2)
        elif len(g3) == 4:
            ano_txt = g3
            mes = int(g2)
        else:
            ano_txt = g3
            mes = int(g2)

        ano = int(ano_txt)
        if len(ano_txt) == 2:
            ano = 2000 + ano if ano <= 49 else 1900 + ano
        if not (1 <= mes <= 12):
            raise ValueError(f"Mês fora do intervalo 1-12 em '{texto}'")
        return mes, ano

    raise ValueError(f"Não consegui interpretar o mês/ano em: '{rotulo}'")


def format_mes_ano(mes: int, ano: int, base_str: str = None) -> str:
    abrev = INV_MES_ABREV.get(mes)
    if not abrev:
        raise ValueError(f"Mês inválido: {mes}")

    ano2 = ano % 100
    rotulo = f"{abrev}/{ano2:02d}"

    if base_str:
        base = base_str.strip()
        if base.islower():
            rotulo = rotulo.lower()
        elif base[0].isupper() and base[1:].islower():
            rotulo = rotulo.capitalize()

    return rotulo


def proximo_mes(mes: int, ano: int):
    if mes == 12:
        return 1, ano + 1
    return mes + 1, ano


# ========================= Duplicar última página ==========================

def duplicar_ultima_pagina(doc, word_app):
    """
    Duplica a ÚLTIMA página e devolve o Range da página nova.
    Usa quebra de seção (Next Page).
    """
    doc.Activate()
    sel = word_app.Selection

    sel.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_LAST)
    ultima_pagina_range = sel.Bookmarks("\\Page").Range
    ultima_pagina_range.Copy()

    sel.EndKey(Unit=WD_STORY)
    sel.InsertBreak(WD_SECTION_BREAK_NEXT_PAGE)
    sel.Paste()

    sel.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_LAST)
    nova_pagina_range = sel.Bookmarks("\\Page").Range

    return nova_pagina_range


# ========================= Excel embutido ==========================

def ler_matriz_de_excel_shape(inline_shape):
    ole = inline_shape.OLEFormat
    ole.Activate()
    excel_obj = ole.Object

    if hasattr(excel_obj, "Worksheets"):
        try:
            ws = excel_obj.ActiveSheet
        except Exception:
            ws = excel_obj.Worksheets(1)
    else:
        ws = excel_obj

    used = ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count

    dados = []
    for r in range(1, rows + 1):
        linha = []
        for c in range(1, cols + 1):
            try:
                cel = used.Cells(r, c)
                valor = cel.Text
                if valor is None:
                    valor = ""
            except Exception:
                valor = ""
            linha.append(str(valor))
        dados.append(linha)

    while dados and not any(str(c).strip() for c in dados[-1]):
        dados.pop()

    return dados


# ========================= Processamento da tabela ==========================

def processar_tabela(dados, valor_novo_str: str, valor_novo_num: Decimal):
    if not dados or len(dados) < 2:
        raise ValueError("Tabela sem linhas suficientes.")

    def linha_vazia(l):
        return not any(str(c).strip() for c in l)

    while dados and linha_vazia(dados[0]):
        dados.pop(0)
    while dados and linha_vazia(dados[-1]):
        dados.pop()

    if len(dados) < 2:
        raise ValueError("Tabela muito pequena após limpeza.")

    total_idx = None
    for i, linha in enumerate(dados):
        texto = str(linha[0] if len(linha) > 0 else "").upper()
        if "TOTAL" in texto:
            total_idx = i
            break
    if total_idx is None:
        raise ValueError("Não encontrei linha 'TOTAL' na primeira coluna.")

    valor_col = None
    linha_total = dados[total_idx]
    for j in range(len(linha_total) - 1, -1, -1):
        txt = str(linha_total[j] or "").strip()
        if not txt:
            continue
        try:
            parse_valor_monetario(txt)
            valor_col = j
            break
        except Exception:
            continue
    if valor_col is None:
        raise ValueError("Não consegui identificar a coluna de valor na linha TOTAL.")

    mes_col = 0
    meses_idxs = []
    for i, linha in enumerate(dados):
        if i == total_idx:
            continue
        if mes_col >= len(linha):
            continue
        txt = str(linha[mes_col] or "").strip()
        if not txt:
            continue
        try:
            parse_mes_ano(txt)
            meses_idxs.append(i)
        except ValueError:
            continue
    if len(meses_idxs) < 2:
        raise ValueError("Não encontrei pelo menos duas linhas de meses.")

    idx_primeiro = meses_idxs[0]
    dados.pop(idx_primeiro)
    if total_idx > idx_primeiro:
        total_idx -= 1

    meses_idxs = []
    for i, linha in enumerate(dados):
        if i == total_idx:
            continue
        if mes_col >= len(linha):
            continue
        txt = str(linha[mes_col] or "").strip()
        if not txt:
            continue
        try:
            parse_mes_ano(txt)
            meses_idxs.append(i)
        except ValueError:
            continue
    if not meses_idxs:
        raise ValueError("Após remover o mês mais antigo, não encontrei mais linhas de meses.")

    idx_ultimo = meses_idxs[-1]
    ultima_linha = dados[idx_ultimo]
    mes_txt = str(ultima_linha[mes_col])
    mes_atual, ano_atual = parse_mes_ano(mes_txt)
    prox_mes, prox_ano = proximo_mes(mes_atual, ano_atual)
    novo_rotulo = format_mes_ano(prox_mes, prox_ano, mes_txt)

    nova_linha = list(ultima_linha)
    ncols = max(len(nova_linha), valor_col + 1, mes_col + 1)
    if len(nova_linha) < ncols:
        nova_linha += [""] * (ncols - len(nova_linha))
    nova_linha[mes_col] = novo_rotulo
    nova_linha[valor_col] = f"R$ {valor_novo_str}"
    dados.insert(idx_ultimo + 1, nova_linha)

    if total_idx > idx_ultimo:
        total_idx += 1

    total_novo_num = Decimal("0.00")
    for i, linha in enumerate(dados):
        if i == total_idx:
            continue
        if mes_col >= len(linha) or valor_col >= len(linha):
            continue
        txt_mes = str(linha[mes_col] or "").strip()
        if not txt_mes:
            continue
        try:
            parse_mes_ano(txt_mes)
        except ValueError:
            continue
        txt_valor = str(linha[valor_col] or "").strip()
        if not txt_valor:
            continue
        total_novo_num += parse_valor_monetario(txt_valor)

    linha_total = dados[total_idx]
    if valor_col >= len(linha_total):
        linha_total += [""] * (valor_col + 1 - len(linha_total))
    linha_total[valor_col] = f"R$ {format_valor_brl(total_novo_num)}"

    for linha in dados:
        for j, cel in enumerate(linha):
            t = str(cel or "").strip()
            if t.startswith("R$"):
                t = re.sub(r"^R\$\s*(.+)$", r"R$ \1", t)
            linha[j] = t

    return dados


# ========================= Tabela Word ==========================

def _ajustar_larguras_colunas(tabela):
    cols = tabela.Columns.Count
    if cols < 2:
        return

    largura_pagina = 0.0
    try:
        doc = tabela.Range.Document
        ps = doc.PageSetup
        largura_pagina = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
    except Exception:
        pass

    total_width = 0.0
    try:
        pw = float(tabela.PreferredWidth or 0)
        if pw > 0:
            total_width = pw
    except Exception:
        pass

    if total_width <= 0:
        try:
            for c in range(1, cols + 1):
                total_width += float(tabela.Columns(c).Width)
        except Exception:
            total_width = 0.0

    if total_width <= 0 and largura_pagina > 0:
        total_width = largura_pagina * 0.55

    if total_width <= 0:
        return

    if cols == 3:
        ratios = [0.22, 0.10, 0.68]
    elif cols == 2:
        ratios = [0.35, 0.65]
    else:
        return

    for i, r in enumerate(ratios, start=1):
        try:
            tabela.Columns(i).Width = total_width * r
        except Exception:
            pass


def _formatar_tabela_visual(tabela):
    tabela.Range.Font.Name = "Arial"
    tabela.Range.Font.Size = 10

    pf = tabela.Range.ParagraphFormat
    pf.SpaceBefore = 0
    pf.SpaceAfter = 0
    pf.LineSpacingRule = 0

    try:
        tabela.Rows.AllowBreakAcrossPages = False
    except Exception:
        pass

    try:
        tabela.Borders.Enable = True
    except Exception:
        pass

    try:
        tabela.TopPadding = 0
        tabela.BottomPadding = 0
        tabela.LeftPadding = 0
        tabela.RightPadding = 0
        tabela.Spacing = 0
    except Exception:
        pass

    try:
        app = tabela.Range.Application
        altura = app.CentimetersToPoints(0.35)
        tabela.Rows.HeightRule = 1  # wdRowHeightAtLeast
        tabela.Rows.Height = altura
        tabela.AllowAutoFit = False
    except Exception:
        pass

    try:
        tabela.Rows.Alignment = 1  # centralizada
    except Exception:
        pass

    _ajustar_larguras_colunas(tabela)


def criar_tabela_word_a_partir_da_matriz(doc, inline_shape, dados):
    if not dados:
        raise ValueError("Matriz de dados vazia.")

    rows = len(dados)
    cols = max(len(linha) for linha in dados)

    rng = inline_shape.Range
    rng.Collapse(0)

    tabela = doc.Tables.Add(rng, rows, cols)

    for r_idx, linha in enumerate(dados, start=1):
        for c_idx in range(1, cols + 1):
            valor = linha[c_idx - 1] if c_idx - 1 < len(linha) else ""
            tabela.Cell(r_idx, c_idx).Range.Text = str(valor)

    for c_idx in range(1, cols + 1):
        tabela.Cell(1, c_idx).Range.Bold = True

    _formatar_tabela_visual(tabela)

    inline_shape.Delete()
    return tabela


def tabela_word_para_matriz(tabela):
    dados = []
    rows = tabela.Rows.Count
    cols = tabela.Columns.Count

    for r in range(1, rows + 1):
        linha = []
        for c in range(1, cols + 1):
            texto = tabela.Cell(r, c).Range.Text
            texto = texto.rstrip("\r\x07")
            linha.append(texto)
        dados.append(linha)

    return dados


def substituir_tabela_word_por_matriz(tabela, dados):
    doc = tabela.Range.Document
    rng = tabela.Range
    tabela.Delete()

    if not dados:
        raise ValueError("Matriz de dados vazia.")

    rows = len(dados)
    cols = max(len(linha) for linha in dados)
    nova = doc.Tables.Add(rng, rows, cols)

    for r_idx, linha in enumerate(dados, start=1):
        for c_idx in range(1, cols + 1):
            valor = linha[c_idx - 1] if c_idx - 1 < len(linha) else ""
            nova.Cell(r_idx, c_idx).Range.Text = str(valor)

    for c_idx in range(1, cols + 1):
        nova.Cell(1, c_idx).Range.Bold = True

    _formatar_tabela_visual(nova)

    return nova


# ========================= Busca da tabela/planilha nova ==========================

def encontrar_tabela_meses_novos(doc, primeiro_indice):
    tabela_encontrada = None
    total = doc.Tables.Count

    for i in range(primeiro_indice, total + 1):
        try:
            tabela = doc.Tables(i)
        except Exception:
            continue

        rows = tabela.Rows.Count
        cols = tabela.Columns.Count
        if cols == 0:
            continue

        achou_total = False
        for r in range(1, rows + 1):
            linha_txt = " ".join(
                tabela.Cell(r, c).Range.Text.rstrip("\r\x07")
                for c in range(1, cols + 1)
            ).upper()
            if "TOTAL" in linha_txt:
                achou_total = True
                break

        if achou_total:
            tabela_encontrada = tabela

    return tabela_encontrada


def encontrar_planilha_excel_nova(doc, inline_inicio, shape_inicio):
    ultimo_tipo = None
    ultimo_obj = None

    total_inline = doc.InlineShapes.Count
    for i in range(inline_inicio, total_inline + 1):
        try:
            ishape = doc.InlineShapes(i)
            ole = ishape.OLEFormat
            class_type = ole.ClassType
        except Exception:
            continue
        if "Excel.Sheet" in str(class_type):
            ultimo_tipo = "inline"
            ultimo_obj = ishape

    try:
        shapes = doc.Shapes
        total_shapes = shapes.Count
    except Exception:
        shapes = None
        total_shapes = 0

    if shapes is not None:
        for i in range(shape_inicio, total_shapes + 1):
            try:
                shp = shapes(i)
                ole = shp.OLEFormat
                class_type = ole.ClassType
            except Exception:
                continue
            if "Excel.Sheet" in str(class_type):
                ultimo_tipo = "shape"
                ultimo_obj = shp

    if ultimo_obj is None:
        return None
    return ultimo_tipo, ultimo_obj


# ========================= Data / "Documento enviado" ==========================

def obter_prefixo_data(doc):
    prefixo = None
    for i in range(1, doc.Paragraphs.Count + 1):
        txt = doc.Paragraphs(i).Range.Text.strip()
        if txt.startswith("Francisco Beltrão"):
            m = re.match(r"(Francisco Beltrão\s*-\s*[^,]*,\s*)", txt)
            if m:
                prefixo = m.group(1)
    if not prefixo:
        prefixo = "Francisco Beltrão - PR, "
    return prefixo


def atualizar_data_francisco(doc, pagina_range, tabela_base,
                             data_extenso: str = "",
                             enviado_para: str = ""):
    """
    Atualiza somente os parágrafos de data e de "Documento enviado para"
    na NOVA página, trocando apenas o texto.
    Não altera alinhamentos, espaçamentos nem outros parágrafos.
    """
    start_min = pagina_range.Start
    end_max = pagina_range.End

    prefixo_default = obter_prefixo_data(doc)

    if data_extenso:
        data_txt = data_extenso.strip()
    else:
        data_txt = data_hoje_extenso()
    if not data_txt.endswith((".", "!", "?")):
        data_txt += "."

    idx_data = None
    idx_env = None

    # Localiza parágrafos de data e de "Documento enviado" na nova página
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        if p.Range.Start < start_min or p.Range.Start > end_max:
            continue
        txt = p.Range.Text.strip()
        if txt.startswith("Francisco Beltrão"):
            idx_data = i
        elif txt.startswith("Documento enviado para"):
            idx_env = i

    # Atualiza texto da data
    if idx_data is not None:
        p_data = doc.Paragraphs(idx_data)
        base_txt = p_data.Range.Text.rstrip("\r\x07")
        m = re.match(r"(Francisco Beltrão\s*-\s*[^,]*,\s*)(.*)", base_txt)
        if m:
            prefixo = m.group(1)
        else:
            prefixo = prefixo_default
        novo_txt = prefixo + data_txt
        p_data.Range.Text = novo_txt + "\r"

    # Atualiza / insere "Documento enviado para"
    enviado_para = (enviado_para or "").strip()
    if enviado_para:
        if idx_env is not None:
            p_env = doc.Paragraphs(idx_env)
            p_env.Range.Text = f"Documento enviado para {enviado_para}.\r"
        else:
            if idx_data is not None:
                # insere um parágrafo logo ANTES da data
                p_data = doc.Paragraphs(idx_data)
                rng = p_data.Range.Duplicate
                rng.Collapse(0)  # início do parágrafo da data
                rng.InsertBefore(f"Documento enviado para {enviado_para}.\r")

    # Apenas garante que a última ocorrência de "Rua Maranhão" esteja centralizada
    try:
        idx_rua = None
        for i in range(1, doc.Paragraphs.Count + 1):
            txt = doc.Paragraphs(i).Range.Text
            if "Rua Maranhão" in txt:
                idx_rua = i
        if idx_rua is not None:
            p = doc.Paragraphs(idx_rua)
            p.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
            p.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    except Exception:
        pass


# ========================= Fluxo principal ==========================

def processar_documento(caminho_doc, valor_novo_usuario, data_extenso, enviado_para):
    if not os.path.isfile(caminho_doc):
        messagebox.showerror("Erro", f"Arquivo não encontrado (Python):\n{caminho_doc}")
        return

    try:
        valor_novo_num, valor_novo_str = interpretar_valor_usuario(valor_novo_usuario)
    except Exception as e:
        messagebox.showerror("Erro", f"Valor monetário inválido:\n{e}")
        return

    try:
        ext = os.path.splitext(caminho_doc)[1] or ".doc"
        tmp_dir = tempfile.mkdtemp(prefix="faturamento_word_")
        tmp_path = os.path.join(tmp_dir, "entrada" + ext)
        shutil.copy2(caminho_doc, tmp_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível criar cópia temporária:\n{e}")
        return

    word_app = win32.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0
    doc = None

    try:
        tmp_path_norm = os.path.normpath(tmp_path)
        doc = word_app.Documents.Open(tmp_path_norm)
        doc.Activate()

        old_tables = doc.Tables.Count
        old_inlines = doc.InlineShapes.Count
        try:
            old_shapes = doc.Shapes.Count
        except Exception:
            old_shapes = 0

        pagina_nova = duplicar_ultima_pagina(doc, word_app)

        tabela = encontrar_tabela_meses_novos(doc, old_tables + 1)

        if tabela is not None:
            dados = tabela_word_para_matriz(tabela)
            dados_novos = processar_tabela(dados, valor_novo_str, valor_novo_num)
            tabela_base = substituir_tabela_word_por_matriz(tabela, dados_novos)
        else:
            resultado_excel = encontrar_planilha_excel_nova(
                doc, old_inlines + 1, old_shapes + 1
            )

            if resultado_excel is None:
                raise RuntimeError(
                    "Não encontrei tabela Word com 'TOTAL' nem planilha Excel embutida na nova página."
                )

            tipo_excel, obj_excel = resultado_excel

            if tipo_excel == "shape":
                try:
                    excel_shape = obj_excel.ConvertToInlineShape()
                except Exception:
                    excel_shape = None
            else:
                excel_shape = obj_excel

            if excel_shape is None:
                raise RuntimeError(
                    "Encontrei objeto Excel, mas não consegui tratá-lo como InlineShape."
                )

            dados_excel = ler_matriz_de_excel_shape(excel_shape)
            dados_novos = processar_tabela(dados_excel, valor_novo_str, valor_novo_num)
            tabela_base = criar_tabela_word_a_partir_da_matriz(doc, excel_shape, dados_novos)

        atualizar_data_francisco(doc, pagina_nova, tabela_base, data_extenso, enviado_para)

        base, ext = os.path.splitext(os.path.normpath(caminho_doc))
        sugestao = base + "_atualizado" + (ext or ".doc")

        novo_caminho = filedialog.asksaveasfilename(
            title="Salvar documento atualizado como",
            initialfile=os.path.basename(sugestao),
            defaultextension=ext or ".doc",
            filetypes=[
                ("Documentos do Word", "*.doc;*.docx"),
                ("Todos os arquivos", "*.*"),
            ],
        )

        if not novo_caminho:
            messagebox.showinfo("Cancelado", "Operação cancelada; arquivo não foi salvo.")
            return

        novo_caminho_norm = os.path.normpath(novo_caminho)

        if os.path.abspath(novo_caminho_norm) == os.path.abspath(caminho_doc):
            resp = messagebox.askyesno(
                "Sobrescrever arquivo?",
                "O caminho escolhido é o mesmo do arquivo original.\n"
                "Deseja sobrescrever o arquivo? O original será substituído."
            )
            if not resp:
                messagebox.showinfo(
                    "Cancelado",
                    "Escolha outro nome ou pasta para não sobrescrever o original."
                )
                return

        doc.SaveAs(novo_caminho_norm)
        messagebox.showinfo("Sucesso", f"Documento salvo em:\n{novo_caminho_norm}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao falar com o Word:\n{e}")
    finally:
        if doc is not None:
            doc.Close(False)
        word_app.Quit()
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


# ========================= UI ==========================

def criar_ui():
    root = tk.Tk()
    root.title("Atualizar tabela de faturamento (.doc/.docx)")

    caminho_var = tk.StringVar()
    valor_var = tk.StringVar()
    data_var = tk.StringVar()
    enviado_var = tk.StringVar()

    def escolher_arquivo():
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Word (.doc ou .docx)",
            filetypes=[
                ("Documentos do Word", "*.doc;*.docx"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if caminho:
            caminho_var.set(caminho)

    def executar():
        caminho = caminho_var.get().strip()
        valor = valor_var.get().strip()
        data_txt = data_var.get().strip()
        enviado_txt = enviado_var.get().strip()

        if not caminho:
            messagebox.showwarning("Atenção", "Selecione o arquivo .doc ou .docx.")
            return
        if not valor:
            messagebox.showwarning("Atenção", "Informe o valor do novo mês.")
            return

        if data_txt:
            try:
                data_extenso = interpretar_data_usuario(data_txt)
                data_var.set(data_extenso)
            except Exception as e:
                messagebox.showerror("Erro", f"Data inválida:\n{e}")
                return
        else:
            data_extenso = ""

        processar_documento(caminho, valor, data_extenso, enviado_txt)

    tk.Label(root, text="Arquivo Word (.doc / .docx):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    tk.Entry(root, textvariable=caminho_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Procurar...", command=escolher_arquivo).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Valor do novo mês:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    tk.Entry(root, textvariable=valor_var, width=20).grid(row=1, column=1, sticky="w", padx=5, pady=5)

    tk.Label(
        root,
        text=(
            "Data após 'Francisco Beltrão - ...' (opcional):\n"
            "Exemplos: 10102025, 10/10/2025, 10-10-25.\n"
            "Será convertido para: 10 de outubro de 2025.\n"
            "Deixe em branco para usar a data de hoje."
        ),
    ).grid(row=2, column=0, sticky="w", padx=5, pady=5)
    tk.Entry(root, textvariable=data_var, width=40).grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Label(root, text="Documento enviado para (opcional):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    tk.Entry(root, textvariable=enviado_var, width=40).grid(row=3, column=1, padx=5, pady=5, sticky="w")

    tk.Button(root, text="Atualizar documento", command=executar).grid(
        row=4, column=0, columnspan=3, pady=10
    )

    root.mainloop()


if __name__ == "__main__":
    criar_ui()
