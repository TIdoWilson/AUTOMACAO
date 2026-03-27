#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Extrai lançamentos de um PDF Sicoob (baixas), seleciona campos importantes,
separa campos da coluna 'Cliente' em colunas adicionais e
gera Excel com as abas 'Detalhe' e 'Resumo por Cliente'.

- NÃO altera as funções extrair_registros e parsear_registros (versão que funcionou).
- Pede ao usuário para localizar o PDF via diálogo (tkinter) com fallback para console.
- Formata colunas monetárias e ativa autofiltro/congelamento de painéis.

Autor: você :)
"""

import re
import sys
import os
from datetime import datetime

# Dependências
try:
    import pdfplumber
    import pandas as pd
except ImportError as e:
    print("Erro: falta dependência. Instale com: pip install pdfplumber pandas XlsxWriter")
    sys.exit(1)

# ---- Configurações de parsing (ajuste se necessário) ----
ACCOUNT_REGEX = re.compile(r"\bSCO_[A-Z0-9_]+\b")
DATE_REGEX = re.compile(r"\d{2}/\d{2}/\d{4}")
MONEY_REGEX = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")
TOTAL_CLIENTE_REGEX = re.compile(r"Total do cliente", re.IGNORECASE)
DATE_RE_BOUNDS = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")

# ---- Funções utilitárias ----
def clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def br_money_to_float(s):
    if isinstance(s, str):
        return float(s.replace(".", "").replace(",", "."))
    return s

def escolher_arquivo():
    """
    Abre um diálogo para o usuário escolher o PDF.
    Se tkinter não estiver disponível (ou ambiente sem GUI), pede no console.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title="Selecione o PDF do Sicoob",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
        )
        root.update()
        root.destroy()
        if not path:
            print("Nenhum arquivo selecionado.")
            sys.exit(0)
        return path
    except Exception:
        # fallback: console
        print("Não foi possível abrir o seletor gráfico de arquivos.")
        path = input("Digite/cole o caminho do PDF: ").strip().strip('"')
        if not os.path.isfile(path):
            print("Caminho inválido ou arquivo não encontrado.")
            sys.exit(1)
        return path

# ---- EXTRAÇÃO/Parsing (VERSÃO QUE FUNCIONOU) ----
def extrair_registros(pdf_path: str):
    """
    Percorre as páginas e costura linhas de cada lançamento.
    Retorna uma lista de strings (cada uma é um lançamento bruto).

    Lógica robusta:
    - Acumula linhas até alcançar >= 6 valores monetários.
    - Faz flush ao detectar uma nova linha que começa por data.
    - Ignora linhas de 'Total do cliente'.
    """
    registros = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            text = text.replace("\xa0", " ")
            linhas = text.split("\n")

            buffer = ""
            for linha in linhas:
                linha = clean_spaces(linha)
                if TOTAL_CLIENTE_REGEX.search(linha):
                    buffer = ""
                    continue

                if re.match(r"^\d{2}/\d{2}/\d{4}\b", linha) and buffer:
                    # nova data -> flush anterior, se estiver completo
                    if len(MONEY_REGEX.findall(buffer)) >= 6:
                        registros.append(buffer.strip())
                    buffer = linha
                else:
                    buffer += (" " if buffer else "") + linha

                # flush automático se já tem 6+ valores monetários
                if len(MONEY_REGEX.findall(buffer)) >= 6:
                    registros.append(buffer.strip())
                    buffer = ""

            # flush final da página
            if len(MONEY_REGEX.findall(buffer)) >= 6:
                registros.append(buffer.strip())
    return registros

def parsear_registros(registros):
    """
    Transforma cada string bruta em um dicionário com os campos desejados.
    Mantida a versão que funcionou após o ajuste.
    """
    parsed = []
    for rec in registros:
        try:
            amounts = MONEY_REGEX.findall(rec)
            if len(amounts) < 6:
                continue

            # últimos 6 campos monetários (ordem baseada no layout observado)
            vl_baixa, acrescimo, seguro, taxa_adm, desconto, liquido = amounts[-6:]

            # conta corrente
            acc_match = ACCOUNT_REGEX.search(rec)
            conta_corrente = acc_match.group(0) if acc_match else ""

            # datas
            datas = DATE_REGEX.findall(rec)
            dt_baixa = datas[0] if datas else ""
            data_vecto = datas[-1] if len(datas) > 1 else ""

            # cliente: trecho entre dt_baixa e o primeiro valor monetário
            first_money = rec.find(amounts[0])
            cliente = clean_spaces(rec)
            if dt_baixa:
                start = cliente.find(dt_baixa) + len(dt_baixa)
                cliente = cliente[start:first_money] if first_money != -1 else cliente[start:]
            cliente = re.sub(r"\(\d+\)", "", cliente).strip(" -")

            parsed.append({
                "Dt. baixa": dt_baixa,
                "Cliente": cliente,
                "Data vecto": data_vecto,
                "Vl. baixa": vl_baixa,
                "Acréscimo": acrescimo,
                "Líquido": liquido,
                "Conta corrente": conta_corrente
            })
        except Exception:
            continue
    return pd.DataFrame(parsed)

# ---- NOVA FUNÇÃO: expandir campos a partir da coluna "Cliente" ----
def expandir_campos_cliente(df: pd.DataFrame) -> pd.DataFrame:
    """
    A partir de 'Cliente', cria as colunas:
      - nome do cliente, data, documento, titulo, parcela, TC, Unid. princ, port, oper, data vecto
    **Não remove nenhuma** das colunas criadas.
    """

    def parse_cliente(texto: str):
        """
        Divide 'Cliente' em:
        nome do cliente | documento | titulo | parcela | TC | Unid. princ | port | oper
        assumindo que 'Cliente' NÃO tem datas e que os 7 últimos tokens são os campos.
        """
        if not isinstance(texto, str):
            return {"nome do cliente": "", "documento": "", "titulo": "", "parcela": "",
                    "TC": "", "Unid. princ": "", "port": "", "oper": "", "data": "", "data vecto": ""}

        tokens = [t for t in texto.split() if t]

        # guardas
        if len(tokens) < 7:
            # não tem tokens suficientes – considere tudo como nome
            return {"nome do cliente": texto.strip(), "documento": "", "titulo": "", "parcela": "",
                    "TC": "", "Unid. princ": "", "port": "", "oper": "", "data": "", "data vecto": ""}

        # heurística a partir do fim:
        oper        = tokens[-1]
        port        = tokens[-2]
        unid_princ  = tokens[-3]
        tc          = tokens[-4]
        parcela     = tokens[-5]
        titulo      = tokens[-6]
        documento   = tokens[-7]
        nome        = " ".join(tokens[:-7]).strip()

        return {
            "nome do cliente": nome,
            "documento": documento,
            "titulo": titulo,
            "parcela": parcela,
            "TC": tc,
            "Unid. princ": unid_princ,
            "port": port,
            "oper": oper,
            # mantidos vazios porque as datas não estão mais em 'Cliente'
            "data": "",
            "data vecto": ""
        }
    
    parsed_rows = df["Cliente"].map(parse_cliente).tolist()
    parsed_df = pd.DataFrame(parsed_rows, index=df.index)
    df_out = pd.concat([df, parsed_df], axis=1)
    return df_out

# ---- Exportar para Excel ----
def salvar_excel(df: pd.DataFrame, caminho_sugestao: str = None) -> str:
    """
    Gera o Excel com 'Detalhe' e 'Resumo por Cliente' formatados.
    Retorna o caminho do arquivo salvo.

    Observação: se existir a coluna 'nome do cliente', ela é priorizada na aba Detalhe.
    """
    # Colunas numéricas para somar
    for col in ["Vl. baixa", "Acréscimo", "Líquido"]:
        df[col + " (num)"] = df[col].map(br_money_to_float)

    # remove duplicados óbvios
    df = df.drop_duplicates(subset=["Cliente", "Dt. baixa", "Data vecto", "Vl. baixa", "Acréscimo", "Líquido"])

    # resumo por cliente
    summary = (
        df.groupby("Cliente", as_index=False)[["Vl. baixa (num)", "Acréscimo (num)", "Líquido (num)"]]
        .sum()
        .sort_values("Líquido (num)", ascending=False)
    )

    # caminho de saída
    if not caminho_sugestao:
        caminho_sugestao = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            f"Contas_Recebidas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

    # Decide quais colunas mostrar na planilha "Detalhe"
    cols_detalhe = [
        "Dt. baixa",
        "nome do cliente" if "nome do cliente" in df.columns else "Cliente",
        "Unid. princ",            # incluída
        "Vl. baixa",
        "Acréscimo",
        "Líquido",
        "Conta corrente",
    ]
    cols_detalhe = [c for c in cols_detalhe if c in df.columns]


    cols_detalhe = [c for c in cols_detalhe if c in df.columns]  # tolerante caso falte algo

    # exporta para Excel
    with pd.ExcelWriter(caminho_sugestao, engine="xlsxwriter") as writer:
        # Detalhe
        df_export = df[cols_detalhe].copy()
        df_export.to_excel(writer, index=False, sheet_name="Detalhe")
        workbook = writer.book
        ws_det = writer.sheets["Detalhe"]

        header_fmt = workbook.add_format({"bold": True, "border": 1})
        money_fmt = workbook.add_format({"num_format": u'R$ #,##0.00'})
        # datas ficam como texto para evitar problemas de localidade; ajuste se quiser converter
        # Ajuste de larguras
        if "Dt. baixa" in cols_detalhe:
            ws_det.set_column(cols_detalhe.index("Dt. baixa"), cols_detalhe.index("Dt. baixa"), 12)
        if ("nome do cliente" in cols_detalhe) or ("Cliente" in cols_detalhe):
            colname = "nome do cliente" if "nome do cliente" in cols_detalhe else "Cliente"
            ws_det.set_column(cols_detalhe.index(colname), cols_detalhe.index(colname), 45)
        if "Data vecto" in cols_detalhe:
            ws_det.set_column(cols_detalhe.index("Data vecto"), cols_detalhe.index("Data vecto"), 12)

        # Formatação monetária
        for money_col in ["Vl. baixa", "Acréscimo", "Líquido"]:
            if money_col in cols_detalhe:
                idx = cols_detalhe.index(money_col)
                ws_det.set_column(idx, idx, 14, money_fmt)

        if "Conta corrente" in cols_detalhe:
            idx = cols_detalhe.index("Conta corrente")
            ws_det.set_column(idx, idx, 18)

        # cabeçalhos
        for col_idx, col_name in enumerate(df_export.columns):
            ws_det.write(0, col_idx, col_name, header_fmt)

        ws_det.autofilter(0, 0, len(df_export), df_export.shape[1] - 1)
        ws_det.freeze_panes(1, 0)

        # Resumo por Cliente
        summary_export = summary.rename(columns={
            "Vl. baixa (num)": "Vl. baixa",
            "Acréscimo (num)": "Acréscimo",
            "Líquido (num)": "Líquido",
        })
        summary_export.to_excel(writer, index=False, sheet_name="Resumo por Cliente")
        ws_sum = writer.sheets["Resumo por Cliente"]
        for col_idx, col_name in enumerate(summary_export.columns):
            ws_sum.write(0, col_idx, col_name, header_fmt)
        ws_sum.set_column("A:A", 45)
        ws_sum.set_column("B:D", 16, money_fmt)
        ws_sum.autofilter(0, 0, len(summary_export), summary_export.shape[1] - 1)
        ws_sum.freeze_panes(1, 0)

    return caminho_sugestao

# ---- Pipeline principal ----
def processar_pdf(pdf_path: str) -> str:
    print(f"Lendo: {pdf_path}")
    registros = extrair_registros(pdf_path)
    if not registros:
        raise RuntimeError("Nenhum lançamento encontrado. Verifique se o PDF é o esperado.")
    print(f"Registros brutos extraídos: {len(registros)}")

    df = parsear_registros(registros)
    if df.empty:
        raise RuntimeError("Não foi possível parsear os lançamentos.")
    print(f"Lançamentos parseados: {len(df)}")

    # >>> NOVO PASSO: expandir campos da coluna 'Cliente' conforme solicitado
    df = expandir_campos_cliente(df)

    # Sugere saída ao lado do PDF
    saida = os.path.join(
        os.path.dirname(pdf_path),
        f"Contas_Recebidas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    out_path = salvar_excel(df, saida)
    return out_path

# ---- Execução ----
def main():
    try:
        pdf_path = escolher_arquivo()
        out_path = processar_pdf(pdf_path)
        print(f"Pronto! Arquivo gerado:\n{out_path}")
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
