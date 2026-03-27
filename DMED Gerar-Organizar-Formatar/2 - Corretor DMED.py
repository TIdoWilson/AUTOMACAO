# -*- coding: utf-8 -*-
"""
2 - Corretor DMED

- Para cada par {numero}.slk e {numero}.txt na pasta base:
  - Extrai Nome, CPF/CNPJ e Valor Total Destinatario do .slk
  - Corrige no .txt: CPF/CNPJ, Nome e Valor
  - Se CPF/CNPJ do .txt estiver vazio/zerado ou nao existir no .slk, tenta casar pelo nome

Saida: gera "{numero}_corrigido.txt" ao lado do original.
"""

from __future__ import annotations

import os
import re

# ===================== CONFIGURACAO =====================
BASE_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\DMED Gerar-Organizar-Formatar\arquivos base"
SUFIXO_SAIDA = "_corrigido"
LOG_ALTERACOES = "DMED_Corretor_Alterados.txt"
LOG_ERROS = "DMED_Corretor_Erros_Detectados.txt"
# ========================================================


def _somente_digitos(texto: str) -> str:
    return "".join(ch for ch in texto if ch.isdigit())


def _normalizar_nome(texto: str) -> str:
    texto = texto.upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    return re.sub(r"[^A-Z0-9 ]+", "", texto)


def _valor_para_txt(valor_str: str) -> str:
    # Ex: "4.700,00" -> "470000" / "900,00" -> "90000"
    s = valor_str.strip()
    s = s.replace(".", "").replace(",", "")
    return _somente_digitos(s)


def _extrair_nome_destinatario(texto: str) -> str:
    # Espera algo como "Destinatario: 15461 - NOME"
    if ":" in texto:
        parte = texto.split(":", 1)[1].strip()
    else:
        parte = texto.strip()
    if " - " in parte:
        parte = parte.split(" - ", 1)[1].strip()
    return parte


def _parse_slk(path: str) -> list[dict[str, str]]:
    rows: dict[int, dict[int, str]] = {}

    with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f_in:
        for line in f_in:
            if not line.startswith("C;"):
                continue
            parts = line.rstrip("\r\n").split(";")
            x = y = None
            value = None
            for p in parts[1:]:
                if p.startswith("X"):
                    try:
                        x = int(p[1:])
                    except ValueError:
                        pass
                elif p.startswith("Y"):
                    try:
                        y = int(p[1:])
                    except ValueError:
                        pass
                elif p.startswith("K"):
                    v = p[1:]
                    if v.startswith('"') and v.endswith('"'):
                        v = v[1:-1]
                    value = v
            if x is None or y is None or value is None:
                continue
            rows.setdefault(y, {})[x] = value

    registros: list[dict[str, str]] = []
    atual: dict[str, str] | None = None

    for y in sorted(rows.keys()):
        row = rows[y]

        texto_dest = None
        for v in row.values():
            if isinstance(v, str) and ("Destinat" in v):
                texto_dest = v
                break

        if texto_dest:
            nome = _extrair_nome_destinatario(texto_dest)
            cnpj = ""
            for v in row.values():
                if not isinstance(v, str):
                    continue
                digs = _somente_digitos(v)
                if len(digs) in (11, 14):
                    cnpj = digs
                    break
            atual = {"nome": nome, "cnpj": cnpj}
            continue

        if any(isinstance(v, str) and "Total Destinat" in v for v in row.values()):
            if atual is None:
                continue
            valor = ""
            for v in row.values():
                if not isinstance(v, str):
                    continue
                if re.search(r"\d", v) and ("," in v or "." in v):
                    valor = _valor_para_txt(v)
                    break
            if valor:
                atual["valor"] = valor
                registros.append(atual)
            atual = None

    return registros


def _carregar_mapas(registros: list[dict[str, str]]):
    por_cnpj: dict[str, dict[str, str]] = {}
    por_nome: dict[str, dict[str, str]] = {}

    for r in registros:
        cnpj = r.get("cnpj", "")
        nome = r.get("nome", "")
        if cnpj:
            por_cnpj[cnpj] = r
        if nome:
            por_nome[_normalizar_nome(nome)] = r

    return por_cnpj, por_nome


def _corrigir_txt(txt_path: str, slk_path: str) -> tuple[bool, list[tuple[str, str]], str]:
    registros = _parse_slk(slk_path)
    por_cnpj, por_nome = _carregar_mapas(registros)

    out_path = os.path.splitext(txt_path)[0] + f"{SUFIXO_SAIDA}.txt"
    logs_erros: list[tuple[str, str]] = []
    houve_mudanca = False

    linhas_saida: list[str] = []

    with open(txt_path, "r", encoding="utf-8", errors="ignore", newline="") as f_in:
        for line in f_in:
            line_strip = line.rstrip("\r\n")
            parts = line_strip.split("|")
            if len(parts) < 5 or parts[0] != "RPPSS":
                linhas_saida.append(line_strip)
                continue

            cnpj_txt = _somente_digitos(parts[1])
            nome_txt = parts[2].strip()

            registro = None
            if cnpj_txt and cnpj_txt in por_cnpj:
                registro = por_cnpj[cnpj_txt]
            else:
                key_nome = _normalizar_nome(nome_txt)
                if key_nome in por_nome:
                    registro = por_nome[key_nome]

            if not registro:
                linhas_saida.append(line_strip)
                continue

            original = parts[:]
            parts[1] = registro.get("cnpj", parts[1])
            parts[2] = registro.get("nome", parts[2])
            parts[3] = registro.get("valor", parts[3])

            if parts != original:
                houve_mudanca = True
                logs_erros.append(("|".join(original), "|".join(parts)))

            linhas_saida.append("|".join(parts))

    if logs_erros:
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            with open(os.path.join(script_dir, LOG_ERROS), "a", encoding="utf-8") as f_log:
                for antes, depois in logs_erros:
                    f_log.write(f"{antes} -> {depois}\n")
        except Exception as exc:
            print(f"Falha ao registrar erros detectados: {exc}")

    if houve_mudanca:
        with open(out_path, "w", encoding="utf-8", newline="") as f_out:
            for ln in linhas_saida:
                f_out.write(ln + "\n")

    return houve_mudanca, logs_erros, out_path


def main() -> None:
    if not os.path.isdir(BASE_DIR):
        print(f"Pasta nao encontrada: {BASE_DIR}")
        return

    arquivos = os.listdir(BASE_DIR)
    slks = {os.path.splitext(a)[0]: a for a in arquivos if a.lower().endswith(".slk")}
    txts = {os.path.splitext(a)[0]: a for a in arquivos if a.lower().endswith(".txt")}

    pares = sorted(set(slks.keys()) & set(txts.keys()))
    if not pares:
        print("Nenhum par .slk/.txt encontrado.")
        return

    for base in pares:
        slk_path = os.path.join(BASE_DIR, slks[base])
        txt_path = os.path.join(BASE_DIR, txts[base])
        mudou, logs, saida = _corrigir_txt(txt_path, slk_path)
        print(f"OK: {os.path.basename(txt_path)}")
        if mudou:
            try:
                base_nome = os.path.splitext(os.path.basename(txt_path))[0]
                with open(os.path.join(BASE_DIR, LOG_ALTERACOES), "a", encoding="utf-8") as f_log:
                    for antes, depois in logs:
                        f_log.write(f"{base_nome}\t{antes} -> {depois}\n")
            except Exception as exc:
                print(f"Falha ao registrar log: {exc}")
            print(f"CRIADO: {os.path.basename(saida)}")


if __name__ == "__main__":
    main()
