#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Combinador de SPED (EFD Contribuições / SPED em geral) — Matriz + Filiais

Atende:
- GUI (Tkinter): seletor de Matriz (1 arquivo) + Filiais (1+ arquivos).
- Saída padronizada: mesmo nome do arquivo da Matriz, salva em:
  W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Combinar SPEDs Matriz + Filiais\Combinados

Fluxo (alto nível):
  1) Lê Matriz e Filiais.
  2) Pré-check e auto-ajustes mínimos:
     - garante '|' final em todas as linhas (quando faltar).
     - corrige D500 com campo final ausente adicionando pipes finais (padding).
  3) Remove linhas idênticas da(s) filial(is) que já existem na Matriz (dedupe inter-arquivos).
  4) Mescla Bloco 0 (0150/0190/0200/0400/0450/0500/0600/0140) por CHAVE (campo 1),
     inserindo no final da “seção” e mantendo 0990 como última linha do bloco (recalculado).
     - 0200 preserva/mescla filhos 0205/0206/0220
     - 0150 mescla 0175 se existir
     - BUGFIX: evita "list index out of range" quando 2+ filiais trazem a mesma chave inexistente na matriz
  5) Mescla demais blocos (A/C/D/F/I/1) preservando abertura X001 da Matriz, recalculando X990.
  6) Bloco M (corrigido):
     - Mantém integridade de grupos pai/filhos (ex.: M100 + filhos; M500 + filhos; M610 + filhos),
       evitando hierarquia inválida (ex.: "esperado M611 mas veio M100").
     - Consolida M200 e M600 (únicos) somando campos numéricos; mantém COD_CONT do primeiro.
     - Consolida M100 e M500 por COD_CONT (quando existirem em mais de um arquivo), mesclando filhos sem duplicar linhas exatas.
     - Mantém M400/M410 como grupos e garante que grupos M400 fiquem antes de M500.
     - Recalcula M990.
  7) Regera Bloco 9 (9001/9900/9990/9999) com contagens corretas.
  8) Evita reformatar linhas: só reconstrói quando necessário (aberturas/fechamentos/Bloco 9 e consolidações).
"""

from __future__ import annotations

import argparse
import os
import re
from collections import OrderedDict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Optional, Tuple


# -----------------------------
# Config
# -----------------------------

DEFAULT_OUT_DIR = (
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Combinar SPEDs Matriz + Filiais\Combinados"
)

CANONICAL_BLOCK_ORDER = ["0", "A", "C", "D", "F", "I", "M", "1", "9"]

CLOSE_REG_BY_BLOCK = {
    "0": "0990",
    "A": "A990",
    "C": "C990",
    "D": "D990",
    "F": "F990",
    "I": "I990",
    "M": "M990",
    "1": "1990",
    "9": "9990",  # (9999 é fechamento do arquivo)
}

OPEN_REG_BY_BLOCK = {
    "A": "A001",
    "C": "C001",
    "D": "D001",
    "F": "F001",
    "I": "I001",
    "M": "M001",
    "1": "1001",
    "9": "9001",
    "0": "0001",  # especial
}

# Bloco 0: registros que devem ser “puxados” das filiais e encaixados no bloco 0 da matriz por chave (campo 1)
BLOCK0_MERGE_REGS = {"0140", "0150", "0190", "0200", "0400", "0450", "0500", "0600"}

# Registros do bloco 0 que NÃO devem entrar vindos das filiais (cabeçalhos da filial)
BLOCK0_SKIP_FROM_FILIAL = {
    "0000",
    "0001",
    "0005",
    "0100",
    "0110",
    "0111",
    "0120",
    "0130",
    "0135",
    "0990",
}

BLOCK0_CHILDREN = {
    "0150": {"0175"},
    "0200": {"0205", "0206", "0220"},
}

REG_RE = re.compile(r"^\|+([A-Z0-9]{4})\|")

# Correções pontuais de “min fields” (quantidade mínima de campos após o REG).
MIN_FIELDS_AFTER_REG: Dict[str, int] = {
    "D500": 22,  # campos após REG (IND_OPER..COD_CTA) = 22
}


# -----------------------------
# Utilidades de parsing
# -----------------------------

def detect_encoding(path: str) -> str:
    """
    Detecta encoding de forma simples (sem estourar memória).
    """
    sample_size = 256 * 1024  # 256KB
    with open(path, "rb") as f:
        data = f.read(sample_size)
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            data.decode(enc)
            return enc
        except UnicodeDecodeError:
            pass
    return "latin-1"


def read_lines(path: str, encoding: str) -> List[str]:
    """
    Lê linha a linha (evita duplicar memória).
    """
    out: List[str] = []
    with open(path, "r", encoding=encoding, errors="strict", newline="") as f:
        for ln in f:
            ln = ln.rstrip("\r\n")
            if ln != "":
                out.append(ln)
    return out


def get_reg(line: str) -> str:
    m = REG_RE.match(line.lstrip("\ufeff").strip())
    return m.group(1) if m else ""


def get_block(reg: str) -> str:
    return reg[:1] if reg else ""


def split_fields(line: str) -> Tuple[str, List[str]]:
    """
    Retorna (REG, FIELDS).
    Ex: '|0150|ABC|NOME|...|' -> ('0150', ['ABC','NOME',...])
    """
    if not line.startswith("|"):
        return "", []
    parts = line.split("|")
    if len(parts) < 3:
        return "", []
    reg = parts[1]
    fields = parts[2:-1]  # descarta vazio inicial e vazio final
    return reg, fields


def build_line(reg: str, fields: List[str]) -> str:
    return "|" + reg + "|" + "|".join("" if f is None else str(f) for f in fields) + "|"


def replace_first_field(line: str, new_value: str) -> str:
    """
    Para registros de abertura X001 (ou 0001), troca o primeiro campo (IND_MOV).
    """
    reg, fields = split_fields(line)
    if not reg:
        return line
    if not fields:
        fields = [new_value]
    else:
        fields[0] = new_value
    return build_line(reg, fields)


def parse_blocks(lines: List[str]) -> Dict[str, List[str]]:
    blocks: Dict[str, List[str]] = {}
    for ln in lines:
        reg = get_reg(ln)
        if not reg:
            continue
        b = get_block(reg)
        blocks.setdefault(b, []).append(ln)
    return blocks


def strip_block_close(lines: List[str], close_reg: str) -> List[str]:
    if not lines:
        return []
    # remove a ÚLTIMA ocorrência do fechamento
    for i in range(len(lines) - 1, -1, -1):
        if get_reg(lines[i]) == close_reg:
            return lines[:i] + lines[i + 1 :]
    return lines[:]


def find_last_index(lines: List[str], reg: str) -> int:
    for i in range(len(lines) - 1, -1, -1):
        if get_reg(lines[i]) == reg:
            return i
    return -1


def count_block_lines_with_close(content_with_open: List[str], close_reg: str) -> List[str]:
    """
    Recebe linhas do bloco já com abertura (quando aplicável), sem fechamento.
    Recalcula e adiciona o fechamento X990 com QTD_LIN_X correto.
    """
    qtd = len(content_with_open) + 1  # inclui o próprio fechamento
    return content_with_open + [build_line(close_reg, [str(qtd)])]


def _safe_int(s: str) -> Optional[int]:
    try:
        return int((s or "").strip())
    except Exception:
        return None


# -----------------------------
# Pré-check + auto-ajustes mínimos
# -----------------------------

def sanitize_lines(lines: List[str], source_label: str, log: List[str]) -> List[str]:
    """
    Ajustes mínimos e seguros (sem reformatar agressivo):
    - garante '|' final (quando faltar).
    - padding de campos finais para alguns registros (ex.: D500).
    """
    out: List[str] = []
    for idx, ln in enumerate(lines):
        if not ln:
            continue

        if not ln.endswith("|"):
            ln = ln + "|"
            log.append(f"[{source_label}] Ajuste: linha {idx+1} sem '|' final -> corrigido.")

        reg = get_reg(ln)
        if reg in MIN_FIELDS_AFTER_REG:
            _, fields = split_fields(ln)
            need = MIN_FIELDS_AFTER_REG[reg]
            if len(fields) < need:
                missing = need - len(fields)
                # cada '|' adicional representa um campo vazio no final
                ln = ln + ("|" * missing)
                log.append(
                    f"[{source_label}] Ajuste: {reg} com campos insuficientes -> "
                    f"adicionados {missing} campo(s) vazio(s) no final."
                )

        out.append(ln)
    return out


def precheck_file(lines: List[str], source_label: str, log: List[str]) -> None:
    """
    Checagens básicas (não bloqueia; só loga):
    - presença de 0000/0001 e 9999
    - consistência de contagens em fechamentos X990 (quando parseável)
    - consistência do 9999 (quando parseável)
    """
    regs = [get_reg(x) for x in lines if x]
    reg_set = set(regs)

    for must in ("0000", "0001", "9999"):
        if must not in reg_set:
            log.append(f"[{source_label}] AVISO: registro obrigatório ausente: {must}")

    blocks = parse_blocks(lines)
    for b, close_reg in CLOSE_REG_BY_BLOCK.items():
        if b == "9":
            continue
        blines = blocks.get(b, [])
        if not blines:
            continue
        idx_close = find_last_index(blines, close_reg)
        if idx_close < 0:
            log.append(f"[{source_label}] AVISO: bloco {b} sem fechamento {close_reg}.")
            continue
        _, f = split_fields(blines[idx_close])
        qtd_in = _safe_int(f[0] if f else "")
        if qtd_in is None:
            log.append(f"[{source_label}] AVISO: {close_reg} com QTD_LIN não numérico.")
            continue
        qtd_real = len(blines)  # inclui o fechamento
        if qtd_in != qtd_real:
            log.append(
                f"[{source_label}] AVISO: {close_reg} QTD_LIN={qtd_in} mas linhas reais do bloco={qtd_real}."
            )

    idx_9999 = find_last_index(lines, "9999")
    if idx_9999 >= 0:
        _, f = split_fields(lines[idx_9999])
        qtd_in = _safe_int(f[0] if f else "")
        if qtd_in is not None and qtd_in != len(lines):
            log.append(
                f"[{source_label}] AVISO: 9999 QTD_LIN={qtd_in} mas linhas reais do arquivo={len(lines)}."
            )


def dedupe_filial_against_matriz(
    matriz_lines: List[str],
    filial_lines: List[str],
    source_label: str,
    log: List[str],
) -> List[str]:
    """
    Remove da filial qualquer linha idêntica que já exista na matriz.
    """
    mset = set(matriz_lines)
    before = len(filial_lines)
    out = [ln for ln in filial_lines if ln not in mset]
    removed = before - len(out)
    if removed:
        log.append(f"[{source_label}] Dedupe: removidas {removed} linha(s) idêntica(s) já presentes na Matriz.")
    return out


# -----------------------------
# Bloco 0: estrutura por “itens” (chave)
# -----------------------------

@dataclass
class Block0Item:
    reg: str
    key: Optional[str]
    lines: List[str]


def parse_block0_items(lines_without_0990: List[str]) -> List[Block0Item]:
    items: List[Block0Item] = []
    i = 0
    n = len(lines_without_0990)

    while i < n:
        ln = lines_without_0990[i]
        reg = get_reg(ln)

        if reg in BLOCK0_CHILDREN:
            parent_reg = reg
            parent_line = ln
            _, fields = split_fields(parent_line)
            key = fields[0] if fields else ""
            group = [parent_line]
            i += 1
            while i < n:
                nxt = lines_without_0990[i]
                nxt_reg = get_reg(nxt)
                if nxt_reg in BLOCK0_CHILDREN.get(parent_reg, set()):
                    group.append(nxt)
                    i += 1
                else:
                    break
            items.append(Block0Item(reg=parent_reg, key=key, lines=group))
            continue

        if reg in BLOCK0_MERGE_REGS:
            _, fields = split_fields(ln)
            key = fields[0] if fields else ""
            items.append(Block0Item(reg=reg, key=key, lines=[ln]))
            i += 1
            continue

        items.append(Block0Item(reg=reg, key=None, lines=[ln]))
        i += 1

    return items


def merge_block0(matrix_lines: List[str], filial_blocks0: List[List[str]], log: List[str]) -> List[str]:
    """
    Mescla bloco 0 com BUGFIX:
    - Mantém cabeçalhos da Matriz.
    - Puxa das filiais apenas registros BLOCK0_MERGE_REGS (exceto cabeçalho).
    - Se a MESMA chave aparecer em 2+ filiais e não existir na Matriz, não quebra:
      ela é tratada como "pendente já adicionado" (pending_map).
    - 0200 e 0150 mesclam filhos.
    - Recalcula 0990.
    """
    close_reg = CLOSE_REG_BY_BLOCK["0"]

    base = strip_block_close(matrix_lines, close_reg)
    base_items = parse_block0_items(base)

    matrix_idx_map: Dict[Tuple[str, str], int] = {}
    for idx, it in enumerate(base_items):
        if it.key is not None and it.reg in BLOCK0_MERGE_REGS and it.key != "":
            matrix_idx_map[(it.reg, it.key)] = idx

    pending_by_reg: Dict[str, List[Block0Item]] = {r: [] for r in BLOCK0_MERGE_REGS}
    pending_map: Dict[Tuple[str, str], Block0Item] = {}

    def last_index_of_reg(items: List[Block0Item], reg: str) -> int:
        for j in range(len(items) - 1, -1, -1):
            if items[j].reg == reg and items[j].key is not None:
                return j
        return -1

    def merge_children_into(target: Block0Item, incoming: Block0Item) -> int:
        existing = set(target.lines)
        added = 0
        for child_ln in incoming.lines[1:]:
            if child_ln not in existing:
                target.lines.append(child_ln)
                existing.add(child_ln)
                added += 1
        return added

    for filial0 in filial_blocks0:
        filial0 = strip_block_close(filial0, close_reg)
        filial_items = parse_block0_items(filial0)

        for it in filial_items:
            if it.reg in BLOCK0_SKIP_FROM_FILIAL:
                continue
            if it.reg not in BLOCK0_MERGE_REGS:
                continue
            if not it.key:
                continue

            k = (it.reg, it.key)

            # 1) Já existe na Matriz
            if k in matrix_idx_map:
                base_idx = matrix_idx_map[k]
                base_item = base_items[base_idx]
                if it.reg in BLOCK0_CHILDREN:
                    added = merge_children_into(base_item, it)
                    if added:
                        log.append(
                            f"[Bloco 0] Mesclado {added} linha(s) filha(s) em {it.reg} chave={it.key} (na Matriz)."
                        )
                else:
                    if it.lines[0] != base_item.lines[0]:
                        log.append(f"[Bloco 0] Ignorado (chave já existe na Matriz) {it.reg} chave={it.key}.")
                continue

            # 2) Já está pendente (veio de outra filial) -> BUGFIX
            if k in pending_map:
                prev = pending_map[k]
                if it.reg in BLOCK0_CHILDREN:
                    added = merge_children_into(prev, it)
                    if added:
                        log.append(
                            f"[Bloco 0] Mesclado {added} linha(s) filha(s) em {it.reg} chave={it.key} (entre filiais)."
                        )
                else:
                    if it.lines[0] != prev.lines[0]:
                        log.append(f"[Bloco 0] Duplicado entre filiais ignorado: {it.reg} chave={it.key}.")
                continue

            # 3) Novo pendente
            pending_by_reg[it.reg].append(it)
            pending_map[k] = it
            log.append(f"[Bloco 0] Adicionado novo {it.reg} chave={it.key} ({len(it.lines)} linha(s)).")

    rebuilt: List[Block0Item] = base_items[:]
    order_hint = ["0140", "0150", "0190", "0200", "0400", "0450", "0500", "0600"]

    for reg in order_hint:
        if not pending_by_reg.get(reg):
            continue

        pos = last_index_of_reg(rebuilt, reg)
        if pos >= 0:
            insert_at = pos + 1
        else:
            insert_at = len(rebuilt)
            reg_idx = order_hint.index(reg)
            for later in order_hint[reg_idx + 1 :]:
                for j, it in enumerate(rebuilt):
                    if it.reg == later and it.key is not None:
                        insert_at = j
                        break
                if insert_at != len(rebuilt):
                    break

        new_items = pending_by_reg[reg]
        if new_items:
            rebuilt[insert_at:insert_at] = new_items

    out0_no_close: List[str] = []
    for it in rebuilt:
        out0_no_close.extend(it.lines)

    idx_0001 = find_last_index(out0_no_close, "0001")
    if idx_0001 >= 0:
        _, f = split_fields(out0_no_close[idx_0001])
        if f and f[0] != "0":
            out0_no_close[idx_0001] = replace_first_field(out0_no_close[idx_0001], "0")

    qtd = len(out0_no_close) + 1
    return out0_no_close + [build_line(close_reg, [str(qtd)])]


# -----------------------------
# Mescla simples (blocos A/C/D/F/I/1)
# -----------------------------

def merge_simple_block(
    block_char: str,
    matrix_lines: List[str],
    filial_blocks: List[List[str]],
    log: List[str],
) -> List[str]:
    """
    Mescla blocos com abertura X001 e fechamento X990:
    - Mantém abertura da Matriz (ou da primeira filial se Matriz não tiver).
    - Ignora aberturas das filiais.
    - Remove duplicatas por linha exata.
    - Recalcula X990.
    - Ajusta IND_MOV para 0 se houver conteúdo além da abertura.
    """
    close_reg = CLOSE_REG_BY_BLOCK[block_char]
    open_reg = OPEN_REG_BY_BLOCK.get(block_char, "")

    base = matrix_lines[:] if matrix_lines else []
    if not base:
        for fb in filial_blocks:
            if fb:
                base = fb[:]
                break

    base_wo_close = strip_block_close(base, close_reg)

    open_line: Optional[str] = None
    content: List[str] = []
    for ln in base_wo_close:
        r = get_reg(ln)
        if r == open_reg and open_line is None:
            open_line = ln
        elif r == open_reg:
            continue
        else:
            content.append(ln)

    existing = set(([open_line] if open_line else []) + content)
    additions: List[str] = []

    for fb in filial_blocks:
        if not fb:
            continue
        fb_wo_close = strip_block_close(fb, close_reg)
        for ln in fb_wo_close:
            r = get_reg(ln)
            if r in (open_reg, close_reg):
                continue
            if ln in existing:
                continue
            existing.add(ln)
            additions.append(ln)

    merged_wo_close: List[str] = []

    if open_line is None:
        open_line = build_line(open_reg, ["0" if (content or additions) else "1"])

    merged_wo_close.append(open_line)
    merged_wo_close.extend(content)
    merged_wo_close.extend(additions)

    has_content = len(merged_wo_close) > 1
    merged_wo_close[0] = replace_first_field(merged_wo_close[0], "0" if has_content else "1")

    if additions:
        log.append(f"[Bloco {block_char}] Adicionadas {len(additions)} linha(s).")

    return count_block_lines_with_close(merged_wo_close, close_reg)


# -----------------------------
# Bloco M (especial) - corrigido
# -----------------------------

def _to_decimal(v: str) -> Decimal:
    s = (v or "").strip()
    if s == "":
        return Decimal("0")
    s = s.replace(".", "")
    s = s.replace(",", ".")
    return Decimal(s)


def _is_numeric_like(v: str) -> bool:
    s = (v or "").strip()
    if s == "":
        return True
    return bool(re.fullmatch(r"-?\d+(,\d+)?", s) or re.fullmatch(r"-?\d+(\.\d+)?", s))


def _format_decimal_sped(d: Decimal, places: int) -> str:
    if places <= 0:
        return f"{d.quantize(Decimal('1')):f}".replace(".", ",")
    q = Decimal("1." + ("0" * places))
    dd = d.quantize(q)
    return f"{dd:f}".replace(".", ",")


def sum_unique_record(reg: str, occurrences: List[str], log: List[str]) -> Optional[str]:
    """
    Mantém o primeiro registro como base (especialmente COD_CONT) e soma campos numéricos (a partir do 2º campo).
    Preserva a linha original quando existir apenas 1 ocorrência (não re-formata).
    """
    if not occurrences:
        return None
    if len(occurrences) == 1:
        return occurrences[0]

    base_reg, _ = split_fields(occurrences[0])
    if base_reg != reg:
        return occurrences[0]

    max_len = max(len(split_fields(x)[1]) for x in occurrences)

    all_fields: List[List[str]] = []
    max_places_by_idx: Dict[int, int] = {}

    for ln in occurrences:
        _, f = split_fields(ln)
        f = f + [""] * (max_len - len(f))
        all_fields.append(f)
        for i, val in enumerate(f):
            if i == 0:
                continue
            s = (val or "").strip()
            if "," in s:
                max_places_by_idx[i] = max(max_places_by_idx.get(i, 0), len(s.split(",")[1]))
            else:
                max_places_by_idx[i] = max(max_places_by_idx.get(i, 0), 0)

    out_fields = list(all_fields[0])

    for i in range(1, max_len):
        if not all(_is_numeric_like(f[i]) for f in all_fields):
            continue
        total = Decimal("0")
        ok = True
        for f in all_fields:
            try:
                total += _to_decimal(f[i])
            except InvalidOperation:
                ok = False
                break
        if not ok:
            continue
        places = max_places_by_idx.get(i, 0)
        out_fields[i] = _format_decimal_sped(total, places)

    log.append(f"[Bloco M] {reg} consolidado: {len(occurrences)} ocorrência(s) -> 1 (somado).")
    return build_line(reg, out_fields)


@dataclass
class BlockMGroup:
    start_reg: str
    lines: List[str]


def _m_reg_num(reg: str) -> int:
    if reg and reg.startswith("M") and len(reg) == 4 and reg[1:].isdigit():
        return int(reg[1:])
    return 9999


def _m_is_group_start(reg: str) -> bool:
    """
    Heurística segura para manter pai/filhos juntos e evitar quebrar hierarquia:
    - Considera início de grupo:
      * registros que terminam com "00" (M200, M300, M400, M600, ...)
      * e alguns pais conhecidos que não terminam com "00" (ex.: M100, M500, M610, M800)
    """
    if not reg or not reg.startswith("M") or len(reg) != 4:
        return False
    if reg in ("M100", "M500", "M610", "M800"):
        return True
    if reg.endswith("00"):
        return True
    return False


def parse_block_m_groups(body_lines: List[str]) -> List[BlockMGroup]:
    """
    Quebra o corpo do bloco M (sem M001/M990) em grupos, mantendo filhos imediatamente após o pai.
    """
    groups: List[BlockMGroup] = []
    cur_lines: List[str] = []
    cur_start: str = ""

    for ln in body_lines:
        r = get_reg(ln)
        if not r:
            continue
        if not cur_lines:
            cur_lines = [ln]
            cur_start = r
            continue

        if _m_is_group_start(r):
            groups.append(BlockMGroup(start_reg=cur_start, lines=cur_lines))
            cur_lines = [ln]
            cur_start = r
        else:
            cur_lines.append(ln)

    if cur_lines:
        groups.append(BlockMGroup(start_reg=cur_start, lines=cur_lines))

    return groups


def _insert_group_sorted(groups: List[BlockMGroup], g: BlockMGroup) -> None:
    """
    Insere g mantendo uma ordem numérica aproximada por REG (M100 < M200 < ... < M610).
    Insere após o último grupo com REG <= o novo REG (estável).
    """
    num = _m_reg_num(g.start_reg)
    idx = 0
    while idx < len(groups) and _m_reg_num(groups[idx].start_reg) <= num:
        idx += 1
    groups[idx:idx] = [g]


def _extract_cod_cont_from_parent_line(parent_line: str) -> str:
    _, f = split_fields(parent_line)
    return (f[0] if f else "").strip()


def _consolidate_groups_by_cod_cont(
    reg: str, groups: List[BlockMGroup], log: List[str]
) -> List[BlockMGroup]:
    """
    Consolida grupos do tipo reg por COD_CONT (campo 1 do registro pai),
    somando campos numéricos no pai (quando houver mais de 1 ocorrência) e
    unindo filhos por linha exata (sem duplicar).
    """
    buckets: "OrderedDict[str, List[BlockMGroup]]" = OrderedDict()
    for g in groups:
        if g.start_reg != reg or not g.lines:
            continue
        key = _extract_cod_cont_from_parent_line(g.lines[0])
        buckets.setdefault(key, []).append(g)

    out: List[BlockMGroup] = []
    for cod_cont, gs in buckets.items():
        parent_lines = [x.lines[0] for x in gs if x.lines]
        parent_merged = sum_unique_record(reg, parent_lines, log) if len(parent_lines) > 1 else parent_lines[0]

        seen = set()
        merged_children: List[str] = []
        for x in gs:
            for ch in x.lines[1:]:
                if ch not in seen:
                    merged_children.append(ch)
                    seen.add(ch)

        if len(gs) > 1:
            log.append(f"[Bloco M] {reg} grupos consolidados por COD_CONT='{cod_cont}': {len(gs)} -> 1.")
        out.append(BlockMGroup(start_reg=reg, lines=[parent_merged] + merged_children))

    return out


def _ensure_m400_before_m500(groups: List[BlockMGroup], log: List[str]) -> List[BlockMGroup]:
    """
    Garante que TODOS os grupos M400 fiquem antes do primeiro M500.
    Não altera ordem relativa entre M400 movidos.
    """
    first_m500 = next((i for i, g in enumerate(groups) if g.start_reg == "M500"), -1)
    if first_m500 <= 0:
        return groups

    m400_after: List[BlockMGroup] = []
    kept: List[BlockMGroup] = []
    for i, g in enumerate(groups):
        if i > first_m500 and g.start_reg == "M400":
            m400_after.append(g)
        else:
            kept.append(g)

    if not m400_after:
        return groups

    insert_at = next((i for i, g in enumerate(kept) if g.start_reg == "M500"), len(kept))
    kept[insert_at:insert_at] = m400_after
    log.append(f"[Bloco M] Reordenado: {len(m400_after)} grupo(s) M400 movido(s) para antes de M500.")
    return kept


def merge_block_m(matrix_lines: List[str], filial_blocks_m: List[List[str]], log: List[str]) -> List[str]:
    """
    Merge do bloco M corrigido para evitar:
    - inserir M100/M105/M500/M505 no final quebrando hierarquia (ex.: esperado M611).
    - quebrar pai/filhos ao inserir linhas “soltas”.

    Estratégia:
    - extrai corpo (sem M001/M990) da matriz e filiais
    - consolida M200 e M600 (únicos)
    - quebra em grupos pai+filhos por heurística e mescla por grupos (não por linha solta)
    - consolida M100 e M500 por COD_CONT (quando houver múltiplos)
    - insere grupos novos em ordem numérica aproximada do REG
    - garante M400 antes de M500
    - recalcula M990
    """
    close_reg = CLOSE_REG_BY_BLOCK["M"]
    open_reg = OPEN_REG_BY_BLOCK["M"]

    # base = matriz (ou primeira filial se não houver matriz)
    base = matrix_lines[:] if matrix_lines else []
    if not base:
        for fb in filial_blocks_m:
            if fb:
                base = fb[:]
                break

    base_wo_close = strip_block_close(base, close_reg)

    open_line: Optional[str] = None
    base_body: List[str] = []
    for ln in base_wo_close:
        r = get_reg(ln)
        if r == open_reg and open_line is None:
            open_line = ln
        elif r == open_reg:
            continue
        elif r == close_reg:
            continue
        else:
            base_body.append(ln)

    if open_line is None:
        open_line = build_line(open_reg, ["0"])

    # Coleta corpo de todas as filiais (sem M001/M990)
    filial_body_all: List[str] = []
    for fb in filial_blocks_m:
        if not fb:
            continue
        fb_wo_close = strip_block_close(fb, close_reg)
        for ln in fb_wo_close:
            r = get_reg(ln)
            if r in (open_reg, close_reg):
                continue
            filial_body_all.append(ln)

    # 1) Consolida M200/M600 (únicos)
    m200_occ = [ln for ln in base_body if get_reg(ln) == "M200"] + [ln for ln in filial_body_all if get_reg(ln) == "M200"]
    m600_occ = [ln for ln in base_body if get_reg(ln) == "M600"] + [ln for ln in filial_body_all if get_reg(ln) == "M600"]

    summed_m200 = sum_unique_record("M200", m200_occ, log) if m200_occ else None
    summed_m600 = sum_unique_record("M600", m600_occ, log) if m600_occ else None

    base_body = [ln for ln in base_body if get_reg(ln) not in ("M200", "M600")]
    filial_body_all = [ln for ln in filial_body_all if get_reg(ln) not in ("M200", "M600")]

    # 2) Grupos (mantém pai/filhos juntos)
    base_groups = parse_block_m_groups(base_body)
    filial_groups = parse_block_m_groups(filial_body_all)

    # 3) Consolida M100 e M500 por COD_CONT (se existirem em mais de um arquivo)
    base_m100 = [g for g in base_groups if g.start_reg == "M100"]
    base_m500 = [g for g in base_groups if g.start_reg == "M500"]
    fil_m100 = [g for g in filial_groups if g.start_reg == "M100"]
    fil_m500 = [g for g in filial_groups if g.start_reg == "M500"]

    other_base = [g for g in base_groups if g.start_reg not in ("M100", "M500")]
    other_fil = [g for g in filial_groups if g.start_reg not in ("M100", "M500")]

    consolidated_m100 = _consolidate_groups_by_cod_cont("M100", base_m100 + fil_m100, log) if (base_m100 or fil_m100) else []
    consolidated_m500 = _consolidate_groups_by_cod_cont("M500", base_m500 + fil_m500, log) if (base_m500 or fil_m500) else []

    # 4) Mescla grupos "outros" (dedupe por grupo exato)
    merged_groups: List[BlockMGroup] = other_base[:]
    existing_sigs = {tuple(g.lines) for g in merged_groups}

    added_groups = 0
    for g in other_fil:
        sig = tuple(g.lines)
        if sig in existing_sigs:
            continue
        existing_sigs.add(sig)
        _insert_group_sorted(merged_groups, g)
        added_groups += 1

    if added_groups:
        log.append(f"[Bloco M] Adicionados {added_groups} grupo(s) (fora M100/M500/M200/M600).")

    # 5) Insere grupos consolidados M100/M500 e registros únicos M200/M600 nas posições corretas
    for g in consolidated_m100:
        _insert_group_sorted(merged_groups, g)

    if summed_m200:
        _insert_group_sorted(merged_groups, BlockMGroup(start_reg="M200", lines=[summed_m200]))

    for g in consolidated_m500:
        _insert_group_sorted(merged_groups, g)

    if summed_m600:
        _insert_group_sorted(merged_groups, BlockMGroup(start_reg="M600", lines=[summed_m600]))

    # 6) Garante M400 antes de M500
    merged_groups = _ensure_m400_before_m500(merged_groups, log)

    # 7) Reconstrói corpo final
    body: List[str] = []
    for g in merged_groups:
        body.extend(g.lines)

    merged_wo_close = [open_line] + body
    has_content = len(merged_wo_close) > 1
    merged_wo_close[0] = replace_first_field(merged_wo_close[0], "0" if has_content else "1")

    return count_block_lines_with_close(merged_wo_close, close_reg)


# -----------------------------
# Bloco 9 (regerar)
# -----------------------------

def build_block9(all_lines_except_block9: List[str]) -> List[str]:
    """
    Gera:
    - 9001|0|
    - 9900 (um por registro existente + 9001/9900/9990/9999)
    - 9990 (QTD_LIN_9) inclui 9999
    - 9999 (QTD_LIN total do arquivo incluindo 9999)
    """
    reg_counts: Dict[str, int] = {}
    for ln in all_lines_except_block9:
        r = get_reg(ln)
        if not r:
            continue
        reg_counts[r] = reg_counts.get(r, 0) + 1

    mandatory = {"9001", "9900", "9990", "9999"}
    regs = sorted(set(reg_counts.keys()) | mandatory)
    total_9900 = len(regs)

    blk9_lines: List[str] = [build_line("9001", ["0"])]

    for r in regs:
        if r == "9001":
            cnt = 1
        elif r == "9900":
            cnt = total_9900
        elif r == "9990":
            cnt = 1
        elif r == "9999":
            cnt = 1
        else:
            cnt = reg_counts.get(r, 0)
        blk9_lines.append(build_line("9900", [r, str(cnt)]))

    qtd_lin_9 = len(blk9_lines) + 2  # (+9990 e +9999)
    blk9_lines.append(build_line("9990", [str(qtd_lin_9)]))

    total_lines = len(all_lines_except_block9) + len(blk9_lines) + 1
    blk9_lines.append(build_line("9999", [str(total_lines)]))

    return blk9_lines


# -----------------------------
# Pipeline principal
# -----------------------------

def combine_speds(matriz_path: str, filiais_paths: List[str], out_dir: str, overwrite: bool = False) -> Tuple[str, List[str]]:
    log: List[str] = []

    if not os.path.isfile(matriz_path):
        raise FileNotFoundError(f"Matriz não encontrada: {matriz_path}")
    if not filiais_paths:
        raise ValueError("Selecione ao menos 1 arquivo de filial.")

    enc_m = detect_encoding(matriz_path)
    matriz_lines_raw = read_lines(matriz_path, enc_m)

    matriz_lines = sanitize_lines(matriz_lines_raw, "MATRIZ", log)
    precheck_file(matriz_lines, "MATRIZ", log)

    filiais_lines: List[List[str]] = []
    for i, p in enumerate(filiais_paths, start=1):
        if not os.path.isfile(p):
            raise FileNotFoundError(f"Filial não encontrada: {p}")
        enc_f = detect_encoding(p)
        fl_raw = read_lines(p, enc_f)

        label = f"FILIAL#{i}"
        fl = sanitize_lines(fl_raw, label, log)
        precheck_file(fl, label, log)

        fl = dedupe_filial_against_matriz(matriz_lines, fl, label, log)
        filiais_lines.append(fl)

    matriz_blocks = parse_blocks(matriz_lines)
    filiais_blocks_list = [parse_blocks(fl) for fl in filiais_lines]

    merged_blocks: Dict[str, List[str]] = {}

    # Bloco 0
    m0 = matriz_blocks.get("0", [])
    f0s = [fb.get("0", []) for fb in filiais_blocks_list if fb.get("0")]
    merged_blocks["0"] = merge_block0(m0, f0s, log)

    # Blocos simples
    for b in ["A", "C", "D", "F", "I", "1"]:
        mb = matriz_blocks.get(b, [])
        fbs = [fb.get(b, []) for fb in filiais_blocks_list if fb.get(b)]
        if not mb and not fbs:
            continue
        merged_blocks[b] = merge_simple_block(b, mb, fbs, log)

    # Bloco M (especial, corrigido)
    mbm = matriz_blocks.get("M", [])
    fbm = [fb.get("M", []) for fb in filiais_blocks_list if fb.get("M")]
    if mbm or fbm:
        merged_blocks["M"] = merge_block_m(mbm, fbm, log)

    # Monta linhas (exceto bloco 9)
    out_lines_no9: List[str] = []
    for b in CANONICAL_BLOCK_ORDER:
        if b == "9":
            continue
        if b in merged_blocks:
            out_lines_no9.extend(merged_blocks[b])

    block9 = build_block9(out_lines_no9)

    out_lines = out_lines_no9 + block9
    out_text = "\r\n".join(out_lines) + "\r\n"

    os.makedirs(out_dir, exist_ok=True)
    out_name = os.path.basename(matriz_path)
    out_path = os.path.join(out_dir, out_name)

    if os.path.exists(out_path) and not overwrite:
        raise FileExistsError(f"Arquivo já existe e overwrite=False: {out_path}")

    try:
        with open(out_path, "w", encoding=enc_m, errors="strict", newline="") as f:
            f.write(out_text)
    except UnicodeEncodeError:
        enc_out = "utf-8-sig"
        with open(out_path, "w", encoding=enc_out, errors="strict", newline="") as f:
            f.write(out_text)
        log.append(f"AVISO: saída não coube em {enc_m}; gravado em {enc_out}.")

    log.append(f"OK: gerado {out_path}")
    return out_path, log


# -----------------------------
# GUI (Tkinter)
# -----------------------------

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Combinador SPED - Matriz + Filiais")
    root.geometry("900x600")

    matriz_var = tk.StringVar(value="")
    filiais_var = tk.StringVar(value="(nenhum)")
    out_var = tk.StringVar(value=f"Pasta de saída: {DEFAULT_OUT_DIR}")

    selected_filiais: List[str] = []

    def pick_matriz():
        p = filedialog.askopenfilename(
            title="Selecione o SPED da Matriz",
            filetypes=[("Arquivos texto", "*.txt"), ("Todos", "*.*")],
        )
        if p:
            matriz_var.set(p)
            out_preview = os.path.join(DEFAULT_OUT_DIR, os.path.basename(p))
            out_var.set(f"Saída: {out_preview}")

    def pick_filiais():
        nonlocal selected_filiais
        ps = filedialog.askopenfilenames(
            title="Selecione 1+ SPED(s) das Filiais",
            filetypes=[("Arquivos texto", "*.txt"), ("Todos", "*.*")],
        )
        if ps:
            selected_filiais = list(ps)
            filiais_var.set(f"{len(selected_filiais)} arquivo(s) selecionado(s)")

    def do_combine():
        matriz = matriz_var.get().strip()
        if not matriz:
            messagebox.showerror("Erro", "Selecione o arquivo da Matriz.")
            return
        if not selected_filiais:
            messagebox.showerror("Erro", "Selecione ao menos 1 arquivo de Filial.")
            return

        out_path = os.path.join(DEFAULT_OUT_DIR, os.path.basename(matriz))

        overwrite = False
        if os.path.exists(out_path):
            overwrite = messagebox.askyesno(
                "Arquivo já existe",
                f"Já existe:\n{out_path}\n\nDeseja sobrescrever?",
            )
            if not overwrite:
                return

        try:
            text_log.delete("1.0", tk.END)
            path, logs = combine_speds(matriz, selected_filiais, DEFAULT_OUT_DIR, overwrite=overwrite)
            for ln in logs:
                text_log.insert(tk.END, ln + "\n")
            messagebox.showinfo("Concluído", f"Arquivo gerado:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    frm = tk.Frame(root, padx=12, pady=12)
    frm.pack(fill="both", expand=True)

    row = 0
    tk.Label(frm, text="Matriz:").grid(row=row, column=0, sticky="w")
    tk.Entry(frm, textvariable=matriz_var, width=95).grid(row=row, column=1, sticky="we", padx=8)
    tk.Button(frm, text="Selecionar...", command=pick_matriz).grid(row=row, column=2, sticky="e")
    row += 1

    tk.Label(frm, text="Filiais:").grid(row=row, column=0, sticky="w")
    tk.Label(frm, textvariable=filiais_var, anchor="w").grid(row=row, column=1, sticky="we", padx=8)
    tk.Button(frm, text="Selecionar 1+...", command=pick_filiais).grid(row=row, column=2, sticky="e")
    row += 1

    tk.Label(frm, textvariable=out_var, fg="gray30").grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 8))
    row += 1

    tk.Button(frm, text="COMBINAR", height=2, command=do_combine).grid(
        row=row, column=0, columnspan=3, sticky="we", pady=(0, 10)
    )
    row += 1

    tk.Label(frm, text="Log:").grid(row=row, column=0, sticky="w")
    row += 1
    text_log = scrolledtext.ScrolledText(frm, wrap=tk.WORD, height=22)
    text_log.grid(row=row, column=0, columnspan=3, sticky="nsew")

    frm.grid_columnconfigure(1, weight=1)
    frm.grid_rowconfigure(row, weight=1)

    root.mainloop()


# -----------------------------
# CLI + fallback para GUI
# -----------------------------

def main():
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("matriz", nargs="?", help="Arquivo SPED da Matriz (.txt)")
    parser.add_argument("filiais", nargs="*", help="Arquivos SPED das Filiais (.txt) (1+)")
    parser.add_argument("--out", default=DEFAULT_OUT_DIR, help="Pasta de saída (padrão: pasta Combinados)")
    parser.add_argument("--overwrite", action="store_true", help="Sobrescrever arquivo de saída se existir")
    args = parser.parse_args()

    if not args.matriz and not args.filiais:
        run_gui()
        return

    if not args.matriz or not args.filiais:
        parser.error("Informe matriz e 1+ filiais, ou execute sem argumentos para abrir a GUI.")

    out_path, logs = combine_speds(args.matriz, args.filiais, args.out, overwrite=args.overwrite)
    print(out_path)
    for ln in logs:
        print(ln)


if __name__ == "__main__":
    main()
