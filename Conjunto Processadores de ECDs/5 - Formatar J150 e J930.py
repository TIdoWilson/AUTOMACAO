# -*- coding: utf-8 -*-
"""
Formato de saída:
- remove linhas |J150| com valor 0 e, se TODAS as linhas do bloco forem 0,
  substitui o bloco inteiro pelo texto padrão.
- substitui todo o bloco |J930| por um texto fixo.
"""

import os
from datetime import datetime

# ===================== CONFIGURACAO =====================
# Caminho do TXT de entrada
INPUT_PATH = ""

# Pasta e arquivo de log de erros (simples e direto)
ERRO_DIR = r"W:\SPEDs\ECD"
ERRO_ARQ = "ERROS DE FORMATAÇÂO ECD.txt"

# Log de alteracoes (detalhado)
LOG_ARQ = "LOG_FORMATACAO_ECD.txt"

# Texto de substituicao quando TODAS as linhas do bloco |J150| forem 0
REPLACEMENT_BLOCK_J150 = [
    "|J150|0|DRE_002_RB_311                |T|2|DRE_319_TITULO_LUCRO_PREJUIZO |VENDA DE PRODUTOS|0|C|0|C|R||",
    "|J150|1|DRE_003_RB_31101              |T|3|DRE_002_RB_311                |VENDA DE PROD DE FABRICACAO PROPRIA|0|C|0|C|R||",
    "|J150|2|DRE_004_RB_31101001000        |D|4|DRE_003_RB_31101              |VENDAS A VISTA|0|C|0|C|R||",
    "|J150|3|DRE_319_TITULO_LUCRO_PREJUIZO |T|1|                              |LUCRO LÍQUIDO DO EXERCÍCIO|0|C|0|C|R||",
]

# Texto de substituicao do bloco |J930| (sempre substitui o bloco inteiro)
REPLACEMENT_BLOCK_J930 = [
    "|J930|WILSON M. LOPES CONTABILIDADE LTDA|07053914000160|Pessoa Jurídica (e-CNPJ ou e-PJ)|001||WILSON@WILSONLOPES.COM.BR|4635203300||||S|",
    "|J930|WILSON MARCOS LOPES|60298227991|Contador|900|PR034211O0|WILSON@WILSONLOPES.COM.BR|4635203300|PR|PR/2024/366551|21072025|N|",
    "|J930|WILSON MARCOS LOPES |60298227991|Procurador|309|PR034211O0|WILSON@WILSONLOPES.COM.BR|4635203300|PR|||S|",
]
# ========================================================


def _extract_valor_j150(line: str) -> str | None:
    if not line.startswith("|J150|"):
        return None
    parts = line.rstrip("\r\n").split("|")
    if len(parts) < 6:
        return None
    return parts[-5]


def _get_block_code(line: str) -> str | None:
    if line.startswith("|J150|"):
        return "J150"
    if line.startswith("|J930|"):
        return "J930"
    return None


def _detect_encoding(path: str) -> str:
    with open(path, "rb") as f_in:
        sample = f_in.read(8192)
    try:
        sample.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        try:
            sample.decode("cp1252")
            return "cp1252"
        except UnicodeDecodeError:
            return "latin-1"


def _open_text_file(path: str):
    encoding = _detect_encoding(path)
    if encoding == "utf-8":
        return open(path, "r", encoding=encoding, newline="")
    return open(path, "r", encoding=encoding, errors="replace", newline="")


def process_file(input_path: str) -> str:
    base_dir = os.path.dirname(input_path)
    base_name = os.path.basename(input_path)
    output_path = os.path.join(base_dir, base_name)
    original_backup = os.path.join(base_dir, f"ORIGINAL_{base_name}")
    temp_output = os.path.join(base_dir, f".TMP_{base_name}")
    log_path = os.path.join(base_dir, LOG_ARQ)

    if os.path.exists(temp_output):
        raise FileExistsError(f"Ja existe um TMP_ para este arquivo: {temp_output}")

    line_ending = "\n"

    # Permite reprocessamento seguro:
    # - se ORIGINAL_ existir e for mais antigo que o arquivo base atual, substitui ORIGINAL_ pelo arquivo base;
    # - se ORIGINAL_ existir e for mais novo/igual, reaproveita ORIGINAL_;
    # - se ORIGINAL_ nao existir, cria normalmente a partir do arquivo base.
    if os.path.exists(original_backup):
        mtime_original = os.path.getmtime(original_backup)
        mtime_base = os.path.getmtime(input_path)
        if mtime_base > mtime_original:
            os.remove(original_backup)
            os.replace(input_path, original_backup)
        source_path = original_backup
    else:
        os.replace(input_path, original_backup)
        source_path = original_backup

    try:
        with _open_text_file(source_path) as f_in:
            output_lines, counts = _process_lines(f_in, line_ending)

        _atualizar_contadores(output_lines, counts)
        _registrar_log(log_path, original_backup, output_path, counts)

        with open(temp_output, "w", encoding="utf-8", newline="") as f_out:
            for line in output_lines:
                f_out.write(line)

        os.replace(temp_output, output_path)
        return output_path
    finally:
        if os.path.exists(temp_output):
            try:
                os.remove(temp_output)
            except OSError:
                pass


def _get_line_ending(line: str, default: str) -> str:
    if line.endswith("\r\n"):
        return "\r\n"
    if line.endswith("\n"):
        return "\n"
    return default


def _process_lines(f_in, line_ending: str):
    output_lines: list[str] = []
    in_block = False
    current_code: str | None = None
    block_lines: list[str] = []

    counts = {
        "old_j150": 0,
        "new_j150": 0,
        "old_j930": 0,
        "new_j930": 0,
        "j150_block_replaced": False,
        "j150_removed_lines": [],
        "i030_updated": 0,
        "j900_updated": 0,
        "line_ending": line_ending,
    }

    for line in f_in:
        if line.endswith("\r\n"):
            counts["line_ending"] = "\r\n"

        line_code = _get_block_code(line)

        if in_block and line_code != current_code:
            _write_block(current_code, block_lines, output_lines, counts)
            block_lines = []
            in_block = False
            current_code = None

        if line_code is not None:
            in_block = True
            current_code = line_code
            block_lines.append(line)
            continue

        output_lines.append(line)

    if in_block:
        _write_block(current_code, block_lines, output_lines, counts)

    return output_lines, counts


def _write_block(
    code: str | None, block_lines: list[str], output_lines: list[str], counts: dict
) -> None:
    if not code:
        for bl in block_lines:
            output_lines.append(bl)
        return

    if code == "J930":
        counts["old_j930"] += len(block_lines)
        counts["new_j930"] += len(REPLACEMENT_BLOCK_J930)
        line_ending = counts["line_ending"]
        for repl in REPLACEMENT_BLOCK_J930:
            output_lines.append(repl + line_ending)
        return

    if code != "J150":
        for bl in block_lines:
            output_lines.append(bl)
        return

    counts["old_j150"] += len(block_lines)

    valores = []
    for bl in block_lines:
        valor = _extract_valor_j150(bl)
        if valor is not None:
            valores.append(valor)

    if valores and all(v == "0" for v in valores):
        counts["new_j150"] += len(REPLACEMENT_BLOCK_J150)
        counts["j150_block_replaced"] = True
        counts["j150_removed_lines"] = [bl.rstrip("\r\n") for bl in block_lines]
        line_ending = counts["line_ending"]
        for repl in REPLACEMENT_BLOCK_J150:
            output_lines.append(repl + line_ending)
        return

    for bl in block_lines:
        output_lines.append(bl)
        counts["new_j150"] += 1


def _atualizar_contadores(output_lines: list[str], counts: dict) -> None:
    old_j150 = counts["old_j150"]
    new_j150 = counts["new_j150"]
    old_j930 = counts["old_j930"]
    new_j930 = counts["new_j930"]

    delta_j150 = new_j150 - old_j150
    delta_j930 = new_j930 - old_j930

    for idx, line in enumerate(output_lines):
        if line.startswith("|9900|J150|"):
            output_lines[idx] = _set_9900_count(line, new_j150, counts["line_ending"])
        elif line.startswith("|9900|J930|"):
            output_lines[idx] = _set_9900_count(line, new_j930, counts["line_ending"])

    for idx, line in enumerate(output_lines):
        if line.startswith("|I030|TERMO DE ABERTURA|"):
            updated_line, updated = _incrementar_numero_termo(
                line, "I030", "TERMO DE ABERTURA", counts["line_ending"]
            )
            output_lines[idx] = updated_line
            if updated:
                counts["i030_updated"] += 1
        elif line.startswith("|J900|TERMO DE ENCERRAMENTO|"):
            updated_line, updated = _incrementar_numero_termo(
                line, "J900", "TERMO DE ENCERRAMENTO", counts["line_ending"]
            )
            output_lines[idx] = updated_line
            if updated:
                counts["j900_updated"] += 1

    if delta_j150 == 0 and delta_j930 == 0:
        return

    for idx, line in enumerate(output_lines):
        if line.startswith("|J990|"):
            output_lines[idx] = _set_j990_count(
                line, delta_j150 + delta_j930, counts["line_ending"]
            )
            break

    for idx, line in enumerate(output_lines):
        if line.startswith("|9999|"):
            output_lines[idx] = _set_9999_count(
                line, delta_j150 + delta_j930, counts["line_ending"]
            )
            break


def _set_9900_count(line: str, new_count: int, default_ending: str) -> str:
    ending = _get_line_ending(line, default_ending)
    parts = line.rstrip("\r\n").split("|")
    if len(parts) < 4:
        return line
    parts[3] = str(new_count)
    return "|".join(parts) + ending


def _set_j990_count(line: str, delta: int, default_ending: str) -> str:
    ending = _get_line_ending(line, default_ending)
    parts = line.rstrip("\r\n").split("|")
    if len(parts) < 3:
        return line
    try:
        original = int(parts[2])
    except ValueError:
        return line
    parts[2] = str(original + delta)
    return "|".join(parts) + ending


def _set_9999_count(line: str, delta: int, default_ending: str) -> str:
    ending = _get_line_ending(line, default_ending)
    parts = line.rstrip("\r\n").split("|")
    if len(parts) < 3:
        return line
    try:
        original = int(parts[2])
    except ValueError:
        return line
    parts[2] = str(original + delta)
    return "|".join(parts) + ending


def _incrementar_numero_termo(
    line: str, registro: str, descricao: str, default_ending: str
) -> tuple[str, bool]:
    ending = _get_line_ending(line, default_ending)
    parts = line.rstrip("\r\n").split("|")
    if len(parts) < 5:
        return line, False
    if parts[1] != registro or parts[2] != descricao:
        return line, False

    try:
        numero = int(parts[3])
    except ValueError:
        return line, False

    parts[3] = str(numero + 1)
    return "|".join(parts) + ending, True


def _registrar_log(log_path: str, original_path: str, output_path: str, counts: dict) -> None:
    data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    old_j150 = counts["old_j150"]
    new_j150 = counts["new_j150"]
    old_j930 = counts["old_j930"]
    new_j930 = counts["new_j930"]
    delta_total = (new_j150 - old_j150) + (new_j930 - old_j930)

    with open(log_path, "a", encoding="utf-8", newline="") as f_log:
        f_log.write(f"[{data}] ARQUIVO ORIGINAL: {original_path}\n")
        f_log.write(f"[{data}] ARQUIVO SAIDA:    {output_path}\n")
        f_log.write(
            f"[{data}] J930: {old_j930} -> {new_j930} (delta {new_j930 - old_j930})\n"
        )
        f_log.write(
            f"[{data}] J150: {old_j150} -> {new_j150} (delta {new_j150 - old_j150})\n"
        )
        f_log.write(
            f"[{data}] J150 bloco substituido: {'SIM' if counts['j150_block_replaced'] else 'NAO'}\n"
        )
        if counts["j150_removed_lines"]:
            f_log.write(f"[{data}] J150 removidos (linhas originais):\n")
            for line in counts["j150_removed_lines"]:
                f_log.write(f"[{data}] {line}\n")
        f_log.write(f"[{data}] I030 termo abertura incrementado: {counts['i030_updated']} linha(s)\n")
        f_log.write(
            f"[{data}] J900 termo encerramento incrementado: {counts['j900_updated']} linha(s)\n"
        )
        f_log.write(f"[{data}] DELTA TOTAL LINHAS: {delta_total}\n")
        f_log.write(f"[{data}] FIM\n")


def main() -> None:
    import sys

    input_path = None
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    elif INPUT_PATH:
        input_path = INPUT_PATH

    if not input_path:
        raise ValueError("Informe o caminho do arquivo (argumento) ou configure INPUT_PATH.")
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Arquivo nao encontrado: {input_path}")

    output_path = process_file(input_path)
    print(f"OK: {output_path}")


def _registrar_erro(msg: str) -> None:
    os.makedirs(ERRO_DIR, exist_ok=True)
    path = os.path.join(ERRO_DIR, ERRO_ARQ)
    data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(path, "a", encoding="utf-8", newline="") as f_log:
        f_log.write(f"[{data}] {msg}\n")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        _registrar_erro(f"Nao foi possivel formatar o arquivo. Motivo: {exc}")
        raise
