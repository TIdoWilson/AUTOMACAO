#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
import datetime
import shutil
import threading
import subprocess
from pathlib import Path
from queue import Queue, Empty
import logging
from logging.handlers import RotatingFileHandler

# ========= CONFIG =========
# Pasta raiz a monitorar
ROOT_DIR = Path(r"C:\Robos\PASTA_MONITORADA")

# Comando para processar o arquivo detectado (recebe {file})
# Use o MESMO python do serviço para evitar problemas de PATH.
PY = r'C:\ROBOS\venv\Scripts\python.exe'

# === mapeamento de scripts por subpasta (nome da pasta do script) ===
SCRIPT_MAP = {
    "CONCILIADOR RAZAO IOB": fr'"{PY}" "C:\Robos\CONCILIADOR_IOB.PY" "{{file}}"',
    "SEPARADOR DE FERIAS POR FUNCIONARIOS": fr'"{PY}" "C:\Robos\SEPARAR_FERIAS_POR_FUNCIONARIO.PY" "{{file}}"',
    "SEPARADOR DE RELATORIO DE FERIAS": fr'"{PY}" "C:\Robos\SEPARAR_RELATORIO_DE_FERIAS.py" "{{file}}"',
    "SEPARADOR DE HOLERITES AGRUPADOS": fr'"{PY}" "C:\Robos\SEPARAR_HOLERITES.py" "{{file}}"',
    "LOTES INTERNETS": fr'"{PY}" "C:\Robos\LOTE_INTERNETS.py" "{{file}}"',
    "LOTES TOSCAN": fr'"{PY}" "C:\Robos\LOTE_TOSCAN.py" "{{file}}"',
}

# Intervalo de varredura (s)
SCAN_INTERVAL = 5

# "ignore_existing" => ignora o que já existe ao iniciar
# "process_existing" => processa tudo que já está lá na primeira rodada
STARTUP_MODE = "process_existing"

# Extensões permitidas
ALLOW_EXTENSIONS = {".pdf", ".xlsx", ".csv", ".txt"}

# Nome da subpasta intermediária
PROCESSING_DIRNAME = "em processamento"

# Pastas que não devem ser varridas como "pastas de script"
EXCLUDED_DIRS = {
    PROCESSING_DIRNAME,  # "em processamento"
    "originais",
    "resultado",
    "resultados",
}

# Logs do app
LOG_DIR = Path(r"C:\Robos")
LOG_FILE = LOG_DIR / "monitor.log"
# ========= FIM CONFIG =========


def setup_logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("monitor")
    logger.setLevel(logging.INFO)
    fh = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


logger: logging.Logger


def list_files_recursive(root: Path):
    """
    Lista (pasta_script, arquivo) para arquivos que estejam
    DIRETAMENTE no 3º nível:

        root/<SETOR>/<PASTA_SCRIPT>/<arquivo>

    Ignora:
    - subpastas em EXCLUDED_DIRS (ex.: "em processamento", "originais", "resultado", "resultados")
    - qualquer coisa abaixo desse nível.
    """
    # 1º nível: setores (RH, CONTABIL, TESTES, etc.)
    for setor in root.iterdir():
        if not setor.is_dir():
            continue

        # 2º nível: pastas dos scripts
        for pasta_script in setor.iterdir():
            if not pasta_script.is_dir():
                continue
            if pasta_script.name in EXCLUDED_DIRS:
                continue

            # 3º nível: arquivos diretamente dentro da pasta do script
            for f in pasta_script.iterdir():
                if f.is_file():
                    if ALLOW_EXTENSIONS and f.suffix.lower() not in ALLOW_EXTENSIONS:
                        continue
                    yield pasta_script, f


def get_rel_subfolder(root: Path, folder: Path) -> str:
    """
    Retorna o caminho relativo root -> pasta_script, no formato:
        RH/SEPARADOR DE HOLERITES AGRUPADOS
    """
    try:
        rel = folder.relative_to(root)
    except ValueError:
        rel = folder
    return str(rel).replace("\\", "/")


def resolve_command_for_folder(rel_subfolder: str) -> str | None:
    """
    Localiza o comando do SCRIPT_MAP para a subpasta.
    Agora usa SEMPRE o último componente do caminho, que é o nome da pasta do script.

    Exemplo:
        rel_subfolder = 'RH/SEPARADOR DE HOLERITES AGRUPADOS'
        script_name   = 'SEPARADOR DE HOLERITES AGRUPADOS'
    """
    if rel_subfolder in SCRIPT_MAP:
        return SCRIPT_MAP[rel_subfolder]

    script_name = rel_subfolder.split("/")[-1] if rel_subfolder else rel_subfolder
    if script_name in SCRIPT_MAP:
        return SCRIPT_MAP[script_name]

    logger.warning(
        f"Nenhum monitor configurado para a pasta: '{rel_subfolder}' "
        f"(script '{script_name}'). Nenhuma ação será executada."
    )
    return None


def move_to_processing(folder: Path, file_path: Path) -> Path | None:
    """
    Move o arquivo para 'em processamento' dentro da mesma pasta_script.
    Exemplo:
        ...\RH\SEPARADOR DE HOLERITES AGRUPADOS\arquivo.pdf
        -> ...\RH\SEPARADOR DE HOLERITES AGRUPADOS\em processamento\arquivo.pdf
    """
    processing_dir = folder / PROCESSING_DIRNAME
    processing_dir.mkdir(parents=True, exist_ok=True)
    target = processing_dir / file_path.name
    try:
        size1 = file_path.stat().st_size
        time.sleep(0.5)
        size2 = file_path.stat().st_size
        if size1 != size2:
            logger.info(f"Arquivo ainda crescendo; tentar novamente: {file_path}")
            return None
        shutil.move(str(file_path), str(target))
        return target
    except Exception as e:
        logger.error(f"Falha ao mover {file_path} -> {target}: {e}")
        return None


class Job:
    def __init__(self, rel_folder: str, file_in_processing: Path, cmd_template: str):
        self.rel_folder = rel_folder          # ex.: 'RH/SEPARADOR DE HOLERITES AGRUPADOS'
        self.file_in_processing = file_in_processing  # caminho completo em 'em processamento'
        self.cmd_template = cmd_template

    def command(self) -> str:
        return self.cmd_template.format(folder=self.rel_folder, file=str(self.file_in_processing))


class Worker(threading.Thread):
    def __init__(self, queue: Queue):
        super().__init__(daemon=True)
        self.queue = queue

    def run(self):
        while True:
            try:
                job: Job = self.queue.get(timeout=1)
            except Empty:
                continue

            cmd = job.command()
            src_file = Path(job.file_in_processing)  # ...\em processamento\<arquivo>.pdf
            processing_dir = src_file.parent         # ...\em processamento
            pasta_script = processing_dir.parent     # pasta do script (ex.: ...\RH\SEPARADOR DE HOLERITES AGRUPADOS)
            originais = pasta_script / "originais"
            resultados = pasta_script / "resultados"
            originais.mkdir(exist_ok=True)
            resultados.mkdir(exist_ok=True)

            try:
                logger.info(f"Executando: {cmd}")
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True)

                if result.returncode == 0:
                    logger.info(f"OK (ret={result.returncode}).")
                    if result.stdout.strip():
                        logger.info(f"STDOUT:\n{result.stdout.strip()}")
                    if result.stderr.strip():
                        logger.info(f"STDERR (não fatal):\n{result.stderr.strip()}")

                    # 1) mover o PDF ORIGINAL (de 'em processamento') -> 'originais'
                    try:
                        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        dest_original = originais / f"{src_file.stem}_{ts}{src_file.suffix}"
                        shutil.move(str(src_file), str(dest_original))
                        logger.info(f"Original movido para: {dest_original}")
                    except Exception as e:
                        logger.error(f"Erro ao mover original para 'originais': {e}")

                    # 2) mover a pasta de resultados (gerada ao lado do PDF dentro de 'em processamento') -> 'resultados'
                    try:
                        base_prefix = src_file.stem.upper()
                        moved_results = False

                        for p in processing_dir.iterdir():
                            # queremos qualquer pasta cujo nome comece com o nome do PDF (em maiúsculo)
                            if p.is_dir() and p.name.upper().startswith(base_prefix):
                                dest_result = resultados / p.name
                                if dest_result.exists():
                                    extra = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                                    dest_result = resultados / f"{p.name}_{extra}"

                                shutil.move(str(p), str(dest_result))
                                logger.info(f"Pasta de resultados movida para: {dest_result}")
                                moved_results = True
                                break

                        if not moved_results:
                            logger.warning(
                                "Não encontrei pasta de resultados ligada a '%s' dentro de 'em processamento'.",
                                src_file.name,
                            )
                    except Exception as e:
                        logger.error(f"Erro ao mover resultados para 'resultados': {e}")

                    # 3) se 'em processamento' ficou vazia, apagar (pasta temporária)
                    try:
                        if processing_dir.exists() and processing_dir.is_dir():
                            vazio = True
                            for _ in processing_dir.iterdir():
                                vazio = False
                                break
                            if vazio:
                                processing_dir.rmdir()
                                logger.info(f"Pasta temporária removida: {processing_dir}")
                    except Exception as e:
                        logger.error(f"Erro ao remover pasta temporária 'em processamento': {e}")

                else:
                    logger.error(
                        f"ERRO ao executar (ret={result.returncode}).\n"
                        f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
                    )
            finally:
                self.queue.task_done()


def main():
    global logger
    logger = setup_logger()
    if not ROOT_DIR.exists():
        logger.error(f"Diretório raiz não existe: {ROOT_DIR}")
        raise SystemExit(1)

    q = Queue()
    Worker(q).start()

    # snapshot inicial
    seen = set()
    if STARTUP_MODE == "ignore_existing":
        for _, f in list_files_recursive(ROOT_DIR):
            try:
                f_abs = f.resolve()
                mtime = f_abs.stat().st_mtime
                key = (str(f_abs), mtime)
                seen.add(key)
            except Exception:
                pass

    logger.info(f"Iniciando monitoramento em: {ROOT_DIR}")
    while True:
        try:
            for pasta_script, f in list_files_recursive(ROOT_DIR):
                f_abs = f.resolve()
                mtime = f_abs.stat().st_mtime
                key = (str(f_abs), mtime)

                # se já vimos ESTE arquivo com ESTE mtime, ignore
                if key in seen:
                    continue
                seen.add(key)

                rel_folder = get_rel_subfolder(ROOT_DIR, pasta_script)
                cmd_template = resolve_command_for_folder(rel_folder)
                if not cmd_template:
                    logger.warning(f"Nenhum script configurado para '{rel_folder}'. Ignorando {f_abs.name}.")
                    continue

                moved = move_to_processing(pasta_script, f_abs)
                if not moved:
                    # se não conseguiu mover (arquivo crescendo, erro, etc.), permita tentar de novo no próximo loop
                    if key in seen:
                        seen.remove(key)
                    continue

                q.put(Job(rel_folder, moved, cmd_template))
                logger.info(f"Novo arquivo: {rel_folder} :: {moved.name} (enfileirado)")

            time.sleep(SCAN_INTERVAL)
            logger.info("Scan concluído; aguardando nova varredura...")

        except Exception as e:
            logger.exception(f"Erro no loop principal: {e}")
            time.sleep(SCAN_INTERVAL)


if __name__ == "__main__":
    main()
