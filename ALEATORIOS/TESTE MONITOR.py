#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import time
import shutil
import threading
import subprocess
from pathlib import Path
from queue import Queue, Empty
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler



# ========= CONFIG =========
# Pasta raiz a monitorar
ROOT_DIR = Path(r"C:\Users\Usuario\Desktop\python")

# Script que processa o PDF (recebe o caminho do PDF como 1º argumento)
SCRIPT_MAP = {
    "*": r'py -3 "W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\RH\SEPARAR FERIAS POR FUNCIONARIO.PY" "{file}"'
    # Se preferir: 'python "...PY" "{file}"'
}


LOG_DIR = ROOT_DIR  # ou Path(r"C:\Logs")
LOG_FILE = LOG_DIR / "monitor.log"
LOG_DIR.mkdir(parents=True, exist_ok=True)

handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=5, encoding="utf-8")
logging.basicConfig(
    level=logging.INFO,
    handlers=[handler],
    format="%(asctime)s [%(levelname)s] %(message)s",
)

def log(msg: str):
    logging.info(msg)

# se quiser registrar tracebacks:
def log_err(msg: str, exc: Exception | None = None):
    if exc:
        logging.exception(f"{msg} :: {exc}")
    else:
        logging.error(msg)

# Intervalo de varredura (s)
SCAN_INTERVAL = 60

# Considerar arquivos já existentes ao iniciar?
# "ignore_existing" ou "process_existing"
STARTUP_MODE = "ignore_existing"

# Processar somente PDFs
ALLOW_EXTENSIONS = {".pdf"}

# Nível de logs: DEBUG/INFO/WARNING/ERROR
LOG_LEVEL = logging.INFO

# Pasta e arquivo de log
LOG_DIR = ROOT_DIR / "_logs"
LOG_FILE = LOG_DIR / "monitor.log"

# Nome da subpasta intermediária
PROCESSING_DIRNAME = "em processamento"
# ========= FIM CONFIG =========


def setup_logger():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("monitor")
    logger.setLevel(LOG_LEVEL)

    # Rotaciona em ~5 MB, mantém 5 arquivos
    file_handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
    file_fmt = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    file_handler.setFormatter(file_fmt)
    logger.addHandler(file_handler)

    # Console também (útil para testes manuais)
    console = logging.StreamHandler()
    console.setFormatter(file_fmt)
    logger.addHandler(console)

    return logger


log = None  # será setado no main()


def list_files_recursive(root: Path):
    """Lista (pasta, arquivo) para todos os arquivos nas subpastas (exceto 'em processamento')."""
    for sub in root.rglob("*"):
        if not sub.is_dir():
            parent = sub.parent
            # ignora qualquer coisa dentro de "em processamento"
            if parent.name == PROCESSING_DIRNAME or PROCESSING_DIRNAME in [p.name for p in parent.parents]:
                continue
            # filtro de extensão
            if ALLOW_EXTENSIONS and sub.suffix.lower() not in ALLOW_EXTENSIONS:
                continue
            yield parent, sub


def get_rel_subfolder(root: Path, folder: Path) -> str:
    try:
        rel = folder.relative_to(root)
    except ValueError:
        rel = folder
    return str(rel).replace("\\", "/")


def resolve_command_for_folder(rel_subfolder: str) -> str | None:
    if rel_subfolder in SCRIPT_MAP:
        return SCRIPT_MAP[rel_subfolder]
    top = rel_subfolder.split("/", 1)[0] if rel_subfolder else rel_subfolder
    if top in SCRIPT_MAP:
        return SCRIPT_MAP[top]
    return SCRIPT_MAP.get("*")


def move_to_processing(folder: Path, file_path: Path) -> Path | None:
    """Move o arquivo para 'em processamento' dentro da mesma subpasta."""
    processing_dir = folder / PROCESSING_DIRNAME
    processing_dir.mkdir(parents=True, exist_ok=True)
    target = processing_dir / file_path.name

    try:
        # tentativa simples de evitar pegar arquivo ainda sendo gravado
        size1 = file_path.stat().st_size
        time.sleep(0.5)
        size2 = file_path.stat().st_size
        if size1 != size2:
            log.info(f"Arquivo ainda crescendo; vai tentar de novo na próxima varredura: {file_path}")
            return None

        shutil.move(str(file_path), str(target))
        return target
    except Exception as e:
        log.error(f"Falha ao mover {file_path} -> {target}: {e}")
        return None


class Job:
    def __init__(self, rel_folder: str, file_in_processing: Path, cmd_template: str):
        self.rel_folder = rel_folder
        self.file_in_processing = file_in_processing
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
            src_file = Path(job.file_in_processing)  # dentro de ".../em processamento/<arquivo>"
            # A subpasta monitorada é o pai de "em processamento"
            subfolder = src_file.parent.parent
            finalizados = subfolder / "finalizados"
            resultados = subfolder / "resultados"
            finalizados.mkdir(exist_ok=True)
            resultados.mkdir(exist_ok=True)

            try:
                log.info(f"Executando: {cmd}")
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True)

                if result.returncode == 0:
                    log.info(f"OK (ret={result.returncode}).\nSTDOUT:\n{result.stdout.strip()}")
                    if result.stderr.strip():
                        log.info(f"STDERR (não fatal):\n{result.stderr.strip()}")

                    # === Pós-processamento ===
                    # 1) mover o PDF ORIGINAL (de 'em processamento') -> 'finalizados'
                    try:
                        dest_original = finalizados / src_file.name
                        shutil.move(str(src_file), str(dest_original))
                        log.info(f"Original movido para: {dest_original}")
                    except Exception as e:
                        log.error(f"Erro ao mover original para 'finalizados': {e}")

                    # 2) achar a pasta "FERIAS - <empresa>" gerada ao lado do PDF (dentro de 'em processamento')
                    try:
                        parent = src_file.parent  # 'em processamento'
                        moved_results = False
                        for p in parent.glob("FERIAS - *"):
                            if p.is_dir():
                                dest_result = resultados / p.name
                                # Se já existir destino, gera nome único
                                i = 1
                                base_name = p.name
                                while dest_result.exists():
                                    i += 1
                                    dest_result = resultados / f"{base_name} {i}"
                                shutil.move(str(p), str(dest_result))
                                log.info(f"Pasta de resultados movida para: {dest_result}")
                                moved_results = True
                                break
                        if not moved_results:
                            log.warning("Não encontrei pasta 'FERIAS - <empresa>' para mover.")
                    except Exception as e:
                        log.error(f"Erro ao mover resultados para 'resultados': {e}")

                else:
                    log.error(
                        f"ERRO ao executar (ret={result.returncode}).\n"
                        f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
                    )
            except Exception as e:
                log.exception(f"Falha ao executar job: {e}")
            finally:
                self.queue.task_done()


def main():
    global log
    logger = setup_logger()
    log = logger.info  # atalho simples para mensagens informativas
    logging.getLogger("monitor")  # para .exception e .error

    if not ROOT_DIR.exists():
        logging.getLogger("monitor").error(f"Diretório raiz não existe: {ROOT_DIR}")
        raise SystemExit(1)

    q = Queue()
    worker = Worker(q)
    worker.start()

    seen = set()
    if STARTUP_MODE == "ignore_existing":
        for _, f in list_files_recursive(ROOT_DIR):
            try:
                seen.add(f.resolve())
            except Exception:
                pass

    logging.getLogger("monitor").info(f"Iniciando monitoramento em: {ROOT_DIR}")
    while True:
        try:
            for folder, f in list_files_recursive(ROOT_DIR):
                f_abs = f.resolve()
                if f_abs in seen:
                    continue
                seen.add(f_abs)

                rel_folder = get_rel_subfolder(ROOT_DIR, folder)
                cmd_template = resolve_command_for_folder(rel_folder)
                if not cmd_template:
                    logging.getLogger("monitor").warning(f"Nenhum script configurado para '{rel_folder}'. Ignorando {f_abs.name}.")
                    continue

                moved = move_to_processing(folder, f_abs)
                if not moved:
                    # não conseguiu mover (talvez em gravação); retira dos 'vistos' para tentar na próxima
                    if f_abs in seen:
                        seen.remove(f_abs)
                    continue

                job = Job(rel_folder, moved, cmd_template)
                q.put(job)
                logging.getLogger("monitor").info(f"Novo arquivo detectado -> movido e enfileirado: {rel_folder} :: {moved.name}")

            time.sleep(SCAN_INTERVAL)

        except KeyboardInterrupt:
            logging.getLogger("monitor").info("Encerrando por KeyboardInterrupt…")
            break
        except Exception as e:
            logging.getLogger("monitor").exception(f"Erro no loop principal: {e}")
            time.sleep(SCAN_INTERVAL)


if __name__ == "__main__":
    main()
