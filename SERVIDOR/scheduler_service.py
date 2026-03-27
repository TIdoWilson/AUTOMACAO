import os
import sys
import time
import signal
import logging
import subprocess
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

# --------- LOG ----------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger("jobs")

ROOT = r"C:\ROBOS"
PY   = sys.executable  # usa o mesmo Python do venv/serviço

BUSCA_NOTAS = os.path.join(ROOT, "Servidor - Busca_Notas_5.0.py")
EMITIR_ISS  = os.path.join(ROOT, "Servidor - Emitir_ISS.py")

def run_script(script_path):
    cmd = [PY, script_path]
    log.info(f"Executando: {cmd}")
    try:
        # stdout/err vão para o log do serviço (WinSW captura); ajuste se quiser arquivo dedicado
        result = subprocess.run(cmd, capture_output=True, text=True)
        log.info(f"RET={result.returncode}")
        if result.stdout:
            log.info(f"STDOUT:\n{result.stdout.strip()}")
        if result.stderr:
            log.warning(f"STDERR:\n{result.stderr.strip()}")
    except Exception as e:
        log.exception(f"Falha ao executar {script_path}: {e}")

def job_busca_notas():
    run_script(BUSCA_NOTAS)

def job_emitir_iss():
    run_script(EMITIR_ISS)

stop_flag = False
def handle_stop(signum, frame):
    global stop_flag
    log.info(f"Sinal {signum} recebido. Encerrando…")
    stop_flag = True

signal.signal(signal.SIGINT, handle_stop)
signal.signal(signal.SIGTERM, handle_stop)

def main():
    # ISS_DAY: por padrão 5 (pode variar conforme sua necessidade)
    iss_day = int(os.getenv("ISS_DAY", "5"))

    sched = BackgroundScheduler(timezone="America/Sao_Paulo")

    # 01 de cada mês às 00:00 — Busca Notas
    sched.add_job(
        job_busca_notas,
        CronTrigger(day="1", hour=0, minute=0),
        id="busca_notas_mensal",
        max_instances=1, coalesce=True, misfire_grace_time=3600
    )

    # [variável] dia X de cada mês às 00:00 — Emitir ISS
    sched.add_job(
        job_emitir_iss,
        CronTrigger(day=str(iss_day), hour=0, minute=0),
        id="emitir_iss_mensal",
        max_instances=1, coalesce=True, misfire_grace_time=3600
    )

    sched.start()
    log.info("Agendador iniciado. Jobs carregados.")

    try:
        while not stop_flag:
            time.sleep(1)
    finally:
        log.info("Parando agendador…")
        sched.shutdown(wait=True)
        log.info("Agendador finalizado.")

if __name__ == "__main__":
    main()
