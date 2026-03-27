# -*- coding: utf-8 -*-
"""
Orquestrador ECD

Observacao: este orquestrador executa apenas a troca de empresa e os enters iniciais.
Deve ser executado APOS os demais scripts terem rodado e ANTES do Script 1
na primeira "rodada" do fluxo.
"""

from __future__ import annotations

import datetime as _dt
import os
import json
import subprocess
import sys
import time

try:
    import pyautogui as pag
except Exception as exc:
    raise SystemExit("pyautogui é obrigatório para a automação de UI") from exc

try:
    from pywinauto import Desktop
except Exception:
    Desktop = None

# =====================
# Configuracoes
# =====================
NOME_SCRIPT = "0 - Orquestrador ECD"
BASE_DIR = r"W:\\SPEDs\\ECD\\2025"
ANO_PASSADO = _dt.date.today().year - 1
LOG_PATH = os.path.join(os.path.dirname(__file__), "log_orquestrador_ecd.txt")
RUN_LOG_PATH = os.path.join(
    os.path.dirname(__file__),
    f"log_orquestrador_ecd_live_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
)
STATUS_PATH = r"W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\central-utils\\data\\ecd-status\\ecd_status.json"
POLL_SECONDS = 15
SCRIPT_1 = os.path.join(os.path.dirname(__file__), "1 - Relatório de Notas Explicativas ECD.py")
SCRIPT_2 = os.path.join(os.path.dirname(__file__), "2- Apuração das Demonstrações Contábeis para o SPED - ECD.py")
SCRIPT_3 = os.path.join(os.path.dirname(__file__), "3 - DFC - Demonstração de Fluxo de Caixa - Comparativo.py")
SCRIPT_4 = os.path.join(os.path.dirname(__file__), "4 - Escrituração Contábil Digital (ECD).py")
SCRIPT_5 = os.path.join(os.path.dirname(__file__), "5 - Formatar J150 e J930.py")
CAMPO_EMPRESA = (1637, 55)
RET_DFC_NAO_ENCONTRADA = 20


# =====================
# Utilitarios
# =====================

def _log(msg: str) -> None:
    timestamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{timestamp}] {msg}"
    print(linha)
    for p in (LOG_PATH, RUN_LOG_PATH):
        with open(p, "a", encoding="utf-8") as f:
            f.write(linha + "\n")


def _executar_cmd_ao_vivo(cmd: list[str], tag: str) -> int:
    try:
        with subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        ) as proc:
            if proc.stdout is not None:
                for linha in proc.stdout:
                    texto = (linha or "").rstrip("\r\n")
                    if texto:
                        _log(f"[{tag}] {texto}")
            rc = proc.wait()
            return int(rc)
    except Exception as exc:
        _log(f"[{tag}] Falha ao executar comando: {exc}")
        return 1


def _caminho_destino(empresa: str) -> str:
    return os.path.join(BASE_DIR, empresa)


def _carregar_status() -> dict:
    try:
        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        _log(f"Arquivo de status nao encontrado: {STATUS_PATH}")
    except json.JSONDecodeError:
        _log("Arquivo de status com JSON invalido.")
    except Exception as exc:
        _log(f"Falha ao ler status: {exc}")
    return {}


def _salvar_status(status: dict) -> bool:
    try:
        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(status, f, ensure_ascii=False, indent=2)
            f.write("\n")
        return True
    except Exception as exc:
        _log(f"Falha ao salvar status: {exc}")
        return False


def _listar_empresas_para_processar(status: dict) -> list[dict]:
    empresas: list[dict] = []
    companies = status.get("companies") or {}
    if not isinstance(companies, dict):
        return empresas
    for codigo, dados in companies.items():
        if not isinstance(dados, dict):
            continue
        if not dados.get("completed", False):
            continue
        if str(dados.get("arquivosNaPasta", "N")).strip().upper() == "S":
            continue
        nome = str(dados.get("name") or "").strip()
        if not nome:
            _log(f"Empresa sem nome no status (codigo={codigo}).")
            continue
        empresas.append(
            {
                "codigo": str(dados.get("code") or codigo).strip(),
                "nome": nome,
                "dfc": bool(dados.get("dfc", False)),
                "simples": str(dados.get("simples") or "Normal").strip(),
            }
        )
    return empresas


def _focar_janela_contabilidade() -> bool:
    if Desktop is None:
        _log("pywinauto nao disponivel. Nao foi possivel focar a janela.")
        return False
    try:
        for w in Desktop(backend="win32").windows():
            try:
                if w.class_name() == "TfrmPrincipal" and "Contabilidade" in (w.window_text() or ""):
                    w.set_focus()
                    try:
                        w.maximize()
                    except Exception:
                        pass
                    _log("Janela de Contabilidade focada e maximizada.")
                    return True
            except Exception:
                continue
    except Exception as exc:
        _log(f"Falha ao focar janela de Contabilidade: {exc}")
    _log("Janela de Contabilidade nao localizada.")
    return False


def _modo_so_focar(argv: list[str]) -> bool:
    return "--so-focar" in [a.strip().lower() for a in argv]

def _modo_so_trocar(argv: list[str]) -> bool:
    return "--so-trocar" in [a.strip().lower() for a in argv]


def _executar_script(
    caminho_script: str,
    empresa: str,
    dfc: bool | None = None,
    simples: str | None = None,
    extra_args: list[str] | None = None,
) -> int:
    if not os.path.isfile(caminho_script):
        _log(f"Script nao encontrado: {caminho_script}")
        return 1
    _log(f"Executando: {caminho_script}")
    try:
        cmd = [sys.executable, caminho_script, empresa]
        if dfc is not None:
            cmd.append("--dfc" if dfc else "--no-dfc")
        if simples:
            cmd.append(f"--simples={simples}")
        if extra_args:
            cmd.extend(extra_args)
        rc = _executar_cmd_ao_vivo(cmd, os.path.basename(caminho_script))
        _log(f"Retorno do script: {rc}")
        return rc
    except Exception as exc:
        _log(f"Falha ao executar script: {exc}")
        return 1


def _executar_script_arquivo(caminho_script: str, caminho_arquivo: str) -> int:
    if not os.path.isfile(caminho_script):
        _log(f"Script nao encontrado: {caminho_script}")
        return 1
    if not os.path.isfile(caminho_arquivo):
        _log(f"Arquivo nao encontrado para processar: {caminho_arquivo}")
        return 1
    _log(f"Executando: {caminho_script}")
    _log(f"Arquivo alvo: {caminho_arquivo}")
    try:
        cmd = [sys.executable, caminho_script, caminho_arquivo]
        rc = _executar_cmd_ao_vivo(cmd, os.path.basename(caminho_script))
        _log(f"Retorno do script: {rc}")
        return rc
    except Exception as exc:
        _log(f"Falha ao executar script: {exc}")
        return 1


def _localizar_arquivo_ecd(destino: str) -> str | None:
    try:
        txts = [f for f in os.listdir(destino) if f.lower().endswith(".txt")]
    except Exception as exc:
        _log(f"Falha ao listar TXT em {destino}: {exc}")
        return None

    candidatos = []
    for nome in txts:
        low = nome.lower()
        if low.startswith("original_") or low.startswith(".tmp_"):
            continue
        candidatos.append(os.path.join(destino, nome))

    if not candidatos:
        return None

    candidatos.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidatos[0]

def _arquivo_existe_por_nome(destino: str, nome_arquivo: str) -> bool:
    alvo = nome_arquivo.strip().lower()
    try:
        for nome in os.listdir(destino):
            if nome.strip().lower() == alvo:
                caminho = os.path.join(destino, nome)
                if os.path.isfile(caminho):
                    return True
    except Exception as exc:
        _log(f"Falha ao verificar arquivo {nome_arquivo} em {destino}: {exc}")
    return False

def _marcar_empresa_sem_dfc(status: dict, codigo: str) -> None:
    companies = status.get("companies") or {}
    if not isinstance(companies, dict):
        return
    for chave, dados in companies.items():
        if not isinstance(dados, dict):
            continue
        codigo_item = str(dados.get("code") or chave).strip()
        if codigo_item == codigo:
            dados["dfc"] = False
            companies[chave] = dados
            status["companies"] = companies
            _salvar_status(status)
            _log(f"Status atualizado: empresa {codigo} marcada com dfc=False.")
            return


# =====================
# Fluxo principal
# =====================

def _processar_empresa(codigo: str, nome: str, dfc: bool, simples: str, status: dict) -> None:
    destino = _caminho_destino(nome)
    os.makedirs(destino, exist_ok=True)
    nota_existente = _arquivo_existe_por_nome(destino, "nota_explicativa.rtf")
    dfc_existente = _arquivo_existe_por_nome(destino, "DFC.rtf")
    ecd_existente_antes = _localizar_arquivo_ecd(destino)

    _focar_janela_contabilidade()
    _log("Selecionando caixa de empresa.")
    pag.click(CAMPO_EMPRESA[0], CAMPO_EMPRESA[1])
    pag.press("backspace", presses=5)
    pag.write(codigo, interval=0.02)
    pag.press("enter", presses=2)
    _log(f"Iniciando: {NOME_SCRIPT}")
    _log(f"Empresa: {nome}")
    _log(f"Codigo: {codigo}")
    _log(f"DFC: {'sim' if dfc else 'nao'} | Simples: {simples}")
    _log(f"Diretorio destino: {destino}")

    passos = [
        f"Trocar empresa no sistema usando o codigo: {codigo}.",
        "Ir ate a aba 'contabil'.",
        "Selecionar a area da empresa, enviar 5 backspaces e digitar a empresa atual.",
        "Enviar dois enters para confirmar.",
    ]

    for idx, passo in enumerate(passos, start=1):
        _log(f"Passo {idx}: {passo}")

    vai_pular_1 = nota_existente
    vai_pular_3 = (not dfc) or dfc_existente
    vai_pular_2 = vai_pular_1 and vai_pular_3

    rc1 = 0
    if vai_pular_1:
        _log("nota_explicativa.rtf ja existe. Pulando script 1.")
    else:
        rc1 = _executar_script(SCRIPT_1, nome)

    rc2 = 0
    if vai_pular_2:
        _log("Scripts 1 e 3 serao pulados. Pulando script 2.")
    else:
        rc2 = _executar_script(SCRIPT_2, nome)

    rc3 = 0
    if dfc:
        if dfc_existente:
            _log("DFC.rtf ja existe. Pulando script 3.")
        else:
            rc3 = _executar_script(SCRIPT_3, nome)
            if rc3 == RET_DFC_NAO_ENCONTRADA:
                _log("Script 3 informou ausencia de estrutura DFC. Continuando com dfc=False.")
                dfc = False
                _marcar_empresa_sem_dfc(status, codigo)
                rc3 = 0
    else:
        _log("DFC falso: pulando script 3.")

    extra_args_script4: list[str] = []
    if ecd_existente_antes:
        _log(f"ECD TXT ja existente detectado. Script 4 em modo somente movimento: {ecd_existente_antes}")
        extra_args_script4.append("--somente-movimento")
    rc4 = _executar_script(SCRIPT_4, nome, dfc=dfc, simples=simples, extra_args=extra_args_script4)

    if any(rc != 0 for rc in (rc1, rc2, rc3, rc4)):
        _log("Falha em um ou mais scripts. Status nao sera atualizado.")
        return

    caminho_ecd = ecd_existente_antes or _localizar_arquivo_ecd(destino)
    if not caminho_ecd:
        _log("Nenhum TXT de ECD localizado para formatar.")
        return
    rc5 = _executar_script_arquivo(SCRIPT_5, caminho_ecd)
    if rc5 != 0:
        _log("Falha ao executar o script 5. Status nao sera atualizado.")
        return

    rtf_necessarios = ["Nota_explicativa.rtf"]
    if dfc:
        rtf_necessarios.append("DFC.rtf")
    faltando = [arq for arq in rtf_necessarios if not os.path.isfile(os.path.join(destino, arq))]
    txts = [f for f in os.listdir(destino) if f.lower().endswith(".txt")]
    if faltando or not txts:
        if faltando:
            _log(f"Arquivos RTF faltando: {', '.join(faltando)}")
        if not txts:
            _log("Nenhum arquivo TXT encontrado na pasta.")
        _log("Nao marcou como concluido. Verificar pasta da empresa.")
        return

    companies = status.get("companies") or {}
    for chave, dados in companies.items():
        if not isinstance(dados, dict):
            continue
        if str(dados.get("code") or chave).strip() == codigo:
            dados["arquivosNaPasta"] = "S"
            break
    _salvar_status(status)
    _log("Finalizado (sem automacao de UI).")


def _agora() -> _dt.datetime:
    return _dt.datetime.now()


def _dentro_janela_execucao(agora: _dt.datetime) -> bool:
    t = agora.time()
    return (t >= _dt.time(8, 0) and t < _dt.time(11, 45)) or (t >= _dt.time(13, 30) and t < _dt.time(17, 50))


def _proxima_janela(agora: _dt.datetime) -> _dt.datetime:
    hoje = agora.date()
    j1_ini = _dt.datetime.combine(hoje, _dt.time(8, 0))
    j1_fim = _dt.datetime.combine(hoje, _dt.time(11, 45))
    j2_ini = _dt.datetime.combine(hoje, _dt.time(13, 30))
    j2_fim = _dt.datetime.combine(hoje, _dt.time(17, 50))

    if agora < j1_ini:
        return j1_ini
    if agora < j1_fim:
        return agora
    if agora < j2_ini:
        return j2_ini
    if agora < j2_fim:
        return agora
    amanha = hoje + _dt.timedelta(days=1)
    return _dt.datetime.combine(amanha, _dt.time(8, 0))


def _sleep_ate(when: _dt.datetime) -> None:
    delta = (when - _agora()).total_seconds()
    if delta > 0:
        _log(f"Aguardando janela de execucao: {int(delta)}s")
        time.sleep(delta)


def main() -> int:
    _log(f"Log ao vivo desta execucao: {RUN_LOG_PATH}")
    if _modo_so_focar(sys.argv):
        _log("Modo so-focar ativo.")
        _focar_janela_contabilidade()
        return 0

    if _modo_so_trocar(sys.argv):
        _log("Modo so-trocar ativo.")
        status = _carregar_status()
        empresas = _listar_empresas_para_processar(status)
        if not empresas:
            _log("Nenhuma empresa pendente para processar.")
            return 1
        empresa = empresas[0]
        _focar_janela_contabilidade()
        _log("Selecionando caixa de empresa.")
        pag.click(CAMPO_EMPRESA[0], CAMPO_EMPRESA[1])
        pag.press("backspace", presses=5)
        pag.write(empresa["codigo"], interval=0.02)
        pag.press("enter", presses=2)
        _log(f"Empresa trocada: {empresa['codigo']} - {empresa['nome']}")
        return 0

    _log(f"Iniciando: {NOME_SCRIPT}")
    _log(f"Arquivo de status: {STATUS_PATH}")
    _log(f"Intervalo de verificação: {POLL_SECONDS}s")
    _log("Janelas: 08:00-11:45 e 13:30-17:50")

    while True:
        agora = _agora()
        if not _dentro_janela_execucao(agora):
            _sleep_ate(_proxima_janela(agora))
            continue
        status = _carregar_status()
        empresas = _listar_empresas_para_processar(status)
        if not empresas:
            _log("Nenhuma empresa pendente para processar.")
        for empresa in empresas:
            if not _dentro_janela_execucao(_agora()):
                _log("Janela encerrada. Finalizando apos empresa atual.")
                break
            _processar_empresa(empresa["codigo"], empresa["nome"], empresa["dfc"], empresa["simples"], status)

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    raise SystemExit(main())


