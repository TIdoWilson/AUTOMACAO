import asyncio
import json
import os
import re
import unicodedata
import traceback
from datetime import datetime, timedelta
from base64 import urlsafe_b64decode
import csv

import httpx

# ======================
# CONFIGURAÇÕES GERAIS
# ======================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HEADERS_FILE = os.path.join(BASE_DIR, "headers_iob.json")

USER_DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DOWNLOAD_DIR = os.path.join(USER_DOWNLOADS_DIR, "RECIBOS_SPED")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

LOG_FILE = os.path.join(DOWNLOAD_DIR, "log.txt")

DOCS_BASE = "https://docs.iob.com.br/api"

# Limites de paralelismo
MAX_CONC_WORKFLOWS = 10   # quantos workflows processar ao mesmo tempo
MAX_RETRIES = 2           # tentativas para chamadas gerais (lista, fullDetails, etc.)
MAX_DOWNLOAD_RETRIES = 5  # tentativas específicas para download de PDF

# ======================
# LOG
# ======================

_LOG_CTX = {"item": "", "empresa": ""}


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]


def set_ctx(item: str = "", empresa: str = ""):
    _LOG_CTX["item"] = item
    _LOG_CTX["empresa"] = empresa


def log(msg: str):
    line = f"[{_ts()}][{_LOG_CTX.get('item','')}]({_LOG_CTX.get('empresa','')}) {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def log_exc(prefix: str = "EXCEPTION"):
    tb = traceback.format_exc()
    log(f"{prefix}: {tb.strip()}")


# ======================
# FUNÇÕES AUXILIARES
# ======================

def strip_accents(s: str) -> str:
    return unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode("ascii")


def clean_filename_strict(name: str) -> str:
    name = strip_accents(name).strip()
    for ch in r'<>:"/\|?*':
        name = name.replace(ch, "_")
    name = re.sub(r"\s+", " ", name).strip()
    return name or "arquivo"


def periodo_mes_anterior() -> str:
    """
    Retorna o período do mês anterior no formato 'MMyyyy' (ex: 102025).
    """
    today = datetime.today()
    first = today.replace(day=1)
    prev = first - timedelta(days=1)
    return f"{prev.month:02d}{prev.year}"


def strip_conditional_headers(headers: dict) -> dict:
    """
    Remove If-None-Match / If-Modified-Since (case-insensitive).
    Útil antes de chamar /fullDetails e /documents/.../download.
    """
    h = dict(headers)
    to_remove = []
    for k in h.keys():
        if k.lower() in ("if-none-match", "if-modified-since"):
            to_remove.append(k)
    for k in to_remove:
        h.pop(k, None)
    return h


def load_headers_from_file(path: str) -> dict:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Arquivo de headers não encontrado: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # Pequeno sanity check
    if "authorization" not in data and "Authorization" not in data:
        log("ATENÇÃO: header 'authorization' não encontrado no JSON.")

    # Não queremos mandar cabeçalhos condicionais na API
    data = strip_conditional_headers(data)

    return data


def decode_jwt_info(token: str):
    """
    Só para logar iat/exp do token (sem validação criptográfica).
    """
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return None
        payload_b64 = parts[1]
        # Ajusta padding
        padding = "=" * (-len(payload_b64) % 4)
        payload_b64 += padding
        data = json.loads(urlsafe_b64decode(payload_b64.encode("ascii")))
        return data
    except Exception:
        return None


def is_token_expired(token: str, skew_seconds: int = 60) -> bool:
    """
    Retorna True se o token já expirou ou vai expirar nos próximos skew_seconds.
    Se não conseguir ler o exp, assume expirado (fail safe).
    """
    info = decode_jwt_info(token)
    if not info or "exp" not in info:
        return True
    now_ts = datetime.now().timestamp()
    exp_ts = float(info["exp"])
    return exp_ts <= (now_ts + skew_seconds)


def log_token_expirations(headers: dict):
    # authorization (Bearer ...)
    auth = headers.get("authorization") or headers.get("Authorization")
    if auth and auth.lower().startswith("bearer "):
        jwt_ = auth.split(" ", 1)[1].strip()
        info = decode_jwt_info(jwt_)
        if info and "exp" in info and "iat" in info:
            iat = datetime.fromtimestamp(info["iat"])
            exp = datetime.fromtimestamp(info["exp"])
            log(
                f"Token 'authorization' iat={iat}, exp={exp} "
                f"(duração ~{(exp - iat).total_seconds()/60:.1f} minutos)."
            )

    # x-hypercube-idp-access-token
    h_token = headers.get("x-hypercube-idp-access-token")
    if h_token:
        info = decode_jwt_info(h_token)
        if info and "exp" in info and "iat" in info:
            iat = datetime.fromtimestamp(info["iat"])
            exp = datetime.fromtimestamp(info["exp"])
            log(
                f"Token 'x-hypercube-idp-access-token' iat={iat}, exp={exp} "
                f"(duração ~{(exp - iat).total_seconds()/60:.1f} minutos)."
            )


def abort_if_tokens_expired(headers: dict, skew_seconds: int = 60):
    """
    Verifica se algum dos tokens relevantes já expirou ou está prestes a expirar
    e aborta a execução se estiver.
    """
    expired_names = []

    # authorization (Bearer ...)
    auth = headers.get("authorization") or headers.get("Authorization")
    if auth and auth.lower().startswith("bearer "):
        jwt_ = auth.split(" ", 1)[1].strip()
        if is_token_expired(jwt_, skew_seconds=skew_seconds):
            expired_names.append("authorization")

    # x-hypercube-idp-access-token
    h_token = headers.get("x-hypercube-idp-access-token")
    if h_token and is_token_expired(h_token, skew_seconds=skew_seconds):
        expired_names.append("x-hypercube-idp-access-token")

    if expired_names:
        log(
            f"ATENÇÃO: token(s) expirado(s) ou prestes a expirar: {', '.join(expired_names)}. "
            f"Abortando para não continuar com credencial velha."
        )
        raise RuntimeError(
            "Token(s) expirado(s) ou inválido(s). Atualize o headers_iob.json antes de rodar novamente."
        )


# ======================
# RELATÓRIO DE FALHAS
# ======================

FAILED_DOWNLOADS: list[dict] = []


def register_failed_download(cnpj: str, empresa: str, period: str, document_id: str | int, output_path: str):
    FAILED_DOWNLOADS.append(
        {
            "cnpj": cnpj,
            "empresa": empresa,
            "periodo": period,
            "document_id": str(document_id),
            "output_path": output_path,
        }
    )


def save_failures(base_dir: str = DOWNLOAD_DIR):
    """
    Salva as falhas acumuladas em JSON e CSV.
    """
    if not FAILED_DOWNLOADS:
        log("Nenhum download falhou. Nenhum relatório de falhas gerado.")
        return

    os.makedirs(base_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    csv_path = os.path.join(base_dir, f"falhas_download_{timestamp}.csv")
    json_path = os.path.join(base_dir, f"falhas_download_{timestamp}.json")

    # JSON
    try:
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(FAILED_DOWNLOADS, jf, ensure_ascii=False, indent=2)
        log(f"Relatório de falhas (JSON) salvo em: {json_path}")
    except Exception as e:
        log(f"Erro ao salvar relatório de falhas JSON: {e}")

    # CSV
    fieldnames = ["cnpj", "empresa", "periodo", "document_id", "output_path"]
    try:
        with open(csv_path, "w", encoding="utf-8", newline="") as cf:
            writer = csv.DictWriter(cf, fieldnames=fieldnames)
            writer.writeheader()
            for item in FAILED_DOWNLOADS:
                row = {key: item.get(key, "") for key in fieldnames}
                writer.writerow(row)
        log(f"Relatório de falhas (CSV) salvo em: {csv_path}")
    except Exception as e:
        log(f"Erro ao salvar relatório de falhas CSV: {e}")


# ======================
# EXTRAÇÃO / PDF
# ======================

def extract_recibo_document_from_workflow(wf_json: dict) -> dict | None:
    """
    A partir do JSON de /api/workflows/processes/{id}?fullDetails=true,
    encontra o Document que corresponde ao Recibo.
    """
    if not wf_json:
        return None

    states = wf_json.get("States") or []
    best_doc = None
    best_score = 0

    for st in states:
        attachments = st.get("Attachments") or []
        for att in attachments:
            doc = (att or {}).get("Document") or {}
            name = doc.get("name") or ""
            ext = (doc.get("extension") or "").lower()
            mime = (doc.get("mimeType") or "").lower()
            is_pdf = doc.get("isPdf") or ext == "pdf" or "application/pdf" in mime
            if not is_pdf:
                continue

            norm = strip_accents(name).lower()
            score = 0

            if "recibo" in norm and "entrega" in norm:
                score = 100
            elif "recibo" in norm:
                score = 80

            if "comprovante" in norm:
                score = max(score, 70)
                if "entrega" in norm or "transmissao" in norm:
                    score = max(score, 90)

            if "efd_contribuicoes" in norm:
                score += 5

            if score > best_score:
                best_score = score
                best_doc = doc

    return best_doc


TRANSIENT_STATUS_CODES = {502, 503, 504}


async def download_pdf_document(
    client: httpx.AsyncClient,
    document_id: int | str,
    filename_hint: str,
    empresa: str,
    cnpj: str,
    period: str,
    max_retries: int = MAX_DOWNLOAD_RETRIES
) -> str | None:
    """
    Faz GET em /documents/{id}/download e salva como PDF.
    Nome usa empresa + cnpj + período.
    """
    try:
        doc_id = str(document_id)
    except Exception:
        log(f"download_pdf_document: document_id inválido: {document_id!r}")
        return None

    url = f"{DOCS_BASE}/documents/{doc_id}/download"

    # Monta nome do arquivo com empresa + CNPJ + período
    clean_emp = clean_filename_strict(empresa)
    cnpj = (cnpj or "").strip() or "SEM_CNPJ"
    base_stem = f"{clean_emp} - {period} EFD CONTRIBUICOES"
    base_name = base_stem + ".pdf"
    full_pdf_path = os.path.join(DOWNLOAD_DIR, base_name)

    for attempt in range(1, max_retries + 1):
        log(f"download_pdf_document: GET {url} (tentativa {attempt}/{max_retries})")

        # Garante que não vamos mandar If-None-Match / If-Modified-Since nesta chamada
        req_headers = strip_conditional_headers(dict(client.headers))

        try:
            resp = await client.get(url, headers=req_headers)
        except Exception as e:
            log(f"download_pdf_document: erro de rede: {e}")
            if attempt < max_retries:
                delay = 2 ** (attempt - 1)
                log(f"download_pdf_document: aguardando {delay:.1f}s antes de nova tentativa...")
                await asyncio.sleep(delay)
                continue
            return None

        ct = (resp.headers.get("content-type") or "").lower()
        log(f"download_pdf_document: status={resp.status_code}, content-type={ct}")

        if not resp.is_success:
            body_snip = resp.text[:300]
            log(f"download_pdf_document: HTTP {resp.status_code}. Corpo (início): {body_snip}")

            # Erros transitórios específicos (502/503/504) -> re-tenta com backoff
            if resp.status_code in TRANSIENT_STATUS_CODES and attempt < max_retries:
                delay = 2 ** (attempt - 1)
                log(
                    f"download_pdf_document: erro transitório ({resp.status_code}). "
                    f"Aguardando {delay:.1f}s antes de nova tentativa..."
                )
                await asyncio.sleep(delay)
                continue

            # Demais erros -> não adianta insistir
            return None

        content = resp.content or b""
        size = len(content)
        log(f"download_pdf_document: recebidos {size} bytes.")

        magic_pdf = content.startswith(b"%PDF-")
        is_pdf_ct = "application/pdf" in ct or "application/octet-stream" in ct

        if not (magic_pdf or is_pdf_ct):
            preview = (content[:200] or b"").decode("latin1", errors="ignore")
            log(
                f"download_pdf_document: conteúdo não parece PDF. "
                f"Magic={magic_pdf}, content-type={ct}, preview={preview[:120]!r}"
            )
            if attempt < max_retries:
                delay = 2 ** (attempt - 1)
                log(f"download_pdf_document: aguardando {delay:.1f}s antes de nova tentativa...")
                await asyncio.sleep(delay)
                continue
            # Não salva nada neste caso
            return None

        try:
            with open(full_pdf_path, "wb") as f:
                f.write(content)
            log(f"download_pdf_document: PDF salvo em {full_pdf_path}")
            return full_pdf_path
        except Exception as e:
            log(f"download_pdf_document: erro ao salvar arquivo: {e}")
            if attempt < max_retries:
                delay = 2 ** (attempt - 1)
                log(f"download_pdf_document: aguardando {delay:.1f}s antes de nova tentativa...")
                await asyncio.sleep(delay)
                continue
            return None

    return None


# ======================
# WORKFLOWS
# ======================

async def fetch_workflows_list(client: httpx.AsyncClient) -> list[dict]:
    """
    Busca a lista de workflows concluídos (state 2 / requestStateType 3).
    """
    url = f"{DOCS_BASE}/workflows/processes?processState=2&requestStateType=3"
    log(f"Requisitando lista de workflows em: {url}")

    for attempt in range(1, MAX_RETRIES + 1):
        # Garante que não vamos mandar If-None-Match / If-Modified-Since
        req_headers = strip_conditional_headers(dict(client.headers))

        try:
            resp = await client.get(url, headers=req_headers)
        except Exception as e:
            log(f"Erro ao requisitar workflows (tentativa {attempt}): {e}")
            if attempt < MAX_RETRIES:
                await asyncio.sleep(1.0)
                continue
            return []

        if not resp.is_success:
            log(f"HTTP {resp.status_code} ao requisitar workflows (tentativa {attempt}).")
            if attempt < MAX_RETRIES and resp.status_code in TRANSIENT_STATUS_CODES:
                await asyncio.sleep(1.0)
                continue
            return []

        try:
            data = resp.json()
        except Exception as e:
            log(f"Erro ao parsear JSON da lista de workflows: {e}")
            return []

        workflows = []
        if isinstance(data, list):
            workflows = data
        elif isinstance(data, dict):
            for key in ("items", "data", "results", "Processes", "workflows", "value", "requests"):
                v = data.get(key)
                if isinstance(v, list):
                    workflows = v
                    break

        log(f"Total de workflows retornados pela API: {len(workflows)}")
        return workflows

    return []


async def fetch_full_details(client: httpx.AsyncClient, wf_id: str) -> dict | None:
    url = f"{DOCS_BASE}/workflows/processes/{wf_id}?fullDetails=true"
    for attempt in range(1, MAX_RETRIES + 1):
        # Garante que não vamos mandar If-None-Match / If-Modified-Since
        req_headers = strip_conditional_headers(dict(client.headers))

        try:
            resp = await client.get(url, headers=req_headers)
        except Exception as e:
            log(f"Erro ao chamar fullDetails ({wf_id}) tentativa {attempt}: {e}")
            if attempt < MAX_RETRIES:
                await asyncio.sleep(1.0)
                continue
            return None

        if resp.status_code in TRANSIENT_STATUS_CODES:
            log(f"fullDetails retornou HTTP {resp.status_code} para {wf_id}.")
            if attempt < MAX_RETRIES:
                await asyncio.sleep(1.0)
                continue
            return None

        if not resp.is_success:
            log(f"fullDetails retornou HTTP {resp.status_code} para {wf_id}.")
            return None

        try:
            return resp.json()
        except Exception as e:
            log(f"Erro ao parsear JSON de fullDetails para {wf_id}: {e}")
            if attempt < MAX_RETRIES:
                await asyncio.sleep(0.5)
                continue
            return None

    return None


async def process_single_workflow(
    sem: asyncio.Semaphore,
    client: httpx.AsyncClient,
    idx: int,
    wf: dict,
    alvo_period: str
):
    """
    Processa um workflow: valida tipo/período, acha recibo e baixa PDF.
    """
    wid = wf.get("id") or wf.get("sanid") or "?"
    state = wf.get("state")
    set_ctx(item=f"wf#{idx}", empresa="")

    log(f"Processando workflow id={wid}, state={state}")

    if state != 2:
        log("Ignorando workflow (state != 2).")
        return

    # Tentativa de extrair nome / cnpj / período já do stub (Request)
    req = wf.get("Request") or {}
    client_info = req.get("Client") or wf.get("Client") or {}
    empresa = client_info.get("fullName") or client_info.get("firstName") or ""
    set_ctx(item=f"wf#{idx}", empresa=empresa)

    documents = (req.get("Documents") or [])
    doc0 = documents[0] if documents else {}

    category = doc0.get("category")
    doc_type = doc0.get("type")
    metadata = ((doc0.get("metadata") or {}).get("fileMetadata") or {})
    period_meta = metadata.get("period") or ""
    cnpj = metadata.get("cnpj") or metadata.get("taxpayerNumber") or ""

    if category != "SPED" or doc_type != "EFD_CONTRIBUICOES":
        log("Workflow não é SPED EFD_CONTRIBUICOES; ignorando.")
        return

    if period_meta != alvo_period:
        log(f"Período {period_meta} diferente do alvo {alvo_period}; ignorando.")
        return

    log(f"Workflow válido para {empresa} (CNPJ={cnpj}, período={period_meta}).")

    async with sem:
        wf_json = await fetch_full_details(client, wid)
        if not wf_json:
            log("Não foi possível obter fullDetails; pulando workflow.")
            # Registra falha de forma genérica (sem document_id)
            register_failed_download(cnpj, empresa, period_meta, f"workflow:{wid}", "")
            return

        recibo_doc = extract_recibo_document_from_workflow(wf_json)
        if not recibo_doc:
            log("Não foi possível localizar documento de Recibo neste workflow.")
            register_failed_download(cnpj, empresa, period_meta, f"workflow:{wid}", "")
            return

        doc_id = recibo_doc.get("id")
        doc_name = recibo_doc.get("name") or f"Recibo_{wid}.pdf"
        log(f"Documento de recibo identificado: id={doc_id}, name={doc_name}")

        saved_path = await download_pdf_document(
            client=client,
            document_id=doc_id,
            filename_hint=doc_name,
            empresa=empresa,
            cnpj=cnpj,
            period=period_meta,
            max_retries=MAX_DOWNLOAD_RETRIES
        )

        if saved_path:
            log(f"PDF salvo com sucesso: {saved_path}")
        else:
            log("Falha ao salvar PDF via API para este workflow.")
            register_failed_download(cnpj, empresa, period_meta, doc_id, "")


# ======================
# MAIN
# ======================

async def main():
    # Início da sessão de log
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80 + f"\nSessão: {_ts()}\n" + "=" * 80 + "\n")
    except Exception:
        pass

    headers = load_headers_from_file(HEADERS_FILE)
    log("Headers carregados do arquivo JSON.")
    log_token_expirations(headers)

    # Não continuar se o token estiver expirado / prestes a expirar
    abort_if_tokens_expired(headers, skew_seconds=60)

    alvo_period = periodo_mes_anterior()
    log(f"Período alvo (metadata.fileMetadata.period) = {alvo_period}")

    async with httpx.AsyncClient(headers=headers, timeout=20.0, follow_redirects=True) as client:
        workflows = await fetch_workflows_list(client)
        if not workflows:
            log("Nenhum workflow retornado pela API. Encerrando.")
            # Mesmo assim tenta salvar relatório de falhas (se houver)
            save_failures()
            return

        sem = asyncio.Semaphore(MAX_CONC_WORKFLOWS)

        tasks = []
        for idx, wf in enumerate(workflows, start=1):
            tasks.append(
                asyncio.create_task(
                    process_single_workflow(sem, client, idx, wf, alvo_period)
                )
            )

        await asyncio.gather(*tasks)

    # Ao final, gera relatório de falhas (se houver)
    save_failures()

    log("Processamento finalizado.")
    print(f"Log salvo em: {LOG_FILE}")


if __name__ == "__main__":
    asyncio.run(main())
