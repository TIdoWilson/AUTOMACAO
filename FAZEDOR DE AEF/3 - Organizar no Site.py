# 3 - Organizar no Site

# =========================
# Configuracoes
# =========================

import os
import sys
import time
import argparse
from datetime import datetime
from pathlib import Path
import re


BASE_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\FAZEDOR DE AEF"
PASTA_ARQUIVOS = os.path.join(BASE_DIR, "Arquivos")
CAMINHO_EMPRESAS = os.path.join(BASE_DIR, "empresas.txt")

# Site
URL_LOGIN = "https://aefhondabr.nx-services.com/hondabr/index.html#/login"

# Arquivo gerado pelo "2 - Fromatador XLSX.py"
PADRAO_ARQUIVO_FINAL = "final_{empresa}.xlsx"

NOME_LOG = "Log - Organizar no Site.txt"

DELAY_ENTRE_ACOES = 0.2
PAUSA_ENTRE_EMPRESAS = 0.5

# Perfis de login (credenciais em ENV / .env)
PERFIS_LOGIN = ["LOBO", "CASTRO", "TELEMACO", "MOTOACAO", "RIO BRANCO"]
PERFIL_PADRAO = "RIO BRANCO"
DOTENV_PADROES = [
    Path(BASE_DIR) / ".env",
    Path(__file__).resolve().parent / ".env",
]

# Playwright
PLAYWRIGHT_HEADLESS = False
PLAYWRIGHT_TIMEOUT_MS = 30_000
PAUSAR_APOS_LOGIN = True

# Retries (pagina de submissions as vezes demora a "popular" a tabela)
TENTATIVAS_CLICAR_EDITAR = 3
ESPERA_ENTRE_TENTATIVAS_S = 5

# Edicao (apos clicar em Editar)
TEXTO_MENU_VEICULOS_NOVOS = "Veiculos Novos"
OPCAO_VEICULOS_NOVOS = "Ativo"
TENTATIVAS_CONFIRMAR = 3
ESPERA_ENTRE_TENTATIVAS_CONFIRMAR_S = 3

# XLSX (valores calculados)
ABA_ATIVO = "Ativo"
COL_NIVEL = 1
COL_DESCRICAO = 2
COL_VALOR = 3


# =========================
# Utilitarios
# =========================


def _normalizar_empresa(texto: str) -> str:
    return texto.strip().lstrip("\ufeff")


def carregar_empresas(caminho: str) -> list[str]:
    if not os.path.isfile(caminho):
        print(f"ERRO: arquivo de empresas nao encontrado: {caminho}")
        sys.exit(1)

    with open(caminho, "r", encoding="utf-8") as arquivo:
        empresas = [_normalizar_empresa(linha) for linha in arquivo.readlines() if linha.strip()]

    if not empresas:
        print("ERRO: lista de empresas vazia.")
        sys.exit(1)

    return empresas


def localizar_arquivo_final(empresa: str) -> str | None:
    alvo = PADRAO_ARQUIVO_FINAL.format(empresa=empresa)
    caminho = os.path.join(PASTA_ARQUIVOS, empresa, alvo)
    if os.path.isfile(caminho):
        return caminho

    # Fallback: busca em toda a pasta Arquivos (caso o arquivo nao esteja dentro da subpasta da empresa)
    alvo_lower = alvo.lower()
    for raiz, _, arquivos in os.walk(PASTA_ARQUIVOS):
        for nome in arquivos:
            if nome.lower() == alvo_lower:
                return os.path.join(raiz, nome)

    return None


def caminho_log() -> str:
    return os.path.join(BASE_DIR, NOME_LOG)


def log_linha(msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{ts}] {msg}"
    print(linha)
    with open(caminho_log(), "a", encoding="utf-8") as arq:
        arq.write(linha + "\n")


def montar_tarefas(empresas: list[str]) -> list[tuple[str, str]]:
    tarefas: list[tuple[str, str]] = []
    for emp in empresas:
        caminho = localizar_arquivo_final(emp)
        if not caminho:
            log_linha(f"AVISO: nao encontrei arquivo final para a empresa {emp}.")
            continue
        tarefas.append((emp, caminho))
    return tarefas


# =========================
# Automacao (site)
# =========================


def _normalizar_perfil(perfil: str) -> str:
    # ENV nao aceita espacos: "RIO BRANCO" -> "RIO_BRANCO"
    return perfil.strip().upper().replace("-", "_").replace(" ", "_")


def _carregar_dotenv_se_existir() -> None:
    """
    Carrega variaveis de um .env se existir (sem sobrescrever ENV do Windows).
    Se python-dotenv nao estiver instalado, apenas ignora.
    """
    try:
        from dotenv import load_dotenv  # type: ignore
    except Exception:
        return

    for p in DOTENV_PADROES:
        try:
            if p.is_file():
                load_dotenv(dotenv_path=p, override=False)
        except Exception:
            continue


def obter_credenciais_por_perfil(perfil: str) -> tuple[str, str]:
    """
    Procura credenciais do perfil em ENV (e opcionalmente .env).

    Padroes aceitos (ex.: perfil "RIO BRANCO" => "RIO_BRANCO"):
    - RIO_BRANCO_USUARIO / RIO_BRANCO_SENHA
    - RIO_BRANCO_USER / RIO_BRANCO_PASS
    - AEF_RIO_BRANCO_USUARIO / AEF_RIO_BRANCO_SENHA
    - AEF_RIO_BRANCO_USER / AEF_RIO_BRANCO_PASS
    - AEF_SITE_RIO_BRANCO_USUARIO / AEF_SITE_RIO_BRANCO_SENHA
    - AEF_SITE_RIO_BRANCO_USER / AEF_SITE_RIO_BRANCO_PASS
    """
    _carregar_dotenv_se_existir()

    pfx = _normalizar_perfil(perfil)

    candidatos = [
        (f"{pfx}_USUARIO", f"{pfx}_SENHA"),
        (f"{pfx}_USER", f"{pfx}_PASS"),
        (f"AEF_{pfx}_USUARIO", f"AEF_{pfx}_SENHA"),
        (f"AEF_{pfx}_USER", f"AEF_{pfx}_PASS"),
        (f"AEF_SITE_{pfx}_USUARIO", f"AEF_SITE_{pfx}_SENHA"),
        (f"AEF_SITE_{pfx}_USER", f"AEF_SITE_{pfx}_PASS"),
    ]

    for k_user, k_pass in candidatos:
        user = os.getenv(k_user, "").strip()
        senha = os.getenv(k_pass, "").strip()
        if user and senha:
            return user, senha

    chaves = []
    for k_user, k_pass in candidatos:
        chaves.append(f"{k_user} / {k_pass}")
    raise RuntimeError(
        "Credenciais nao encontradas no ENV/.env para o perfil "
        f"'{perfil}'. Chaves aceitas: " + "; ".join(chaves)
    )


def _importar_dependencias_ui():
    # Mantido por compatibilidade: este script esta migrando para Playwright.
    return


def _importar_dependencias_playwright():
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError  # noqa: F401
    except Exception as exc:
        print("ERRO: playwright nao encontrado.")
        print("Instale: pip install playwright")
        print("E depois: playwright install chromium")
        print(f"Detalhe: {exc}")
        raise


def _importar_dependencias_openpyxl():
    try:
        import openpyxl  # noqa: F401
    except Exception as exc:
        print("ERRO: openpyxl nao encontrado.")
        print("Instale: pip install openpyxl")
        print(f"Detalhe: {exc}")
        raise


def _salvar_screenshot(page, nome_base: str) -> str | None:
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome = f"{nome_base}_{ts}.png"
        caminho = os.path.join(BASE_DIR, nome)
        page.screenshot(path=caminho, full_page=True)
        return caminho
    except Exception:
        return None


def _e_nivel_l(texto: str) -> bool:
    t = (texto or "").strip().upper()
    if not t.startswith("L"):
        return False
    return t[1:].isdigit()


def _numero_nivel(texto: str) -> int:
    t = (texto or "").strip().upper()
    if not t.startswith("L"):
        return 999999
    try:
        return int(t[1:])
    except Exception:
        return 999999


def ler_linhas_ativo_xlsx(caminho_xlsx: str) -> list[dict]:
    """
    Le a aba "Ativo" do XLSX usando data_only=True para pegar o VALOR calculado.
    Retorna lista de itens: {nivel, descricao, valor}.
    """
    _importar_dependencias_openpyxl()
    import openpyxl

    if not os.path.isfile(caminho_xlsx):
        raise RuntimeError(f"Arquivo XLSX nao encontrado: {caminho_xlsx}")

    wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)
    if ABA_ATIVO in wb.sheetnames:
        ws = wb[ABA_ATIVO]
    else:
        ws = wb[wb.sheetnames[0]]

    itens: list[dict] = []
    for r in range(1, ws.max_row + 1):
        nivel = ws.cell(r, COL_NIVEL).value
        if nivel is None:
            continue
        nivel_txt = str(nivel).strip()
        if not _e_nivel_l(nivel_txt):
            continue

        descricao = ws.cell(r, COL_DESCRICAO).value
        valor = ws.cell(r, COL_VALOR).value

        if valor is None:
            # Sem valor calculado (pode acontecer se o arquivo nao foi salvo com calculo no Excel).
            continue
        if not isinstance(valor, (int, float)):
            continue

        itens.append(
            {
                "nivel": nivel_txt.upper(),
                "descricao": (str(descricao).strip() if descricao is not None else ""),
                "valor": float(valor),
            }
        )

    itens.sort(key=lambda x: _numero_nivel(x["nivel"]))
    return itens


def _formatar_valor_para_site(valor: float, usa_virgula: bool) -> str:
    # Mantem 2 casas e sem separador de milhar.
    s = f"{valor:.2f}"
    if usa_virgula:
        s = s.replace(".", ",")
    return s


def _detectar_decimal_virgula(frame) -> bool:
    """
    Tenta inferir se o input na tela usa virgula como decimal.
    Se nao der para ler, default False (ponto).
    """
    try:
        inp = frame.locator("input[type='text'], input[type='number']").first
        if inp.count() == 0:
            return False
        v = inp.input_value(timeout=1000) or ""
        return "," in v and "." not in v
    except Exception:
        return False


def _competencia_mes_anterior(hoje: datetime | None = None) -> str:
    if hoje is None:
        hoje = datetime.now()
    # Mes anterior com ajuste de ano.
    if hoje.month == 1:
        ano = hoje.year - 1
        mes = 12
    else:
        ano = hoje.year
        mes = hoje.month - 1

    meses = {
        1: "jan",
        2: "fev",
        3: "mar",
        4: "abr",
        5: "mai",
        6: "jun",
        7: "jul",
        8: "ago",
        9: "set",
        10: "out",
        11: "nov",
        12: "dez",
    }
    return f"{meses[mes]} {ano}"


def _aguardar_carregar_pagina(page) -> None:
    try:
        page.wait_for_load_state("domcontentloaded", timeout=PLAYWRIGHT_TIMEOUT_MS)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=PLAYWRIGHT_TIMEOUT_MS)
    except Exception:
        pass


def _esperar_frame_submissions(page, timeout_ms: int) -> object:
    """
    A tela de submissions pode demorar a popular a tabela e, em alguns casos,
    o conteudo pode estar em um iframe. Aqui fazemos polling em todos os frames
    ate encontrar um elemento tipico da tabela.
    """
    inicio = time.monotonic()
    timeout_s = max(1, int(timeout_ms / 1000))

    while time.monotonic() - inicio < timeout_s:
        candidatos = []
        try:
            candidatos.append(page.main_frame)
        except Exception:
            pass
        try:
            for fr in page.frames:
                if fr not in candidatos:
                    candidatos.append(fr)
        except Exception:
            pass

        for fr in candidatos:
            try:
                el = fr.locator("a:has-text('Editar')").first
                if el.count() > 0:
                    el.wait_for(state="visible", timeout=500)
                    return fr
            except Exception:
                pass
            try:
                el = fr.locator("text=Periodo").first
                if el.count() > 0:
                    el.wait_for(state="visible", timeout=500)
                    return fr
            except Exception:
                pass
            try:
                el = fr.locator("tr.normal").first
                if el.count() > 0:
                    el.wait_for(state="visible", timeout=500)
                    return fr
            except Exception:
                pass

        time.sleep(0.5)

    return page


def _ir_para_submissions(page) -> None:
    url = "https://aefhondabr.nx-services.com/hondabr/index.html#/submissions"
    log_linha(f"Acessando: {url}")
    page.goto(url)
    _aguardar_carregar_pagina(page)

    frame = _esperar_frame_submissions(page, timeout_ms=PLAYWRIGHT_TIMEOUT_MS)

    # A tabela dessa tela pode variar; ja vimos pelo menos 2 layouts:
    # - layout 1: <tr class="normal"> ... <a class="editForm">Editar</a>
    # - layout 2: header "Periodo" e link "Editar" sem classe.
    try:
        page.wait_for_function(
            "() => window.location && window.location.hash && window.location.hash.toLowerCase().includes('submissions')",
            timeout=10_000,
        )
    except Exception:
        pass

    try:
        frame.locator("text=Periodo").first.wait_for(state="visible", timeout=PLAYWRIGHT_TIMEOUT_MS)
        return frame
    except Exception:
        pass

    try:
        frame.locator("a:has-text('Editar')").first.wait_for(state="visible", timeout=PLAYWRIGHT_TIMEOUT_MS)
        return frame
    except Exception:
        pass

    try:
        frame.locator("tr.normal").first.wait_for(state="visible", timeout=PLAYWRIGHT_TIMEOUT_MS)
        return frame
    except Exception:
        caminho = _salvar_screenshot(page, "submissions_nao_carregou")
        raise RuntimeError(
            "Nao consegui carregar a tabela de submissions."
            + (f" Screenshot: {caminho}" if caminho else "")
        )


def _resolver_frame_conteudo_submissions(page):
    """
    Em algumas execucoes, o conteudo da tabela aparece em um iframe.
    Esta funcao tenta achar o frame onde os seletores da tabela existem.
    """
    frames = []
    try:
        frames = page.frames
    except Exception:
        return page

    candidatos = []
    try:
        candidatos.append(page.main_frame)
    except Exception:
        pass
    for fr in frames:
        if fr not in candidatos:
            candidatos.append(fr)

    # Tenta identificar o frame pela presenca de elementos tipicos da tela.
    for fr in candidatos:
        try:
            if fr.locator("text=Periodo").count() > 0:
                return fr
        except Exception:
            pass
        try:
            if fr.locator("a:has-text('Editar')").count() > 0:
                return fr
        except Exception:
            pass
        try:
            if fr.locator("tr.normal").count() > 0:
                return fr
        except Exception:
            pass

    return page


def _clicar_editar_na_competencia(frame, page, competencia: str) -> None:
    # Ex.: "jan 2026" (do HTML que voce colou).
    competencia_norm = competencia.strip().lower()

    # Tenta localizar a linha que contenha a competencia.
    # Layout A (HTML que voce colou): <tr class="normal"> ... <a>jan 2026</a> ... <a class="editForm">Editar</a>
    # Layout B (seu print): competencia aparece como texto na 1a coluna.
    td_primeira_coluna = frame.locator("td:nth-child(1)").filter(
        has_text=re.compile(rf"^{re.escape(competencia_norm)}$", re.I)
    )
    linha = frame.locator("tr").filter(has=td_primeira_coluna).first

    # Fallbacks
    if linha.count() == 0:
        link_comp = frame.locator("a").filter(has_text=re.compile(rf"^{re.escape(competencia_norm)}$", re.I)).first
        if link_comp.count() > 0:
            linha = frame.locator("tr", has=link_comp).first
        else:
            linha = frame.locator("tr").filter(
                has_text=re.compile(rf"\\b{re.escape(competencia_norm)}\\b", re.I)
            ).first

    try:
        linha.wait_for(state="visible", timeout=10_000)
    except Exception:
        caminho = _salvar_screenshot(page, "competencia_nao_encontrada")
        raise RuntimeError(
            f"Competencia nao encontrada na tabela: '{competencia}'."
            + (f" Screenshot: {caminho}" if caminho else "")
        )

    # "Editar" na mesma linha:
    # - layout 1: <a class="editForm">Editar</a>
    # - layout 2: link "Editar" sem classe
    editar = linha.locator("a.editForm").first
    if editar.count() == 0:
        editar = linha.locator("a").filter(has_text=re.compile(r"^Editar$", re.I)).first
    try:
        editar.wait_for(state="visible", timeout=10_000)
    except Exception:
        caminho = _salvar_screenshot(page, "botao_editar_nao_encontrado")
        raise RuntimeError(
            "Nao encontrei o botao 'Editar' na linha da competencia."
            + (f" Screenshot: {caminho}" if caminho else "")
        )

    log_linha(f"Clicando em Editar da competencia: {competencia_norm}")
    editar.click()
    _aguardar_carregar_pagina(page)

    caminho_ok = _salvar_screenshot(page, "apos_clicar_editar")
    if caminho_ok:
        log_linha(f"Screenshot: {caminho_ok}")


def _clicar_editar_na_competencia_com_retry(frame, page, competencia: str) -> None:
    ultima_exc: Exception | None = None
    for tentativa in range(1, TENTATIVAS_CLICAR_EDITAR + 1):
        try:
            _clicar_editar_na_competencia(frame, page, competencia=competencia)
            return
        except Exception as exc:
            ultima_exc = exc
            log_linha(f"AVISO: tentativa {tentativa}/{TENTATIVAS_CLICAR_EDITAR} falhou: {exc}")
            if tentativa < TENTATIVAS_CLICAR_EDITAR:
                time.sleep(ESPERA_ENTRE_TENTATIVAS_S)
                _aguardar_carregar_pagina(page)
                continue
            raise ultima_exc


def _esperar_frame_edicao(page, timeout_ms: int):
    """
    Apos clicar em Editar, a tela pode carregar em outro frame/contexto.
    Polling ate encontrar:
    - texto "Veiculos Novos" (com ou sem acento)
    - botao/submit "Confirmar"
    """
    inicio = time.monotonic()
    timeout_s = max(1, int(timeout_ms / 1000))

    padrao_veiculos = re.compile(r"Ve[ií]culos\\s+Novos", re.I)

    while time.monotonic() - inicio < timeout_s:
        candidatos = []
        try:
            candidatos.append(page.main_frame)
        except Exception:
            pass
        try:
            for fr in page.frames:
                if fr not in candidatos:
                    candidatos.append(fr)
        except Exception:
            pass

        for fr in candidatos:
            try:
                if fr.locator("input[type='submit'][value='Confirmar'], input[value='Confirmar']").count() > 0:
                    return fr
            except Exception:
                pass
            try:
                if fr.locator("text=/Ve[ií]culos\\s+Novos/i").count() > 0:
                    return fr
            except Exception:
                pass
            try:
                if fr.locator("text=" + TEXTO_MENU_VEICULOS_NOVOS).count() > 0:
                    return fr
            except Exception:
                pass

        time.sleep(0.5)

    return page


def _selecionar_veiculos_novos_ativo_e_confirmar(page) -> None:
    _aguardar_carregar_pagina(page)
    frame = _esperar_frame_edicao(page, timeout_ms=PLAYWRIGHT_TIMEOUT_MS)

    # Localiza o dropdown associado ao texto "Veiculos Novos" (com ou sem acento).
    alvo_txt = frame.locator("text=/Ve[ií]culos\\s+Novos/i").first
    if alvo_txt.count() == 0:
        alvo_txt = frame.locator(f"text={TEXTO_MENU_VEICULOS_NOVOS}").first

    if alvo_txt.count() == 0:
        caminho = _salvar_screenshot(page, "edicao_veiculos_novos_nao_encontrado")
        raise RuntimeError(
            "Nao encontrei o texto/menu 'Veiculos Novos' na tela de edicao."
            + (f" Screenshot: {caminho}" if caminho else "")
        )

    select = None
    try:
        # Tenta pegar um <select> na mesma linha/tabela.
        select = alvo_txt.locator("xpath=ancestor::tr[1]//select[1]").first
    except Exception:
        select = None

    if not select or select.count() == 0:
        try:
            select = alvo_txt.locator("xpath=ancestor::td[1]//select[1]").first
        except Exception:
            select = None

    if not select or select.count() == 0:
        try:
            select = alvo_txt.locator("xpath=following::select[1]").first
        except Exception:
            select = None

    if not select or select.count() == 0:
        caminho = _salvar_screenshot(page, "edicao_select_nao_encontrado")
        raise RuntimeError(
            "Nao encontrei o dropdown (<select>) para 'Veiculos Novos'."
            + (f" Screenshot: {caminho}" if caminho else "")
        )

    # Seleciona "Ativo"
    try:
        select.wait_for(state="visible", timeout=10_000)
    except Exception:
        pass

    log_linha(f"Selecionando dropdown '{TEXTO_MENU_VEICULOS_NOVOS}' -> '{OPCAO_VEICULOS_NOVOS}'")
    try:
        select.select_option(label=OPCAO_VEICULOS_NOVOS)
    except Exception:
        # Fallback: tenta por value.
        select.select_option(value=OPCAO_VEICULOS_NOVOS)

    # Clica Confirmar
    confirmar = frame.locator(
        "input#main\\:j_id86, input[type='submit'][value='Confirmar'], input[value='Confirmar']"
    ).first
    if confirmar.count() == 0:
        caminho = _salvar_screenshot(page, "edicao_confirmar_nao_encontrado")
        raise RuntimeError(
            "Nao encontrei o botao 'Confirmar'."
            + (f" Screenshot: {caminho}" if caminho else "")
        )

    for tentativa in range(1, TENTATIVAS_CONFIRMAR + 1):
        try:
            confirmar.wait_for(state="visible", timeout=10_000)
            log_linha("Clicando em Confirmar.")
            confirmar.click()
            _aguardar_carregar_pagina(page)
            caminho_ok = _salvar_screenshot(page, "apos_confirmar")
            if caminho_ok:
                log_linha(f"Screenshot: {caminho_ok}")
            return
        except Exception as exc:
            log_linha(f"AVISO: falha ao clicar Confirmar (tentativa {tentativa}/{TENTATIVAS_CONFIRMAR}): {exc}")
            if tentativa < TENTATIVAS_CONFIRMAR:
                time.sleep(ESPERA_ENTRE_TENTATIVAS_CONFIRMAR_S)
                _aguardar_carregar_pagina(page)
                continue
            raise


def _preencher_ativo_no_site(page, caminho_xlsx: str) -> None:
    """
    Preenche os valores do Ativo no site, batendo por codigo L1/L2/... e preenchendo apenas o valor.
    """
    linhas = ler_linhas_ativo_xlsx(caminho_xlsx)
    if not linhas:
        raise RuntimeError(
            "Nao encontrei linhas L* com valor calculado na aba Ativo do XLSX. "
            "Se o arquivo tiver formulas, confirme se foi salvo com os valores calculados."
        )

    _aguardar_carregar_pagina(page)
    frame = _esperar_frame_edicao(page, timeout_ms=PLAYWRIGHT_TIMEOUT_MS)
    usa_virgula = _detectar_decimal_virgula(frame)

    log_linha(f"Preenchendo Ativo (linhas: {len(linhas)}; decimal_virgula={usa_virgula}).")

    faltando: list[str] = []
    preenchidas = 0

    for item in linhas:
        nivel = item["nivel"]
        valor = item["valor"]

        # Procura a linha pelo texto do nivel.
        # Preferencia: td com texto exato "Lx"
        td_nivel = frame.locator("td").filter(has_text=re.compile(rf"^{re.escape(nivel)}$", re.I))
        row = frame.locator("tr").filter(has=td_nivel).first
        if row.count() == 0:
            # Fallback: qualquer linha com o texto Lx como palavra.
            row = frame.locator("tr").filter(has_text=re.compile(rf"\\b{re.escape(nivel)}\\b", re.I)).first

        if row.count() == 0:
            faltando.append(nivel)
            continue

        # Primeiro input editavel na linha
        inp = row.locator("input[type='text'], input[type='number']").first
        if inp.count() == 0:
            faltando.append(nivel)
            continue

        txt_val = _formatar_valor_para_site(valor, usa_virgula=usa_virgula)
        try:
            inp.scroll_into_view_if_needed(timeout=2000)
        except Exception:
            pass
        try:
            inp.fill(txt_val)
            preenchidas += 1
        except Exception:
            # Fallback: click + type (alguns campos bloqueiam fill)
            try:
                inp.click()
                inp.press("Control+A")
                inp.type(txt_val, delay=20)
                preenchidas += 1
            except Exception:
                faltando.append(nivel)

    log_linha(f"Ativo preenchido: {preenchidas}/{len(linhas)}.")
    if faltando:
        log_linha(f"AVISO: niveis nao preenchidos (nao encontrados/sem input): {', '.join(faltando[:30])}")
        if len(faltando) > 30:
            log_linha(f"AVISO: (+{len(faltando) - 30} niveis omitidos do log)")

    caminho_ok = _salvar_screenshot(page, "apos_preencher_ativo")
    if caminho_ok:
        log_linha(f"Screenshot: {caminho_ok}")


def executar_fluxo_playwright(
    url_login: str,
    usuario: str,
    senha: str,
    somente_login: bool,
    ate_editar: bool,
    ate_confirmar: bool,
    preencher_ativo: bool,
    caminho_xlsx_ativo: str,
    competencia: str,
    fechar_apos: bool,
    pausar: bool,
) -> None:
    _importar_dependencias_playwright()
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=PLAYWRIGHT_HEADLESS)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(PLAYWRIGHT_TIMEOUT_MS)

        log_linha(f"Abrindo: {url_login}")
        page.goto(url_login)
        _aguardar_carregar_pagina(page)

        # Campos do print: placeholders "Username" e "Password"
        user = page.locator(
            "input[placeholder='Username'], input[placeholder='User'], input[name*='user' i], input[id*='user' i]"
        ).first
        pwd = page.locator(
            "input[placeholder='Password'], input[type='password'], input[name*='pass' i], input[id*='pass' i]"
        ).first

        try:
            user.wait_for(state="visible", timeout=20_000)
            pwd.wait_for(state="visible", timeout=20_000)
        except PWTimeoutError:
            caminho = _salvar_screenshot(page, "login_campos_nao_encontrados")
            raise RuntimeError(
                "Nao encontrei os campos de login (Username/Password)."
                + (f" Screenshot: {caminho}" if caminho else "")
            )

        user.fill(usuario)
        pwd.fill(senha)

        btn = page.get_by_role("button", name=re.compile(r"^login$", re.I))
        if btn.count() == 0:
            btn = page.locator("button:has-text('LOGIN'), button:has-text('Login'), input[type='submit']").first
            btn.click()
        else:
            btn.first.click()

        # SPA (hash). Em geral sai de "#/login" para outra rota.
        try:
            page.wait_for_function(
                "() => window.location && window.location.hash && !window.location.hash.toLowerCase().includes('login')",
                timeout=20_000,
            )
        except PWTimeoutError:
            # Fallback: se os campos ainda estiverem visiveis, assume falha
            try:
                if user.is_visible() or pwd.is_visible():
                    caminho = _salvar_screenshot(page, "login_falhou")
                    raise RuntimeError(
                        "Login nao saiu da tela de login (possivel usuario/senha invalidos ou captcha/politica)."
                        + (f" Screenshot: {caminho}" if caminho else "")
                    )
            except Exception:
                pass

        _aguardar_carregar_pagina(page)

        log_linha("Login executado. Verifique se entrou no sistema.")
        caminho_ok = _salvar_screenshot(page, "login_ok")
        if caminho_ok:
            log_linha(f"Screenshot: {caminho_ok}")

        if not somente_login:
            frame = _ir_para_submissions(page)
            if ate_editar:
                _clicar_editar_na_competencia_com_retry(frame, page, competencia=competencia)
                if ate_confirmar:
                    _selecionar_veiculos_novos_ativo_e_confirmar(page)
                    if preencher_ativo:
                        _preencher_ativo_no_site(page, caminho_xlsx=caminho_xlsx_ativo)

        if fechar_apos:
            context.close()
            browser.close()
            return

        # Mantem aberto (processo assistido).
        if sys.stdin.isatty():
            if pausar:
                input("Pressione ENTER para fechar o navegador...")
                context.close()
                browser.close()
            else:
                print("Navegador mantido aberto. Para encerrar, feche a janela ou use Ctrl+C no terminal.")
                while True:
                    time.sleep(1)
        else:
            while True:
                time.sleep(1)


def organizar_no_site(empresa: str, caminho_arquivo: str) -> None:
    """
    TODO: implementar o fluxo do site.

    Para completar, preciso que voce confirme:
    - qual e o site (URL / sistema)
    - qual e o passo a passo (login, menu, tela, campos)
    - se usa Chrome/Edge e se ha extensoes/popup
    - onde anexar o arquivo (upload) e como confirmar sucesso
    """
    _importar_dependencias_ui()
    raise NotImplementedError("Fluxo do site ainda nao implementado (apos login).")


# =========================
# Main
# =========================


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Organiza/Envia arquivos finais no site (assistido).")
    p.add_argument(
        "--empresa",
        help="Processa apenas uma empresa (ex.: 22). Se omitido, usa empresas.txt.",
        default="",
    )
    p.add_argument(
        "--perfil",
        default=PERFIL_PADRAO,
        help="Perfil de login (LOBO/CASTRO/TELEMACO/MOTOACAO/RIO BRANCO).",
    )
    p.add_argument(
        "--url",
        default="",
        help="URL do login (sobrescreve URL_LOGIN).",
    )
    p.add_argument(
        "--somente-login",
        action="store_true",
        help="Apenas abre o site e tenta logar (nao processa empresas/arquivos).",
    )
    p.add_argument(
        "--ate-editar",
        action="store_true",
        help="Apos logar, vai em /submissions e clica em Editar na competencia do mes anterior (ou --competencia).",
    )
    p.add_argument(
        "--ate-confirmar",
        action="store_true",
        help="Apos clicar em Editar, seleciona 'Veiculos Novos' = Ativo e clica em Confirmar.",
    )
    p.add_argument(
        "--preencher-ativo",
        action="store_true",
        help="Apos Confirmar, preenche os valores do Ativo no site a partir do final_{empresa}.xlsx.",
    )
    p.add_argument(
        "--competencia",
        default="",
        help="Override da competencia (ex.: \"jan 2026\"). Se vazio, usa mes anterior.",
    )
    p.add_argument(
        "--fechar-apos-login",
        action="store_true",
        help="Fecha o navegador automaticamente apos o login.",
    )
    p.add_argument(
        "--nao-pausar-apos-login",
        action="store_true",
        help="Nao pausa aguardando ENTER apos logar.",
    )
    p.add_argument(
        "--somente-listar",
        action="store_true",
        help="Apenas lista as empresas/arquivos encontrados (nao executa automacao).",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()

    if not args.somente_listar:
        perfis_norm = {_normalizar_perfil(p): p for p in PERFIS_LOGIN}
        perfil_in = args.perfil.strip()
        perfil_norm = _normalizar_perfil(perfil_in)
        if perfil_norm not in perfis_norm:
            log_linha(f"ERRO: perfil invalido: {args.perfil}. Perfis: {', '.join(PERFIS_LOGIN)}")
            return 1
        perfil_ok = perfis_norm[perfil_norm]

        url = (args.url or URL_LOGIN).strip()
        if not url:
            log_linha("ERRO: informe --url (link do login) ou preencha URL_LOGIN no topo do script.")
            return 1

        try:
            usuario, senha = obter_credenciais_por_perfil(perfil_ok)
            log_linha(f"Perfil de login: {perfil_ok} (usuario: {usuario}).")
        except Exception as exc:
            log_linha(f"ERRO: {exc}")
            return 1

        if args.somente_login:
            try:
                competencia = args.competencia.strip() or _competencia_mes_anterior()
                # Default: manter aberto. Para fechar automatico, use --fechar-apos-login.
                pausar = (not args.fechar_apos_login) and PAUSAR_APOS_LOGIN and (not args.nao_pausar_apos_login)
                executar_fluxo_playwright(
                    url_login=url,
                    usuario=usuario,
                    senha=senha,
                    somente_login=True,
                    ate_editar=False,
                    ate_confirmar=False,
                    preencher_ativo=False,
                    caminho_xlsx_ativo="",
                    competencia=competencia,
                    fechar_apos=args.fechar_apos_login,
                    pausar=pausar,
                )
                return 0
            except Exception as exc:
                log_linha(f"ERRO: {exc}")
                return 2

        if args.ate_editar:
            try:
                competencia = args.competencia.strip() or _competencia_mes_anterior()
                log_linha(f"Competencia alvo: {competencia}")
                pausar = (not args.fechar_apos_login) and PAUSAR_APOS_LOGIN and (not args.nao_pausar_apos_login)

                caminho_xlsx = ""
                if args.preencher_ativo:
                    if not args.empresa:
                        log_linha("ERRO: para --preencher-ativo, informe --empresa (ex.: --empresa 22).")
                        return 1
                    caminho_xlsx = localizar_arquivo_final(_normalizar_empresa(args.empresa)) or ""
                    if not caminho_xlsx:
                        log_linha(f"ERRO: nao encontrei o arquivo final da empresa {args.empresa}.")
                        return 1
                    log_linha(f"XLSX Ativo: {caminho_xlsx}")

                executar_fluxo_playwright(
                    url_login=url,
                    usuario=usuario,
                    senha=senha,
                    somente_login=False,
                    ate_editar=True,
                    ate_confirmar=bool(args.ate_confirmar),
                    preencher_ativo=bool(args.preencher_ativo),
                    caminho_xlsx_ativo=caminho_xlsx,
                    competencia=competencia,
                    fechar_apos=args.fechar_apos_login,
                    pausar=pausar,
                )
                return 0
            except Exception as exc:
                log_linha(f"ERRO: {exc}")
                return 2

    if args.empresa:
        empresas = [_normalizar_empresa(args.empresa)]
    else:
        empresas = carregar_empresas(CAMINHO_EMPRESAS)

    tarefas = montar_tarefas(empresas)
    if not tarefas:
        log_linha("ERRO: nenhuma tarefa encontrada (nenhum arquivo final localizado).")
        return 1

    if args.somente_listar:
        log_linha("Modo somente-listar.")
        for emp, arq in tarefas:
            log_linha(f"OK: {emp} -> {arq}")
        return 0

    log_linha(f"Iniciando automacao (tarefas: {len(tarefas)}).")
    for i, (emp, arq) in enumerate(tarefas, start=1):
        log_linha(f"[{i}/{len(tarefas)}] Empresa {emp}: {arq}")
        try:
            organizar_no_site(emp, arq)
            log_linha(f"OK: empresa {emp}.")
        except NotImplementedError as exc:
            log_linha(f"ERRO: {exc}")
            return 2
        except Exception as exc:
            log_linha(f"ERRO: falha na empresa {emp}: {exc}")
            return 3

        time.sleep(PAUSA_ENTRE_EMPRESAS)

    log_linha("Finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
