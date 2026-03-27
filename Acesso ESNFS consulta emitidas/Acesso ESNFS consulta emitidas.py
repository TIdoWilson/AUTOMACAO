# -*- coding: utf-8 -*-
"""
Fluxo ESNFS:
1) Conecta no Chrome via CDP (porta 9222)
2) Abre tela inicial e aguarda login manual
3) Acessa tela de consulta
4) Altera Origem para "Recebida"
5) Seleciona Tomador configurado
6) Clica em "Pesquisar"
7) Percorre as paginas e baixa os PDFs das notas (lupa da 1a coluna)
"""

from __future__ import annotations

import datetime as dt
import os
import re
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path
from shutil import which

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# ===================== Configuracao =====================
PORTA_CDP = 9222
CDP_ENDPOINT = f"http://127.0.0.1:{PORTA_CDP}"

URL_LOGIN = "https://www.esnfs.com.br/"
URL_CONSULTA = "https://www.esnfs.com.br/nfsconsultanota.load.logic"

ORIGEM_VALUE_RECEBIDA = "2"

TOMADOR_TEXTO = "49426814000181 - JRC ADMINISTRADORA DE BENS LTDA (318442)"
# Se quiser travar por value do option:
# TOMADOR_OPTION_VALUE = "3328515"
TOMADOR_OPTION_VALUE = None

TIMEOUT_PADRAO_MS = 60_000
TIMEOUT_CURTO_MS = 6_000


# ===================== Helpers =====================
def limpar_nome_arquivo(texto: str) -> str:
    s = (texto or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:180] or "sem_nome"


def nome_arquivo_chave(chave: str) -> str:
    digitos = re.sub(r"\D", "", chave or "")
    if len(digitos) >= 8:
        return digitos
    return limpar_nome_arquivo(chave or "chave_sem_valor")


def caminho_pdf_unico(destino: Path) -> Path:
    if not destino.exists():
        return destino
    base = destino.stem
    ext = destino.suffix
    pasta = destino.parent
    i = 2
    while True:
        novo = pasta / f"{base}_{i}{ext}"
        if not novo.exists():
            return novo
        i += 1


# ===================== Chrome 9222 =====================
def encontrar_chrome_exe() -> str:
    candidatos = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        which("chrome"),
        which("chrome.exe"),
    ]
    for caminho in candidatos:
        if caminho and os.path.exists(caminho):
            return caminho
    raise FileNotFoundError("chrome.exe nao encontrado. Ajuste o caminho manualmente.")


def cdp_ativo(porta: int = PORTA_CDP) -> bool:
    url = f"http://127.0.0.1:{porta}/json/version"
    try:
        with urllib.request.urlopen(url, timeout=1) as resposta:
            return resposta.status == 200
    except Exception:
        return False


def iniciar_chrome_9222(porta: int = PORTA_CDP) -> subprocess.Popen | None:
    chrome = encontrar_chrome_exe()
    perfil = str(Path(tempfile.gettempdir()) / f"chrome-dev-{porta}")
    os.makedirs(perfil, exist_ok=True)

    args = [
        chrome,
        f"--remote-debugging-port={porta}",
        f"--user-data-dir={perfil}",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    for _ in range(60):
        if cdp_ativo(porta):
            return proc
        time.sleep(0.25)
    raise TimeoutError("Chrome com CDP nao respondeu a tempo.")


# ===================== Fluxo ESNFS =====================
def obter_pagina(browser):
    if browser.contexts and browser.contexts[0].pages:
        return browser.contexts[0].pages[0]
    contexto = browser.new_context()
    return contexto.new_page()


def aguardar_login_manual() -> None:
    input("Faca o login manual no ESNFS e pressione ENTER para continuar...")


def abrir_consulta_com_tentativas(browser, tentativas: int = 5):
    ultimo_erro = None
    for tentativa in range(1, tentativas + 1):
        page_consulta = browser.contexts[0].new_page()
        page_consulta.set_default_timeout(TIMEOUT_PADRAO_MS)
        try:
            page_consulta.goto(URL_CONSULTA, wait_until="domcontentloaded")
            page_consulta.locator("#parametrosTela\\.origemEmissaoNfse").first.wait_for(
                state="visible", timeout=15_000
            )
            return page_consulta
        except Exception as e:
            ultimo_erro = e
            url_atual = ""
            try:
                url_atual = page_consulta.url or ""
            except Exception:
                pass

            if "nfssecurity.check.logic" in url_atual.lower():
                input(
                    "ESNFS abriu a validacao de seguranca/certificado. "
                    "Conclua no navegador e pressione ENTER para tentar novamente..."
                )

            try:
                page_consulta.close()
            except Exception:
                pass

            if tentativa < tentativas:
                time.sleep(1.0)
                continue
            raise RuntimeError(f"Falha ao abrir tela de consulta apos {tentativas} tentativas: {ultimo_erro}")
    raise RuntimeError(f"Falha ao abrir tela de consulta: {ultimo_erro}")


def selecionar_origem_recebida(page) -> None:
    seletor_origem = "#parametrosTela\\.origemEmissaoNfse"
    page.locator(seletor_origem).wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    page.select_option(seletor_origem, value=ORIGEM_VALUE_RECEBIDA)


def _achar_select_tomador(page):
    seletores = [
        "#parametrosTela\\.idPessoa",
        "select[name='parametrosTela.idPessoa']",
        "select[id='parametrosTela.idPessoa']",
        "select[id*='idPessoa']",
        "select[name*='idPessoa']",
    ]
    for seletor in seletores:
        loc = page.locator(seletor)
        if loc.count() > 0:
            try:
                loc.first.wait_for(state="visible", timeout=5_000)
                return loc.first
            except PlaywrightTimeoutError:
                continue
    raise RuntimeError("Nao foi possivel localizar o campo de tomador.")


def selecionar_tomador(page, texto: str, option_value: str | None = None) -> None:
    select = _achar_select_tomador(page)

    if option_value:
        select.select_option(value=option_value)
        page.wait_for_load_state("domcontentloaded")
        return

    selecionou = select.evaluate(
        """(el, alvo) => {
            const normalizar = (s) => (s || "")
                .toUpperCase()
                .replace(/\\s+/g, " ")
                .trim();
            const alvoUpper = normalizar(alvo);
            const opcoes = Array.from(el.options || []);
            const op = opcoes.find(o => normalizar(o.textContent || "").includes(alvoUpper));
            if (!op) return false;
            el.value = op.value;
            el.dispatchEvent(new Event("change", { bubbles: true }));
            return true;
        }""",
        texto,
    )

    if not selecionou:
        raise RuntimeError(f"Tomador nao encontrado no select: {texto}")

    page.wait_for_load_state("domcontentloaded")
    select = _achar_select_tomador(page)

    texto_selecionado = select.evaluate(
        """(el) => {
            const op = el.options[el.selectedIndex];
            return op ? (op.textContent || "").trim() : "";
        }"""
    )
    texto_norm = " ".join((texto or "").upper().split())
    selecionado_norm = " ".join((texto_selecionado or "").upper().split())
    if texto_norm not in selecionado_norm:
        raise RuntimeError(f"Tomador selecionado nao confere: {texto_selecionado}")


def obter_nome_tomador_atual(page) -> str:
    select = _achar_select_tomador(page)
    texto_selecionado = select.evaluate(
        """(el) => {
            const op = el.options[el.selectedIndex];
            return op ? (op.textContent || "").trim() : "";
        }"""
    )
    return texto_selecionado or TOMADOR_TEXTO


def clicar_pesquisar(page) -> None:
    botao = page.locator("input.botaoPesquisar")
    botao.first.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    botao.first.click()
    page.wait_for_load_state("domcontentloaded")


def criar_pasta_saida(page) -> Path:
    base_dir = Path(__file__).resolve().parent
    tomador = limpar_nome_arquivo(obter_nome_tomador_atual(page))
    hoje = dt.datetime.now().strftime("%d-%m-%Y")
    pasta = base_dir / f"PDFs ESNFS - {tomador} - {hoje}"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def mapear_linhas_resultado(page) -> tuple[int, list[dict]]:
    dados = page.evaluate(
        """() => {
            const tabela = document.querySelector("#tabelaDinamica");
            if (!tabela) return { tableIndex: -1, linhas: [] };

            const normalizar = (s) => (s || "").replace(/\\s+/g, " ").trim();
            const rows = Array.from(tabela.querySelectorAll("tbody tr"));
            const linhas = [];

            rows.forEach((row, rowIndex) => {
                const tds = row.querySelectorAll("td");
                if (!tds || tds.length < 2) return;
                const linkLupa = tds[0].querySelector("a");
                const temLupa = Boolean(linkLupa && linkLupa.querySelector("i.fa.fa-search"));
                if (!temLupa) return;
                const numeroNfse = normalizar(tds[1].textContent || "");
                if (!numeroNfse) return;
                const href = linkLupa.getAttribute("href") || "";
                const m = href.match(/viewNota\\(["']([^"']+)["']\\)/i);
                const token = m ? m[1] : "";
                linhas.push({ rowIndex, chave: numeroNfse, token });
            });
            return { tableIndex: 0, linhas };
        }"""
    )
    return int(dados.get("tableIndex", -1)), list(dados.get("linhas", []))


def clicar_lupa_da_linha(page, table_index: int, row_index: int):
    if table_index != 0:
        raise RuntimeError("Indice de tabela invalido para #tabelaDinamica.")
    tabela = page.locator("#tabelaDinamica")
    row = tabela.locator("tbody tr").nth(row_index)
    celula = row.locator("td").first

    seletores = [
        "a:has(i.fa.fa-search)",
        "a i.fa.fa-search",
        "a",
    ]

    alvo = None
    for sel in seletores:
        loc = celula.locator(sel)
        if loc.count() > 0:
            alvo = loc.first
            break
    if not alvo:
        raise RuntimeError(f"Nao achei botao de lupa na linha {row_index}.")

    popup = None
    try:
        with page.expect_popup(timeout=TIMEOUT_CURTO_MS) as popup_info:
            alvo.click()
        popup = popup_info.value
    except PlaywrightTimeoutError:
        pass

    return popup


def _tentar_download_por_clique(page_alvo, destino: Path) -> bool:
    seletores = [
        "a[href*='.pdf']",
        "button:has-text('PDF')",
        "a:has-text('PDF')",
        "input[value*='PDF' i]",
        "button:has-text('Imprimir')",
        "a:has-text('Imprimir')",
        "input[value*='Imprimir' i]",
        "button:has-text('Visualizar')",
        "a:has-text('Visualizar')",
    ]
    for seletor in seletores:
        loc = page_alvo.locator(seletor)
        if loc.count() <= 0:
            continue
        try:
            with page_alvo.expect_download(timeout=TIMEOUT_CURTO_MS) as dlinfo:
                loc.first.click()
            download = dlinfo.value
            destino_final = caminho_pdf_unico(destino)
            download.save_as(str(destino_final))
            return True
        except PlaywrightTimeoutError:
            continue
        except Exception:
            continue
    return False


def _tentar_download_por_url_pdf(page_alvo, destino: Path) -> bool:
    urls = []
    try:
        url_atual = page_alvo.url or ""
        # No ESNFS o PDF abre inline em URL tipo ".../esenfs.view.logic?aut=..."
        if url_atual:
            urls.append(url_atual)
    except Exception:
        pass

    try:
        hrefs = page_alvo.evaluate(
            """() => Array.from(document.querySelectorAll("a[href]"))
                .map(a => a.href || "")
                .filter(h => /\\.pdf(\\?|$)/i.test(h));"""
        )
        for u in hrefs:
            if u and u not in urls:
                urls.append(u)
    except Exception:
        pass

    for url in urls:
        try:
            resp = page_alvo.context.request.get(
                url,
                timeout=TIMEOUT_PADRAO_MS,
                headers={"referer": (page_alvo.url or "")},
            )
            if not resp.ok:
                continue
            content_type = (resp.headers.get("content-type") or "").lower()
            conteudo = resp.body()
            if not conteudo:
                continue
            eh_pdf = ("application/pdf" in content_type) or conteudo.startswith(b"%PDF")
            if not eh_pdf:
                continue
            destino_final = caminho_pdf_unico(destino)
            destino_final.write_bytes(conteudo)
            return True
        except Exception:
            continue

    # Fallback forte: buscar o PDF dentro da propria aba com cookies/sessao do navegador.
    try:
        resultado = page_alvo.evaluate(
            """async () => {
                const candidatos = [];
                const pushUrl = (u) => {
                    if (!u) return;
                    try {
                        const abs = new URL(u, window.location.href).href;
                        if (!candidatos.includes(abs)) candidatos.push(abs);
                    } catch (_) {}
                };

                pushUrl(window.location.href);
                document.querySelectorAll("iframe[src], embed[src], object[data], a[href]").forEach((el) => {
                    pushUrl(el.getAttribute("src") || el.getAttribute("data") || el.getAttribute("href"));
                });

                const toBase64 = (bytes) => {
                    let binary = "";
                    const chunk = 0x8000;
                    for (let i = 0; i < bytes.length; i += chunk) {
                        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
                    }
                    return btoa(binary);
                };

                for (const url of candidatos) {
                    try {
                        const resp = await fetch(url, { credentials: "include" });
                        if (!resp.ok) continue;
                        const ct = (resp.headers.get("content-type") || "").toLowerCase();
                        const buf = await resp.arrayBuffer();
                        const bytes = new Uint8Array(buf);
                        const isPdf = ct.includes("application/pdf")
                            || (bytes.length >= 4
                                && bytes[0] === 0x25
                                && bytes[1] === 0x50
                                && bytes[2] === 0x44
                                && bytes[3] === 0x46);
                        if (!isPdf) continue;
                        return { ok: true, b64: toBase64(bytes), url, ct };
                    } catch (_) {}
                }
                return { ok: false };
            }"""
        )
        if resultado and resultado.get("ok") and resultado.get("b64"):
            import base64

            conteudo = base64.b64decode(resultado["b64"])
            destino_final = caminho_pdf_unico(destino)
            destino_final.write_bytes(conteudo)
            return True
    except Exception:
        pass

    return False


def salvar_pdf_da_nota(page_alvo, chave: str, pasta_saida: Path) -> Path:
    nome = f"{nome_arquivo_chave(chave)}.pdf"
    destino = pasta_saida / nome

    if _tentar_download_por_clique(page_alvo, destino):
        return destino
    if _tentar_download_por_url_pdf(page_alvo, destino):
        return destino

    raise RuntimeError("Nao foi possivel baixar o PDF automaticamente para esta nota.")


def baixar_pdf_por_token(contexto, token: str, chave: str, pasta_saida: Path, referer_url: str = "") -> Path | None:
    tk = (token or "").strip()
    if not tk:
        return None

    urls = [
        f"https://www.esnfs.com.br/esenfs.view.logic?aut={tk}",
        f"https://esnfs.com.br/esenfs.view.logic?aut={tk}",
    ]
    headers = {}
    if referer_url:
        headers["referer"] = referer_url

    destino_base = pasta_saida / f"{nome_arquivo_chave(chave)}.pdf"

    for url in urls:
        try:
            resp = contexto.request.get(url, timeout=TIMEOUT_PADRAO_MS, headers=headers)
            if not resp.ok:
                continue
            content_type = (resp.headers.get("content-type") or "").lower()
            conteudo = resp.body()
            if not conteudo:
                continue
            eh_pdf = ("application/pdf" in content_type) or conteudo.startswith(b"%PDF")
            if not eh_pdf:
                continue
            destino_final = caminho_pdf_unico(destino_base)
            destino_final.write_bytes(conteudo)
            return destino_final
        except Exception:
            continue
    return None


def abrir_lupas_da_pagina(page, browser, linhas: list[dict]) -> list[dict]:
    if not browser.contexts:
        return []
    contexto = browser.contexts[0]
    total_antes = len(contexto.pages)

    tokens = [str((i.get("token") or "")).strip() for i in linhas if str((i.get("token") or "")).strip()]
    tokens = list(dict.fromkeys(tokens))
    if not tokens:
        return []

    page.evaluate(
        """(tokens) => {
            for (const tk of tokens) {
                try {
                    if (typeof viewNota === "function") {
                        viewNota(tk);
                        continue;
                    }
                } catch (_) {}
                try {
                    const anchors = Array.from(document.querySelectorAll("#tabelaDinamica tbody tr td:first-child a[href]"));
                    const alvo = anchors.find(a => (a.getAttribute("href") || "").includes(tk));
                    if (alvo) alvo.click();
                } catch (_) {}
            }
        }""",
        tokens,
    )

    esperado = min(total_antes + len(tokens), total_antes + len(linhas))
    fim = time.time() + 10.0
    while time.time() < fim:
        if len(contexto.pages) >= esperado:
            break
        page.wait_for_timeout(100)

    abertas = []
    paginas = list(contexto.pages)
    for item in linhas:
        chave = (item.get("chave") or "").strip()
        token = (item.get("token") or "").strip()
        if not token:
            continue
        popup = None
        for p in paginas:
            if p == page:
                continue
            try:
                if token.lower() in (p.url or "").lower():
                    popup = p
                    break
            except Exception:
                continue
        if popup:
            abertas.append({"chave": chave, "page": popup, "token": token})
        else:
            print(f"[WARN] Aba nao localizada para NFS-e {chave} (token {token}).")
    return abertas


def fechar_abas_exceto_principal(browser, page_principal) -> None:
    if not browser.contexts:
        return
    contexto = browser.contexts[0]
    for aba in list(contexto.pages):
        if aba == page_principal:
            continue
        try:
            aba.close()
        except Exception:
            pass


def obter_pagina_atual_tabela(page) -> int:
    try:
        n = page.evaluate(
            """() => {
                const ativo = document.querySelector("#tabelaDinamica_paginate li.paginate_button.active a");
                if (!ativo) return 0;
                const txt = (ativo.textContent || "").trim();
                const num = parseInt(txt, 10);
                return Number.isFinite(num) ? num : 0;
            }"""
        )
        return int(n or 0)
    except Exception:
        return 0


def ir_para_proxima_pagina(page) -> bool:
    item_next = page.locator("li#tabelaDinamica_next")
    if item_next.count() <= 0:
        return False

    classe_item = (item_next.first.get_attribute("class") or "").lower()
    if "disabled" in classe_item:
        return False

    link = item_next.first.locator("a")
    if link.count() <= 0:
        return False
    link = link.first

    try:
        if not link.is_visible():
            return False
    except Exception:
        return False

    pagina_antes = obter_pagina_atual_tabela(page)
    try:
        # Clique unico para evitar pular pagina (1->3).
        link.click()
    except Exception:
        return False

    fim = time.time() + 12.0
    while time.time() < fim:
        pagina_depois = obter_pagina_atual_tabela(page)
        if pagina_depois and pagina_antes and pagina_depois > pagina_antes:
            return True
        if pagina_depois and not pagina_antes:
            return True
        page.wait_for_timeout(120)

    return False


def processar_resultados(browser, page, pasta_saida: Path) -> None:
    pagina = 1
    total_baixados = 0
    falhas = 0

    while True:
        pagina_ui = obter_pagina_atual_tabela(page)
        if pagina_ui > 0:
            pagina = pagina_ui
        table_index, linhas = mapear_linhas_resultado(page)
        qtd_lupas = len(linhas)
        print(f"[PAGINA {pagina}] Lupas encontradas: {qtd_lupas}")

        if table_index < 0 or qtd_lupas == 0:
            print(f"[PAGINA {pagina}] Nenhum item de resultado encontrado.")

        if qtd_lupas > 0 and table_index >= 0:
            print(f"[PAGINA {pagina}] Baixando por token (aut=...) sem depender de abrir abas...")
            contexto = browser.contexts[0] if browser.contexts else None
            if not contexto:
                raise RuntimeError("Contexto do navegador indisponivel.")

            for item in linhas:
                chave = (item.get("chave") or "").strip()
                token = (item.get("token") or "").strip()
                print(f"[PAGINA {pagina}] Processando NFS-e: {chave}")
                try:
                    caminho = baixar_pdf_por_token(
                        contexto=contexto,
                        token=token,
                        chave=chave,
                        pasta_saida=pasta_saida,
                        referer_url=page.url,
                    )
                    if caminho:
                        total_baixados += 1
                        print(f"[OK] PDF salvo para NFS-e {chave}")
                        continue

                    # Fallback: abre a nota e tenta salvar pela aba (se o token direto falhar)
                    popup = None
                    try:
                        row_index = int(item["rowIndex"])
                        popup = clicar_lupa_da_linha(page, table_index, row_index)
                        if popup:
                            popup.wait_for_timeout(1200)
                            salvar_pdf_da_nota(popup, chave, pasta_saida)
                            total_baixados += 1
                            print(f"[OK] PDF salvo para NFS-e {chave} (fallback aba)")
                        else:
                            raise RuntimeError("fallback sem popup")
                    finally:
                        if popup:
                            try:
                                popup.close()
                            except Exception:
                                pass
                except Exception as e:
                    falhas += 1
                    print(f"[ERRO] Falha na NFS-e {chave}: {e}")
            fechar_abas_exceto_principal(browser, page)

        if not ir_para_proxima_pagina(page):
            print("[OK] Nao ha proxima pagina. Processo completo.")
            print(f"[RESUMO] PDFs baixados: {total_baixados} | Falhas: {falhas}")
            return

        page.wait_for_timeout(700)


def main() -> None:
    chrome_proc = None
    with sync_playwright() as p:
        if not cdp_ativo(PORTA_CDP):
            chrome_proc = iniciar_chrome_9222(PORTA_CDP)

        browser = p.chromium.connect_over_cdp(CDP_ENDPOINT)
        page = obter_pagina(browser)
        page.set_default_timeout(TIMEOUT_PADRAO_MS)

        page.goto(URL_LOGIN, wait_until="domcontentloaded")
        aguardar_login_manual()

        page = abrir_consulta_com_tentativas(browser)
        selecionar_origem_recebida(page)
        selecionar_tomador(page, texto=TOMADOR_TEXTO, option_value=TOMADOR_OPTION_VALUE)
        clicar_pesquisar(page)

        pasta_saida = criar_pasta_saida(page)
        print(f"[INFO] Pasta de saida: {pasta_saida}")

        processar_resultados(browser, page, pasta_saida)

        print("[OK] Encerrando navegador.")
        try:
            browser.close()
        except Exception:
            pass

    if chrome_proc is not None:
        try:
            chrome_proc.terminate()
        except Exception:
            pass


if __name__ == "__main__":
    main()
