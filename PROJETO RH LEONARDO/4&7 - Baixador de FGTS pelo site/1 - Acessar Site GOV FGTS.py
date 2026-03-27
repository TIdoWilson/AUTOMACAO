
import csv
import json
import os
import re
import shutil
import subprocess
import time
import unicodedata
import winreg
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright
from chrome_9222 import PORT, chrome_9222

LOGIN_URL = "https://fgtsdigital.sistema.gov.br/portal/login"
SERVICOS_URL = "https://fgtsdigital.sistema.gov.br/portal/servicos"
EMISSAO_GUIA_RAPIDA_URL = "https://fgtsdigital.sistema.gov.br/cobranca/#/gestao-guias/emissao-guia-rapida"

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"

GRUPOS_PARTE_1 = {"6", "13", "14"}
GRUPOS_PARTE_2 = {"31", "32", "33", "34", "35", "36", "37"}

REGISTRO_BAIXADOS_PADRAO = BASE_DIR / "registro_downloads_fgts.csv"
REGISTRO_ERROS_PADRAO = BASE_DIR / "erros_download_fgts.csv"

AUTOMATIZADO_1A = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1ª PARTE")
AUTOMATIZADO_2A = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\2ª PARTE")


def load_env(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    if not path.exists():
        return data
    for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        data[k.strip()] = v.strip().strip('"').strip("'")
    return data


def merge_env_with_session(env: dict[str, str]) -> dict[str, str]:
    out = dict(env)
    for k, v in os.environ.items():
        if (k.startswith("FGTS_") or k.startswith("CERT_")) and str(v).strip():
            out[k] = str(v).strip()
    return out


def _apenas_digitos(v: str) -> str:
    return re.sub(r"\D", "", v or "")


def _sanitizar_nome(v: str) -> str:
    s = (v or "").strip()
    for ch in '<>:"/\\|?*':
        s = s.replace(ch, "-")
    s = " ".join(s.split())
    return s or "SEM_NOME"


def _normalizar_texto_busca(v: str) -> str:
    s = unicodedata.normalize("NFKD", v or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()


def _parse_grupos(v: str) -> list[str]:
    return [x.strip() for x in (v or "").split(";") if x.strip()]


def _parte_por_grupo(g: str) -> str:
    if g in GRUPOS_PARTE_1:
        return "1a_parte"
    if g in GRUPOS_PARTE_2:
        return "2a_parte"
    return "outros"


def _competencia_atual() -> str:
    return datetime.now().strftime("%Y-%m")


def _competencia_por_vencimento(data_venc: str) -> tuple[str, str, str]:
    m = re.match(r"^\s*(\d{2})/(\d{2})/(\d{4})\s*$", data_venc or "")
    if m:
        _, mm, yyyy = m.groups()
        mes = int(mm)
        ano = int(yyyy)
    else:
        now = datetime.now()
        mes = now.month
        ano = now.year

    mes_comp = mes - 1
    ano_comp = ano
    if mes_comp == 0:
        mes_comp = 12
        ano_comp -= 1

    return str(ano_comp), f"{mes_comp:02d}", f"{ano_comp:04d}-{mes_comp:02d}"


def _vencimento_alvo_padrao() -> str:
    return datetime.now().strftime("20/%m/%Y")


def _competencia_label(data_venc: str) -> str:
    ano, mes, _ = _competencia_por_vencimento(data_venc)
    return f"{mes}-{ano}"


def _chave_baixado(comp: str, grupo: str, doc: str) -> str:
    return f"{comp}|{grupo}|{_apenas_digitos(doc)}"


def _formatar_cnpj(doc: str) -> str:
    d = _apenas_digitos(doc)
    if len(d) == 14:
        return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"
    return doc or d


def _resolver_dir_com_fallback(dir_env: Path, fallbacks: list[Path], label: str) -> Path:
    if dir_env.exists():
        return dir_env
    for alt in fallbacks:
        if alt.exists():
            print(f"[warn] {label} do .env nao encontrado: {dir_env}")
            print(f"[info] usando fallback de {label}: {alt}")
            return alt
    return dir_env


def _slk_ler_celulas(path: Path) -> dict[tuple[int, int], str]:
    text = None
    for enc in ("latin-1", "utf-8"):
        try:
            text = path.read_text(encoding=enc, errors="strict")
            break
        except Exception:
            continue
    if text is None:
        text = path.read_text(encoding="latin-1", errors="ignore")

    cells: dict[tuple[int, int], str] = {}
    for line in text.splitlines():
        if not line.startswith("C;"):
            continue
        xm = re.search(r";X(\d+)", line)
        ym = re.search(r";Y(\d+)", line)
        if not xm or not ym:
            continue
        row = int(ym.group(1))
        col = int(xm.group(1))
        km = re.search(r';K"([^"]*)"', line) or re.search(r";K([^;]+)", line)
        if not km:
            continue
        cells[(row, col)] = km.group(1).strip()
    return cells


def carregar_documentos_do_slk(path: Path) -> list[dict]:
    cells = _slk_ler_celulas(path)
    rows = sorted({r for (r, _) in cells.keys()})
    seen: set[str] = set()
    out: list[dict] = []
    for r in rows:
        codigo = (cells.get((r, 1)) or "").strip()
        documento = (cells.get((r, 2)) or "").strip()
        nome = (cells.get((r, 3)) or "").strip()
        if not codigo and not documento and not nome:
            continue
        if codigo.lower() == "empresas" or documento.lower() == "empresas":
            continue
        doc = _apenas_digitos(documento)
        if len(doc) not in (11, 14):
            continue
        if doc in seen:
            continue
        seen.add(doc)
        out.append({"codigo": codigo or "SEM_CODIGO", "documento": doc, "nome": nome or "SEM_NOME"})
    return out

def localizar_slk_do_grupo(grupos_dir: Path, grupo: str) -> Path | None:
    exato = grupos_dir / f"{grupo}.slk"
    if exato.exists():
        return exato
    padrao = re.compile(rf"(^|\D){re.escape(grupo)}(\D|$)")
    for arq in sorted(grupos_dir.glob("*.slk")):
        if padrao.search(arq.stem):
            return arq
    return None


def carregar_fila_representados_por_grupo(grupos_dir: Path, grupos: list[str]) -> list[dict]:
    fila: list[dict] = []
    for grupo in grupos:
        slk = localizar_slk_do_grupo(grupos_dir, grupo)
        if not slk:
            print(f"[warn] Arquivo .slk do grupo {grupo} nao encontrado em: {grupos_dir}")
            continue
        docs = carregar_documentos_do_slk(slk)
        if not docs:
            print(f"[warn] Nenhum documento valido no grupo {grupo}: {slk.name}")
            continue
        parte = _parte_por_grupo(grupo)
        for d in docs:
            fila.append({"parte": parte, "grupo": grupo, "codigo": d["codigo"], "documento": d["documento"], "nome": d["nome"]})
        print(f"[ok] {parte} grupo {grupo}: {len(docs)} documento(s) de {slk.name}")
    return fila


def carregar_chaves_baixadas(path: Path) -> set[str]:
    if not path.exists():
        return set()
    out: set[str] = set()
    with path.open("r", encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            comp = (row.get("competencia") or "").strip()
            grp = (row.get("grupo") or "").strip()
            doc = _apenas_digitos((row.get("documento") or "").strip())
            if comp and grp and doc:
                out.add(_chave_baixado(comp, grp, doc))
    return out


def registrar_download_sucesso(path: Path, competencia: str, item: dict, arquivo_nome: str, pasta_destino: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    existe = path.exists()
    cols = ["timestamp", "competencia", "parte", "grupo", "codigo", "documento", "nome", "arquivo", "pasta_destino"]
    with path.open("a", encoding="utf-8", newline="") as f:
        wr = csv.DictWriter(f, fieldnames=cols)
        if not existe:
            wr.writeheader()
        wr.writerow(
            {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "competencia": competencia,
                "parte": item["parte"],
                "grupo": item["grupo"],
                "codigo": item["codigo"],
                "documento": item["documento"],
                "nome": item["nome"],
                "arquivo": arquivo_nome,
                "pasta_destino": str(pasta_destino),
            }
        )


def salvar_relatorio_erros_excel(path: Path, erros: list[dict]) -> None:
    if not erros:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    cols = ["timestamp", "competencia", "grupo", "cnpj", "empresa", "motivo"]
    with path.open("w", encoding="utf-8", newline="") as f:
        wr = csv.DictWriter(f, fieldnames=cols)
        wr.writeheader()
        for row in erros:
            wr.writerow(row)
    print(f"[ok] relatorio de erros salvo: {path}")


def preparar_pasta_holerite_bruto(base_dir: Path, codigo: str, nome: str) -> Path:
    mes = datetime.now().strftime("%m")
    dest = base_dir / _sanitizar_nome(codigo) / mes / _sanitizar_nome(nome)
    dest.mkdir(parents=True, exist_ok=True)
    return dest


def close_warning_if_present(page) -> bool:
    try:
        btn = page.locator("button.br-button._footerAction_slfpf_24.secondary:has-text('Fechar')").first
        if btn.count() > 0 and btn.is_visible():
            btn.click(timeout=2000)
            return True
    except Exception:
        pass
    return False


def confirmar_gerar_guia_se_aparecer(page) -> bool:
    for sel in [
        "div:has-text('Gerar guia') button.br-button.primary:has-text('Confirmar')",
        "div:has-text('Gerar guia') button:has-text('Confirmar')",
    ]:
        try:
            btn = page.locator(sel).first
            if btn.count() > 0 and btn.is_visible():
                btn.click(timeout=6000, force=True)
                print("[ok] modal 'Gerar guia' confirmado.")
                return True
        except Exception:
            continue
    return False


def abrir_modal_trocar_perfil(page) -> None:
    for sel in [
        "button.br-button.secondary.botao-barra-perfil:has-text('Trocar Perfil')",
        "button:has-text('Trocar Perfil')",
    ]:
        try:
            btn = page.locator(sel).first
            if btn.count() > 0 and btn.is_visible():
                btn.click(timeout=10000)
                return
        except Exception:
            continue
    raise RuntimeError("Nao consegui abrir o modal de Trocar Perfil.")


def localizar_modal_perfil(page):
    for loc in [
        page.locator("div[role='dialog']:has-text('Empregador a ser representado')").first,
        page.locator("div:has-text('Empregador a ser representado')").first,
        page.locator("div[role='dialog']:has-text('Definir Perfil'), div[role='dialog']:has-text('Trocar Perfil')").first,
    ]:
        try:
            if loc.count() > 0 and loc.is_visible():
                return loc
        except Exception:
            continue
    return page.locator("div:has-text('Definir Perfil'), div:has-text('Trocar Perfil')").first

def preencher_modal_perfil(page, documento: str) -> None:
    modal = localizar_modal_perfil(page)
    try:
        modal.wait_for(state="visible", timeout=8000)
    except Exception:
        # Subsequentes: garante que o modal seja aberto de novo.
        try:
            abrir_modal_trocar_perfil(page)
        except Exception:
            pass
        modal = localizar_modal_perfil(page)
        modal.wait_for(state="visible", timeout=30000)

    selecionou = False
    for tentativa in range(1, 7):
        print(f"[info] preenchendo Perfil com 'Procurador' (tentativa {tentativa}/6)...")
        try:
            perfil_ok = False

            # Tenta abrir o ng-select de Perfil por clique no container/seta.
            for sel in [
                "xpath=.//*[contains(normalize-space(.),'Perfil')]/following::ng-select[1]",
                "ng-select",
            ]:
                try:
                    ng = modal.locator(sel).first
                    if ng.count() == 0:
                        continue
                    ng.scroll_into_view_if_needed(timeout=2000)
                    ng.click(timeout=3000, force=True)
                    seta = ng.locator(".ng-arrow-wrapper").first
                    if seta.count() > 0:
                        seta.click(timeout=2000, force=True)
                    break
                except Exception:
                    continue

            # Tenta escolher clicando na opcao visivel "Procurador".
            for op_sel in [
                "div.ng-dropdown-panel div.ng-option span:has-text('Procurador')",
                "div.ng-dropdown-panel div.ng-option:has-text('Procurador')",
                "div[role='option']:has-text('Procurador')",
                "li:has-text('Procurador')",
            ]:
                try:
                    op = page.locator(op_sel).first
                    if op.count() > 0 and op.is_visible():
                        op.click(timeout=3000, force=True)
                        perfil_ok = True
                        break
                except Exception:
                    continue

            # Sempre reforca por digitacao para garantir os subsequentes.
            perfil = modal.locator("input[role='combobox']").first
            perfil.wait_for(state="visible", timeout=3000)
            perfil.click(timeout=3000, force=True)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.keyboard.type("Procurador", delay=35)
            time.sleep(0.25)
            page.keyboard.press("Enter")
            time.sleep(0.25)

            # Valida se o valor realmente ficou selecionado no campo.
            for ver_sel in [
                "span.ng-value-label:has-text('Procurador')",
                "ng-select.ng-has-value:has-text('Procurador')",
                "div.ng-value:has-text('Procurador')",
            ]:
                try:
                    v = modal.locator(ver_sel).first
                    if v.count() > 0 and v.is_visible():
                        perfil_ok = True
                        break
                except Exception:
                    continue

            # Fallback final de validacao: se o combobox contem "Procurador", segue.
            if not perfil_ok:
                try:
                    perfil_input = modal.locator("input[role='combobox']").first
                    val = (perfil_input.input_value(timeout=1000) or "").strip().lower()
                    if "procurador" in val:
                        perfil_ok = True
                except Exception:
                    pass

            if perfil_ok:
                selecionou = True
                break
        except Exception:
            pass
        time.sleep(0.5)

    if not selecionou:
        raise RuntimeError("Nao consegui preencher o campo Perfil com 'Procurador'.")

    doc_input = None
    for doc_sel in [
        "input[placeholder*='Informe CNPJ' i]",
        "input[placeholder*='CNPJ' i], input[placeholder*='CPF' i]",
        "xpath=.//*[contains(normalize-space(.),'Empregador a ser representado')]/following::input[1]",
    ]:
        try:
            cand = modal.locator(doc_sel).first
            if cand.count() > 0:
                cand.wait_for(state="visible", timeout=4000)
                doc_input = cand
                break
        except Exception:
            continue

    if doc_input is None:
        raise RuntimeError("Nao consegui localizar o campo de CNPJ/CPF no modal de perfil.")

    digits = _apenas_digitos(documento)
    preenchido = False
    for _ in range(3):
        try:
            doc_input.scroll_into_view_if_needed(timeout=1500)
        except Exception:
            pass
        try:
            doc_input.click(timeout=3000, force=True)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.keyboard.type(digits, delay=22)
            time.sleep(0.15)
            atual = (doc_input.input_value(timeout=1000) or "").strip()
            if _apenas_digitos(atual) == digits:
                preenchido = True
                break
        except Exception:
            pass
        try:
            doc_input.fill(digits, timeout=3000)
            atual = (doc_input.input_value(timeout=1000) or "").strip()
            if _apenas_digitos(atual) == digits:
                preenchido = True
                break
        except Exception:
            pass
        time.sleep(0.25)

    if not preenchido:
        raise RuntimeError(f"Nao consegui preencher o campo CNPJ/CPF para documento: {documento}")

    clicou = False
    for sel in [
        "button.br-button.primary:has-text('Selecionar')",
        "button.br-button.primary:has-text('Definir')",
        "button:has-text('Selecionar')",
        "button:has-text('Definir')",
    ]:
        try:
            btn = modal.locator(sel).first
            if btn.count() > 0 and btn.is_visible():
                btn.click(timeout=8000, force=True)
                clicou = True
                break
        except Exception:
            continue

    if not clicou:
        raise RuntimeError("Nao consegui clicar em Selecionar/Definir no modal de perfil.")

    try:
        modal.wait_for(state="hidden", timeout=15000)
    except Exception:
        time.sleep(1.0)

    print(f"[ok] modal preenchido para documento: {documento}")


def esperar_perfil_aplicado(page, documento: str, timeout_ms: int = 20000) -> None:
    alvo = _apenas_digitos(documento)
    limite = time.time() + (timeout_ms / 1000.0)
    while time.time() < limite:
        try:
            txt = page.locator("span.dados-perfil, div:has-text('Empregador:'), body").first.inner_text(timeout=2000)
        except Exception:
            txt = ""
        if "Usuário (Procurador)" in txt and alvo and alvo in _apenas_digitos(txt):
            print(f"[ok] perfil aplicado para documento: {documento}")
            return
        time.sleep(0.5)
    raise RuntimeError(f"Nao confirmou cabecalho de perfil para o documento: {documento}")


def abrir_emissao_rapida_e_pesquisar(page) -> None:
    page.goto(EMISSAO_GUIA_RAPIDA_URL, wait_until="domcontentloaded")

    for _ in range(4):
        try:
            marcado = page.evaluate(
                """() => {
                    const el = document.querySelector('#h-checkbox-2');
                    return el ? !!el.checked : null;
                }"""
            )
        except Exception:
            marcado = None

        if marcado is False:
            break

        try:
            page.locator("label[for='h-checkbox-2']").first.click(timeout=3000)
        except Exception:
            try:
                page.locator("label:has-text('Rescisório')").first.click(timeout=3000)
            except Exception:
                pass
        time.sleep(0.3)

    try:
        page.locator("button.br-button.secondary.ml-2:has-text('Pesquisar')").first.click(timeout=10000)
    except Exception:
        page.get_by_role("button", name="Pesquisar").first.click(timeout=10000)
    print("[ok] botao 'Pesquisar' clicado em Emissao Guia Rapida.")


def baixar_guia_fgts_por_vencimento(page, data_vencimento_alvo: str, pasta_destino: Path) -> tuple[Path | None, str]:
    alvo = (data_vencimento_alvo or "").strip()
    if not alvo:
        return None, "data_vencimento_alvo_vazia"

    pasta_destino.mkdir(parents=True, exist_ok=True)

    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass
    time.sleep(0.8)

    # Regra: se a propria tela informar que nao ha debitos, marca erro e segue.
    try:
        texto_tela = page.locator("body").inner_text(timeout=2500)
    except Exception:
        texto_tela = ""
    texto_norm = _normalizar_texto_busca(texto_tela)
    if "nao ha debitos de interesse" in texto_norm:
        return None, "sem_debitos_de_interesse"

    try:
        marcado = page.evaluate(
            """(alvo) => {
                for (const old of document.querySelectorAll('[data-cdx-emitir-hit="1"]')) {
                    old.removeAttribute('data-cdx-emitir-hit');
                }
                const blocos = Array.from(document.querySelectorAll('div, tr, section, article'));
                const host = blocos.find(el => (el.innerText || '').includes(`Vencimento da Guia: ${alvo}`) || (el.innerText || '').includes(alvo));
                if (!host) return false;
                const botoes = Array.from(host.querySelectorAll('button.ml-2.br-button.primary, button.br-button.primary, button'));
                const btn = botoes.find(b => /emitir\\s*guia|emitir/i.test((b.innerText || '').trim()));
                if (!btn) return false;
                btn.setAttribute('data-cdx-emitir-hit', '1');
                btn.scrollIntoView({block:'center', inline:'center'});
                return true;
            }""",
            alvo,
        )
    except Exception:
        marcado = False

    if not marcado:
        return None, f"botao_emitir_guia_nao_encontrado_{alvo}"

    try:
        btn = page.locator("button[data-cdx-emitir-hit='1']").first
        with page.expect_download(timeout=60000) as dlinfo:
            btn.click(timeout=10000, force=True)
            time.sleep(0.4)
            confirmar_gerar_guia_se_aparecer(page)
        download = dlinfo.value
    except Exception:
        return None, f"download_nao_gerado_{alvo}"

    nome = _sanitizar_nome(download.suggested_filename or "guia_fgts")
    if not nome.lower().endswith(".pdf"):
        nome = f"{nome}.pdf"
    dest = pasta_destino / nome
    download.save_as(str(dest))
    print(f"[ok] download salvo em bruto: {dest}")
    return dest, ""


def baixar_guia_fgts_com_retentativas(page, data_vencimento_alvo: str, pasta_destino: Path, tentativas: int) -> tuple[Path | None, str]:
    ultimo = "download_falhou"
    for i in range(1, max(1, tentativas) + 1):
        print(f"[info] tentativa de download {i}/{max(1, tentativas)}")
        arq, motivo = baixar_guia_fgts_por_vencimento(page, data_vencimento_alvo, pasta_destino)
        if arq:
            return arq, ""
        if motivo == "sem_debitos_de_interesse":
            return None, motivo
        ultimo = motivo or "download_falhou"
        if i < max(1, tentativas):
            time.sleep(1.0)
    return None, ultimo


def organizar_download_no_automatizado(item: dict, arquivo_baixado: Path, data_vencimento_alvo: str) -> Path:
    parte = item.get("parte", "")
    grupo = _sanitizar_nome(item.get("grupo", "SEM_GRUPO"))
    codigo = _sanitizar_nome(item.get("codigo", "SEM_CODIGO"))
    nome_empresa = _sanitizar_nome(item.get("nome", "SEM_NOME"))
    competencia_label = _competencia_label(data_vencimento_alvo)
    ano_comp, mes_comp, _ = _competencia_por_vencimento(data_vencimento_alvo)

    base_1a = Path(os.environ.get("FGTS_AUTOMATIZADO_1A", str(AUTOMATIZADO_1A)))
    base_2a = Path(os.environ.get("FGTS_AUTOMATIZADO_2A", str(AUTOMATIZADO_2A)))
    if not base_1a.exists():
        base_1a = _resolver_dir_com_fallback(base_1a, [Path(r"\\192.0.0.251\Arquivos\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1ª PARTE"), AUTOMATIZADO_1A], "FGTS_AUTOMATIZADO_1A")
    if not base_2a.exists():
        base_2a = _resolver_dir_com_fallback(base_2a, [Path(r"\\192.0.0.251\Arquivos\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\2ª PARTE"), AUTOMATIZADO_2A], "FGTS_AUTOMATIZADO_2A")

    base_raiz = base_1a if parte == "1a_parte" else base_2a
    pasta = base_raiz / ano_comp / mes_comp / grupo / nome_empresa
    pasta.mkdir(parents=True, exist_ok=True)

    destino = pasta / f"{nome_empresa} - FGTS {competencia_label}.pdf"
    if destino.exists():
        i = 2
        while True:
            alt = pasta / f"{nome_empresa} - FGTS {competencia_label} ({i}).pdf"
            if not alt.exists():
                destino = alt
                break
            i += 1

    shutil.copy2(str(arquivo_baixado), str(destino))
    print(f"[ok] arquivo organizado em automatizado: {destino}")
    return destino


def wait_for_login(page) -> None:
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    try:
        page.locator("button.br-button.is-primary.entrar:has-text('Entrar com GOV.BR')").first.click(timeout=15000)
    except Exception:
        try:
            page.get_by_role("button", name="Entrar com GOV.BR").first.click(timeout=10000)
        except Exception:
            print("[warn] botao 'Entrar com GOV.BR' nao foi encontrado.")

    try:
        page.locator("button:has-text('Seu certificado digital')").first.click(timeout=10000)
    except Exception:
        print("[warn] botao 'Seu certificado digital' nao apareceu ainda.")

    print("[info] aguardando redirecionamento para /portal/escolhaPerfil...")
    page.wait_for_url("**/portal/escolhaPerfil**", timeout=120000)
    close_warning_if_present(page)


def ir_para_servicos_e_abrir_trocar_perfil(page) -> None:
    page.goto(SERVICOS_URL, wait_until="domcontentloaded")
    close_warning_if_present(page)
    abrir_modal_trocar_perfil(page)


def cert_exists_in_windows(subject: str = "", issuer: str = "", thumbprint: str = "") -> bool:
    cmd = (
        "Get-ChildItem Cert:\\CurrentUser\\My | "
        "Select-Object Subject,Issuer,Thumbprint | "
        "ConvertTo-Json -Depth 3"
    )
    res = subprocess.run(["powershell", "-NoProfile", "-Command", cmd], capture_output=True, text=True, check=False)
    if res.returncode != 0:
        print("[warn] Nao consegui validar certificado no Windows. Continuando mesmo assim.")
        return True
    try:
        certs = json.loads(res.stdout) if res.stdout else []
    except Exception:
        certs = []
    if isinstance(certs, dict):
        certs = [certs]

    subject_l = subject.lower()
    issuer_l = issuer.lower()
    thumb_l = thumbprint.lower().replace(" ", "")
    for cert in certs:
        cert_subject = str(cert.get("Subject", "")).lower()
        cert_issuer = str(cert.get("Issuer", "")).lower()
        cert_thumb = str(cert.get("Thumbprint", "")).lower().replace(" ", "")
        if ((not subject_l) or (subject_l in cert_subject)) and ((not issuer_l) or (issuer_l in cert_issuer)) and ((not thumb_l) or (thumb_l == cert_thumb)):
            return True
    return False

def _extract_dn_field(dn_value: str, field: str) -> str:
    target = f"{field.upper()}="
    for part in dn_value.split(","):
        cur = part.strip()
        if cur.upper().startswith(target):
            return cur[len(target):].strip()
    return ""


def apply_chrome_auto_select_policy(patterns: list[str], cert_subject: str = "", cert_issuer: str = "", cert_subject_cn: str = "", cert_issuer_cn: str = "") -> list[tuple[str, bool, str]]:
    subject_cn = cert_subject_cn or _extract_dn_field(cert_subject, "CN")
    issuer_cn = cert_issuer_cn or _extract_dn_field(cert_issuer, "CN")

    cert_filter = {}
    if subject_cn:
        cert_filter["SUBJECT"] = {"CN": subject_cn}
    if issuer_cn:
        cert_filter["ISSUER"] = {"CN": issuer_cn}

    reg_path = r"Software\Policies\Google\Chrome\AutoSelectCertificateForUrls"
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
    backups: list[tuple[str, bool, str]] = []
    try:
        for idx, pattern in enumerate(patterns, start=1):
            value_name = str(idx)
            payload = {"pattern": pattern, "filter": cert_filter}
            payload_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

            prev_exists = False
            prev_value = ""
            try:
                prev_value, _ = winreg.QueryValueEx(key, value_name)
                prev_exists = True
            except FileNotFoundError:
                pass

            backups.append((value_name, prev_exists, prev_value))
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, payload_str)
            print(f"[ok] politica AutoSelectCertificateForUrls aplicada [{value_name}]: {payload_str}")
    finally:
        winreg.CloseKey(key)
    return backups


def restore_chrome_auto_select_policy(backups: list[tuple[str, bool, str]]) -> None:
    reg_path = r"Software\Policies\Google\Chrome\AutoSelectCertificateForUrls"
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
    try:
        for value_name, prev_exists, prev_value in backups:
            if prev_exists:
                winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, prev_value)
                print(f"[ok] politica AutoSelectCertificateForUrls restaurada ao valor anterior [{value_name}].")
            else:
                try:
                    winreg.DeleteValue(key, value_name)
                    print(f"[ok] politica AutoSelectCertificateForUrls removida ao finalizar [{value_name}].")
                except FileNotFoundError:
                    print(f"[info] politica AutoSelectCertificateForUrls ja estava ausente ao finalizar [{value_name}].")
    finally:
        winreg.CloseKey(key)


def fechar_todos_chromes() -> None:
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", "Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force"],
            check=False,
            capture_output=True,
            text=True,
        )
        print("[info] todos os Chrome abertos foram fechados.")
    except Exception:
        print("[warn] nao consegui fechar processos do Chrome automaticamente.")


def processar_representados(
    p,
    fila: list[dict],
    holerite_dir: Path,
    registro_path: Path,
    data_vencimento_alvo: str,
    erros_path: Path,
    tentativas_download: int,
    competencia: str,
) -> None:
    chaves_baixadas = carregar_chaves_baixadas(registro_path)
    print(f"[info] competencia de controle: {competencia}")
    print(f"[info] registros ja baixados carregados: {len(chaves_baixadas)}")

    erros: list[dict] = []
    browser = None
    context = None
    try:
        browser = chrome_9222(p, PORT)
        context = browser.new_context()
        page = context.new_page()
        wait_for_login(page)

        for idx, item in enumerate(fila, start=1):
            grupo = item["grupo"]
            codigo = item["codigo"]
            nome = item["nome"]
            documento = item["documento"]
            chave = _chave_baixado(competencia, grupo, documento)
            if chave in chaves_baixadas:
                print(f"[skip] ja registrado como baixado: competencia={competencia} grupo={grupo} documento={documento}")
                continue

            pasta_destino = preparar_pasta_holerite_bruto(holerite_dir, codigo, nome)
            print(f"[info] {idx}/{len(fila)} parte={item['parte']} grupo={grupo} codigo={codigo} documento={documento}")
            print(f"[info] empresa={nome}")
            print(f"[info] pasta arquivos brutos: {pasta_destino}")

            try:
                if idx > 1:
                    ir_para_servicos_e_abrir_trocar_perfil(page)
                preencher_modal_perfil(page, documento)
                esperar_perfil_aplicado(page, documento)
                abrir_emissao_rapida_e_pesquisar(page)

                arquivo_baixado, motivo = baixar_guia_fgts_com_retentativas(page, data_vencimento_alvo, pasta_destino, tentativas_download)
                if not arquivo_baixado:
                    print("[info] sem download valido para a regra de vencimento; registrar como erro.")
                    erros.append({
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "competencia": competencia,
                        "grupo": grupo,
                        "cnpj": _formatar_cnpj(documento),
                        "empresa": nome,
                        "motivo": motivo or "download_nao_encontrado",
                    })
                    continue

                destino_auto = organizar_download_no_automatizado(item, arquivo_baixado, data_vencimento_alvo)
                registrar_download_sucesso(registro_path, competencia, item, destino_auto.name, destino_auto.parent)
                chaves_baixadas.add(chave)
                print(f"[ok] empresa marcada como baixada para a competencia {competencia}.")
            except Exception as exc:
                print(f"[erro] falha no documento {documento}: {exc}")
                print("[info] seguindo para o proximo documento.")
                erros.append({
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "competencia": competencia,
                    "grupo": grupo,
                    "cnpj": _formatar_cnpj(documento),
                    "empresa": nome,
                    "motivo": str(exc),
                })

        salvar_relatorio_erros_excel(erros_path, erros)
    finally:
        try:
            if context:
                context.close()
        except Exception:
            pass
        try:
            if browser:
                browser.close()
        except Exception:
            pass
        fechar_todos_chromes()


def main() -> None:
    env = merge_env_with_session(load_env(ENV_PATH))

    cert_subject = env.get("CERT_SUBJECT", "").strip()
    cert_issuer = env.get("CERT_ISSUER", "").strip()
    cert_thumbprint = env.get("CERT_THUMBPRINT", "").strip()
    cert_subject_cn = env.get("CERT_SUBJECT_CN", "").strip()
    cert_issuer_cn = env.get("CERT_ISSUER_CN", "").strip()

    grupos = _parse_grupos(env.get("FGTS_GRUPOS", "6;13;14;31;32;33;34;35;36;37"))
    grupos_dir = Path(env.get("FGTS_GRUPOS_DIR", str(BASE_DIR / "Grupos")))
    holerite_dir = Path(env.get("FGTS_ARQUIVOS_BRUTOS_DIR", env.get("FGTS_HOLERITE_BRUTO_DIR", str(BASE_DIR / "Arquivos Brutos"))))

    grupos_dir = _resolver_dir_com_fallback(grupos_dir, [BASE_DIR / "Bruto", BASE_DIR / "Grupos"], "FGTS_GRUPOS_DIR")
    holerite_dir = _resolver_dir_com_fallback(holerite_dir, [BASE_DIR / "Bruto", BASE_DIR / "Arquivos Brutos"], "FGTS_ARQUIVOS_BRUTOS_DIR")

    registro_baixados = Path(env.get("FGTS_REGISTRO_BAIXADOS", str(REGISTRO_BAIXADOS_PADRAO)))
    registro_erros = Path(env.get("FGTS_REGISTRO_ERROS", str(REGISTRO_ERROS_PADRAO)))
    tentativas_download = max(1, int(env.get("FGTS_TENTATIVAS_DOWNLOAD", "2").strip() or "2"))
    data_vencimento_alvo = env.get("FGTS_DATA_VENCIMENTO_ALVO", _vencimento_alvo_padrao()).strip()
    _, _, competencia = _competencia_por_vencimento(data_vencimento_alvo)
    print(f"[info] filtro de vencimento para emitir guia: {data_vencimento_alvo}")
    print(f"[info] competencia de organizacao (mes anterior): {competencia}")

    fila = carregar_fila_representados_por_grupo(grupos_dir, grupos)
    if not fila:
        print("[erro] Nenhum representado encontrado nos .slk. Verifique FGTS_GRUPOS_DIR e arquivos de grupo.")
        return

    cert_pattern = env.get("CERT_AUTOSELECT_PATTERN", "https://certificado.sso.acesso.gov.br;https://sso.acesso.gov.br;https://[*.]mte.gov.br").strip()
    cert_autoselect_enabled = env.get("CERT_AUTOSELECT_ENABLED", "1").strip().lower() in ("1", "true", "sim", "yes")
    patterns = [p.strip() for p in cert_pattern.split(";") if p.strip()]

    policy_backups: list[tuple[str, bool, str]] = []
    policy_applied = False
    try:
        if cert_autoselect_enabled:
            policy_backups = apply_chrome_auto_select_policy(
                patterns=patterns or ["https://certificado.sso.acesso.gov.br", "https://sso.acesso.gov.br"],
                cert_subject=cert_subject,
                cert_issuer=cert_issuer,
                cert_subject_cn=cert_subject_cn,
                cert_issuer_cn=cert_issuer_cn,
            )
            policy_applied = True
        else:
            print("[info] politica de auto selecao desativada por CERT_AUTOSELECT_ENABLED.")

        if cert_subject or cert_issuer or cert_thumbprint:
            if cert_exists_in_windows(cert_subject, cert_issuer, cert_thumbprint):
                print("[ok] certificado localizado no Windows (filtro do .env).")
            else:
                print("[warn] certificado do .env nao foi encontrado em Cert:\\CurrentUser\\My.")

        with sync_playwright() as p:
            processar_representados(
                p,
                fila,
                holerite_dir,
                registro_baixados,
                data_vencimento_alvo,
                registro_erros,
                tentativas_download,
                competencia,
            )
            print("[ok] fluxo FGTS concluido.")
    finally:
        if cert_autoselect_enabled and policy_applied:
            restore_chrome_auto_select_policy(policy_backups)


if __name__ == "__main__":
    main()



