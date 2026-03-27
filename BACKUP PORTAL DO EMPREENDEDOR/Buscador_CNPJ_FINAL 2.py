import os
import time
from datetime import datetime
import tempfile
import subprocess
import urllib.request
from pathlib import Path
from shutil import which
from playwright.sync_api import sync_playwright, expect, TimeoutError as PWTimeout
import re
import csv
from chrome_9222 import chrome_9222, PORT
from playwright.sync_api import sync_playwright

mes_ano = datetime.now().strftime("%m-%Y")
data_consulta = datetime.now().strftime("%d-%m-%Y")

CSV_PATH = r"\\192.0.0.251\Arquivos\DOCUMENTOS ESCRITORIO\RH\DET\lista empresas.csv"
LOG_DIR = r"\\192.0.0.251\Arquivos\DOCUMENTOS ESCRITORIO\RH\DET"

# agora log vai pra pasta do mês, junto com os prints
LOG_FOLDER = os.path.join(LOG_DIR, mes_ano)
os.makedirs(LOG_FOLDER, exist_ok=True)

LOG_FILE = os.path.join(LOG_FOLDER, f"log caixa postal {data_consulta}.csv")

def ler_lista_empresas(caminho_csv):
    empresas = []
    try:
        with open(caminho_csv, mode='r', encoding='cp1252') as file:
            reader = csv.reader(file, delimiter=';')
            next(reader)  # Pula o cabeçalho
            for row in reader:
                if len(row) >= 2:
                    cnpj = row[0].strip()
                    razao_social = row[1].strip()
                    empresas.append({'cnpj': cnpj, 'empresa': razao_social})
    except Exception as e:
        print(f"Erro ao ler o arquivo CSV: {e}")
    return empresas

def abrir_trocar_perfil(page):
    # Diferentes formas de clicar no botão "Trocar Perfil"
    seletores = [
        "button.br-button:has-text('Trocar Perfil')",
        "a.br-button:has-text('Trocar Perfil')",
        "[aria-label*='Trocar Perfil']",
        "text=Trocar Perfil",
        "button:has(i.fa-user-friends), a:has(i.fa-user-friends)",
    ]
    clicou = False
    for sel in seletores:
        try:
            el = page.locator(sel).first
            el.scroll_into_view_if_needed(timeout=1500)
            el.click(timeout=3000)
            clicou = True
            break
        except Exception:
            continue

    if not clicou:
        # última tentativa: procurar pelo texto em toda a página e forçar o clique
        try:
            el = page.get_by_text(re.compile(r"Trocar\s+Perfil", re.I)).first
            el.scroll_into_view_if_needed()
            el.click(timeout=3000, force=True)
            clicou = True
        except Exception:
            pass

    # Espera algum elemento característico do modal
    # (qualquer um desses vale como "modal aberto")
    alvos_modal = [
        page.get_by_placeholder(re.compile(r"Informe CNPJ", re.I)),
        page.get_by_role("button", name=re.compile(r"Selecionar", re.I)),
        page.get_by_text(re.compile(r"Perfil\s*\(Obrigatório\)", re.I)),
        page.get_by_text(re.compile(r"Trocar Perfil", re.I)),  # título do modal
    ]

    for alvo in alvos_modal:
        try:
            expect(alvo).to_be_visible(timeout=8000)
            # retorna o container do modal (o ancestral mais próximo)
            return alvo.locator("xpath=ancestor::div[contains(@class,'modal') or @role='dialog'][1]")
        except Exception:
            continue

    # Debug: salva screenshot para ver o que ficou na tela
    page.screenshot(path="debug_falha_trocar_perfil.png", full_page=True)
    raise AssertionError("Não consegui abrir o modal 'Trocar Perfil'. Screenshot salvo em debug_falha_trocar_perfil.png")

def selecionar_perfil_procurador(page, dialog):
    # Abre o combo de Perfil (pode ser um br-select / ng-select)
    # 1ª tentativa: usar role=combobox
    try:
        dialog.get_by_role("combobox", name=re.compile(r"Perfil", re.I)).click()
    except Exception:
        # Fallback: clicar no primeiro combobox do modal
        dialog.get_by_role("combobox").first.click()
    # Seleciona a opção "Procurador"
    try:
        page.get_by_role("option", name=re.compile(r"Procurador", re.I)).click()
    except Exception:
        # Fallback por texto genérico em listas
        page.locator("text=Procurador").first.click()

def preencher_cnpj_e_selecionar(dialog, cnpj: str):
    # pega diretamente o textbox acessível do modal
    campo = dialog.get_by_role("textbox", name=re.compile(r"Empregador\s+a\s+ser\s+Representado", re.I))
    expect(campo).to_be_visible(timeout=5000)

    # limpar e digitar (máscara às vezes atrapalha .fill)
    campo.click()
    try:
        campo.fill("")  # se a máscara não permitir, caímos no except e usamos Ctrl+A
    except Exception:
        pass
    campo.press("Control+A")
    campo.type(cnpj)

    # clicar em Selecionar
    dialog.get_by_role("button", name=re.compile(r"Selecionar", re.I)).click()

def trocar_perfil_para_cada_cnpj(page, caminho_csv: str):
    cnpjs = ler_lista_empresas(caminho_csv)
    for cnpj in cnpjs:
        print(f"Trocando perfil para CNPJ: {cnpj}")
        dialog = abrir_trocar_perfil(page)
        selecionar_perfil_procurador(page, dialog)
        preencher_cnpj_e_selecionar(dialog, cnpj)
        page.wait_for_load_state("domcontentloaded")
        abrir_caixa_postal(page)

def garantir_log() -> None:
    Path(LOG_FOLDER).mkdir(parents=True, exist_ok=True)  # já cria a pasta do mês
    novo = not Path(LOG_FILE).exists()
    if novo:
        with open(LOG_FILE, "w", newline="", encoding="cp1252") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["CNPJ", "Empresa", "Remetente", "Titulo", "Data"])

def escrever_linha_msg(cnpj: str, empresa: str, remetente: str, titulo: str, data: str) -> None:
    with open(LOG_FILE, "a", newline="", encoding="cp1252") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow([cnpj, empresa, remetente, titulo, data])

def abrir_caixa_postal(page):
    target_url_re = re.compile(r"^https://det\.sit\.trabalho\.gov\.br/caixapostal", re.I)
    label_re = re.compile(r"Caixa Postal", re.I)

    # Tenta como link/botão/qualquer texto (tentativa inicial, como no original)
    for loc in [
        page.get_by_role("link", name=label_re),
        page.get_by_role("button", name=label_re),
        page.locator("text=Caixa Postal").first,
    ]:
        try:
            loc.click(timeout=1000)
            break
        except Exception:
            pass

    # espera a página da caixa postal
    page.wait_for_load_state("domcontentloaded")

    # Após o domcontentloaded, tenta clicar e verificar a URL alvo.
    # Se não navegar para a URL desejada, aguarda 3s e tenta novamente.
    for _ in range(5):  # ajuste o número de tentativas se necessário
        # tenta clicar em qualquer uma das formas
        clicked = False
        for loc in [
            page.get_by_role("link", name=label_re),
            page.get_by_role("button", name=label_re),
            page.locator("text=Caixa Postal").first,
        ]:
            try:
                loc.click(timeout=2000)
                clicked = True
                break
            except Exception:
                continue

        # Se conseguiu clicar, espera a navegação para a URL alvo
        if clicked:
            try:
                page.wait_for_url(target_url_re, timeout=5000)
                # sucesso: já está na página da caixa postal -> seguir o script
                return
            except Exception:
                pass  # não navegou para a URL esperada dentro do timeout

        # Se ainda não está na URL alvo, aguarda 3s e tenta novamente
        if not target_url_re.search(page.url):
            time.sleep(3)
        else:
            return  # já está na URL alvo

    # Fallback: se após as tentativas não navegou, ao menos garante que o título existe no DOM
    page.get_by_role("heading", name=label_re).wait_for(state="attached", timeout=30000)
    
def extrair_empregador_info(page):
    """
    Lê o cabeçalho 'Empregador: 00.000.000/0000-00 | RAZÃO'
    """
    body = page.locator("body").inner_text()
    m = re.search(r"Empregador:\s*([\d\.\-/]+)\s*\|\s*([^\n]+)", body)
    cnpj = m.group(1) if m else ""
    empresa = m.group(2).strip() if m else ""
    # normaliza CNPJ só dígitos (opcional)
    cnpj_dig = re.sub(r"\D", "", cnpj)
    return cnpj_dig or cnpj, empresa

def localizar_lista_itens(page):
    """
    Retorna somente os cards de mensagem na coluna esquerda.
    Evita capturar cabeçalhos/paginação.
    """
    itens = page.locator("div.menu-caixa-postal div.tabela_mensagens div.tabela")
    return itens

def parse_bloco_texto(txt: str):
    """
    Extrai (titulo, remetente, data) do bloco de texto de um item da lista à esquerda.
    Normaliza data para "dd/mm/aaaa".
    Se for 'hoje', substitui pela data atual.
    """
    linhas = [l.strip() for l in txt.splitlines() if l.strip()]
    titulo = ""
    remetente = ""
    data = ""

    if not linhas:
        return titulo, remetente, data

    # 1) título
    if re.search(r"^(Aviso|Notifica)", linhas[0], re.I):
        titulo = linhas[1] if len(linhas) > 1 else ""
        possivel_linha_remetente = linhas[2] if len(linhas) > 2 else ""
    else:
        titulo = linhas[0]
        possivel_linha_remetente = linhas[1] if len(linhas) > 1 else ""

    padrao_data = PADRAO_DATA.pattern  # reutiliza o mesmo padrão
    matches = list(PADRAO_DATA.finditer(txt))
    if matches:
        data = normalizar_data(matches[-1].group(0))

    # 3) remetente
    if possivel_linha_remetente:
        linha_sem_data = re.sub(padrao_data, "", possivel_linha_remetente, flags=re.I).strip(" -–|:/\t ")
        linha_sem_data = re.sub(r"(?i)^(órgão\s*emissor|origem|remetente)\s*[:\-–]\s*", "", linha_sem_data).strip()
        remetente = linha_sem_data

    # fallback para remetente se ainda não achou
    if not remetente and linhas:
        for ln in reversed(linhas):
            if re.search(padrao_data, ln, re.I):
                tmp = re.sub(padrao_data, "", ln, flags=re.I).strip(" -–|:/\t ")
                tmp = re.sub(r"(?i)^(órgão\s*emissor|origem|remetente)\s*[:\-–]\s*", "", tmp).strip()
                if tmp:
                    remetente = tmp
                    break

    return titulo, remetente, data

def detalhar_mensagem(page):
    """
    Lê o painel direito após o clique no item.
    Procura containers típicos e devolve o texto.
    """
    seletores_conteudo = [
        "div.container-message",          # já no seu script
        "section#detalheMensagem",
        "div.conteudo-mensagem",
        "div.br-card .conteudo",          # comum em GOVBR
        "article",                        # fallback genérico
    ]
    for sel in seletores_conteudo:
        try:
            box = page.locator(sel).first
            box.wait_for(state="visible", timeout=10000)
            texto = box.inner_text().strip()
            if texto and len(texto.split()) > 3:
                return texto
        except Exception:
            pass
    return "N/A"

def coletar_e_logar_mensagens(page, cnpj, empresa):
    garantir_log()

    # 1) encontre os itens da lista da esquerda
    itens = localizar_lista_itens(page)
    total = itens.count()
    print(f"Encontradas {total} mensagens para esta empresa.")

    if total == 0:
        escrever_linha_msg(cnpj, empresa, "", "Sem mensagens", "", "")
        return

    # 2) percorra visivelmente (evita stale locators)
    for i in range(total):
        item = itens.nth(i)

        # tente obter o bloco de texto do item antes do clique
        try:
            # alguns itens precisam de scroll para renderizar completamente
            item.scroll_into_view_if_needed(timeout=3000)
            bloco = item.inner_text(timeout=3000)
        except Exception:
            bloco = ""

        tipo, assunto, data = parse_bloco_texto(bloco)

        # 3) clique para abrir no painel direito
        clicou = False
        for tentativa in range(2):
            try:
                item.click(timeout=3000)
                clicou = True
                break
            except Exception:
                # fallback: clique em link/botão interno
                try:
                    item.get_by_role("link").first.click(timeout=2000)
                    clicou = True
                    break
                except Exception:
                    pass

        if not clicou:
            print(f"[Aviso] Não consegui clicar no item {i}. Pulando.")
            continue

        # 4) espere carregar o painel e colete o conteúdo
        try:
            page.wait_for_load_state("domcontentloaded", timeout=8000)
        except Exception:
            pass

        historico = detalhar_mensagem(page)

        # Se assunto/tipo vieram vazios do bloco, tente extrair do painel
        if not assunto and historico and len(historico.splitlines()) > 0:
            assunto = historico.splitlines()[0][:200].strip()

        # 5) escreva no CSV
        escrever_linha_msg(cnpj, empresa, tipo, assunto, historico, data)
        # voltar_para_servicos(page)

def voltar_para_servicos(page):
    page.goto("https://det.sit.trabalho.gov.br/servicos")

def criar_pasta_do_mes_e_salvar_print(page, empresa, cnpj):
    caminho_pasta_mes = os.path.join(LOG_DIR, mes_ano)
    
    os.makedirs(caminho_pasta_mes, exist_ok=True)
    
    nome_print = f"{empresa.replace('/', '')} - {cnpj.replace('/', '').replace('.', '').replace('-', '')}.png"
    caminho_print = os.path.join(caminho_pasta_mes, nome_print)
    
    try:
        time.sleep(1)
        page.screenshot(path=caminho_print)
        print(f"Print da tela salvo em: {caminho_print}")
    except Exception as e:
        print(f"Erro ao salvar o print: {e}")

PADRAO_DATA = re.compile(
    r"("                                     # match completo
    r"\b\d{1,2}\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\.?\s+\d{2,4}\b"  # 21 jul 25 / 21 jul. 2025
    r"|\b\d{1,2}/\d{1,2}/\d{2,4}\b"          # 21/07/2025
    r"|\bhoje\b(?:\s*(?:às|as)?\s*\d{1,2}:\d{2})?"  # hoje, hoje 09:12, hoje às 09:12
    r")",
    re.I
)

def normalizar_data(data_raw: str) -> str:
    if not data_raw:
        return ""
    s = data_raw.strip()

    # qualquer ocorrência de "hoje" (com/sem horário)
    if re.search(r"(?i)\bhoje\b", s):
        return datetime.now().strftime("%d/%m/%Y")

    # dd/mm/aa(aa)
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", s)
    if m:
        d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000  # assume século 2000
        return f"{d:02d}/{mth:02d}/{y:04d}"

    # dd mmm aa(aa) em PT-BR (ex: 21 jul 25 / 21 fev 2025)
    m = re.match(r"^\s*(\d{1,2})\s+([A-Za-zçãéáíóúûêô]{3,})\s+(\d{2,4})\s*$", s, re.I)
    if m:
        d = int(m.group(1))
        mon = m.group(2).lower()[:3]
        mapa = {'jan':1,'fev':2,'mar':3,'abr':4,'mai':5,'jun':6,'jul':7,'ago':8,'set':9,'out':10,'nov':11,'dez':12}
        if mon in mapa:
            y = int(m.group(3))
            if y < 100:
                y += 2000
            return f"{d:02d}/{mapa[mon]:02d}/{y:04d}"

    return s

def extrair_msg_do_item(item_locator):
    try:
        titulo = item_locator.locator(".titulo").inner_text(timeout=2500).strip()
    except Exception:
        titulo = ""

    try:
        remetente = item_locator.locator(".origem").inner_text(timeout=2500).strip()
    except Exception:
        remetente = ""

    try:
        bruto = item_locator.inner_text(timeout=2500)
    except Exception:
        bruto = ""

    m = PADRAO_DATA.search(bruto)
    if m:
        data = normalizar_data(m.group(0))   # <<<<< use o match completo
    else:
        # Fallback: se o texto cita "hoje" mas o regex não capturou, force a data de hoje
        data = datetime.now().strftime("%d/%m/%Y") if re.search(r"(?i)\bhoje\b", bruto) else ""

    return remetente, titulo, data

def processar_empresa(page, cnpj, empresa):
    print(f"\nProcessando empresa: {empresa} (CNPJ: {cnpj})")
    
    # entrar sem procuração para CNPJ do escritorio
    if cnpj == "07.053.914/0001-60":
        # Limpa cookies do contexto
        context = page.context
        context.clear_cookies()

        # Limpa local/session storage da aba
        page.evaluate("() => { localStorage.clear(); sessionStorage.clear(); }")

        # Agora sim navega de volta
        page.goto("https://det.sit.trabalho.gov.br/servicos", wait_until="domcontentloaded")
        time.sleep(1)
        page.click("text=Entrar com")
        time.sleep(1)
        page.click("text=Seu certificado digital")
            
        abrir_caixa_postal(page)

        # 2. Salvar print da tela da caixa postal
        criar_pasta_do_mes_e_salvar_print(page, empresa, cnpj)
        
        # 3. Processar mensagens
        try:
            garantir_log()

            # localiza os itens da lista da Caixa Postal (coluna esquerda)
            itens = localizar_lista_itens(page)
            total = itens.count()
            print(f"Encontradas {total} mensagens para esta empresa.")

            if total == 0:
                escrever_linha_msg(cnpj, empresa, "", "Sem mensagens", "")
                # voltar para serviços após processar essa empresa
                voltar_para_servicos(page)
                return

            for i in range(total):
                item = itens.nth(i)
                try:
                    item.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass

                remetente, titulo, data = extrair_msg_do_item(item)

                escrever_linha_msg(
                    (cnpj or "").strip(),
                    (empresa or "").strip(),
                    (remetente or "").strip(),
                    (titulo or "").strip(),
                    (data or "").strip()
                )

            # voltar para serviços após processar essa empresa
            voltar_para_servicos(page)

        except Exception as e:
            print(f"Ocorreu um erro ao processar as mensagens da empresa {empresa}: {e}")

    dialog = abrir_trocar_perfil(page)
    selecionar_perfil_procurador(page, dialog)
    preencher_cnpj_e_selecionar(dialog, cnpj)

    # busca aviso de sem procuração
    try:
        time.sleep(1.5)
        aviso = page.locator("app-modal-perfil")
        if aviso.is_visible(timeout=500):
            print(f"Aviso detectado: Sem procuração")

            from pathlib import Path
            pasta_base = Path("prints")
            pasta_sem_proc = pasta_base / "SEM PROCURACAO"
            pasta_sem_proc.mkdir(parents=True, exist_ok=True)
            nome_arquivo = f"{empresa} - {cnpj.replace('/', '').replace('.', '').replace('-', '')}.png"
            page.screenshot(path=str(pasta_sem_proc / nome_arquivo))

            garantir_log()
            escrever_linha_msg(cnpj, empresa, "", "NÃO TEM PROCURAÇÃO", "")

            # volta pra tela de serviços
            page.goto("https://det.sit.trabalho.gov.br/servicos", wait_until="domcontentloaded")

            return  # sai da função para seguir para a próxima empresa
    except Exception:
        pass

    # tenta detectar se caiu em /cadastro
    caiu_em_cadastro = False
    try:
        page.wait_for_url("**/cadastro", timeout=1000)
        caiu_em_cadastro = True
    except:
        pass

    if caiu_em_cadastro:
        from pathlib import Path
        print(f"Empresa {empresa} (CNPJ: {cnpj}) não possui procuração!")
        pasta_base = Path("prints")
        pasta_sem_proc = pasta_base / "SEM PROCURACAO"
        pasta_sem_proc.mkdir(parents=True, exist_ok=True)
        nome_arquivo = f"{empresa} - {cnpj}.png".replace("/", "-")
        page.screenshot(path=str(pasta_sem_proc / nome_arquivo))
        garantir_log()
        escrever_linha_msg(cnpj, empresa, "", "NÃO TEM PROCURAÇÃO", "")

        # fecha a aba travada e abre uma nova em /servicos
        # Limpa cookies do contexto
        context = page.context
        context.clear_cookies()

        # Limpa local/session storage da aba
        page.evaluate("() => { localStorage.clear(); sessionStorage.clear(); }")

        # Agora sim navega de volta
        page.goto("https://det.sit.trabalho.gov.br/servicos", wait_until="domcontentloaded")
        time.sleep(1)
        page.click("text=Entrar com")
        time.sleep(1)
        page.click("text=Seu certificado digital")
        
        # devolve a nova aba para o loop principal (se precisar usar depois)
        return

    abrir_caixa_postal(page)

    # 2. Salvar print da tela da caixa postal
    criar_pasta_do_mes_e_salvar_print(page, empresa, cnpj)
    
    # 3. Processar mensagens
    try:
        garantir_log()

        # localiza os itens da lista da Caixa Postal (coluna esquerda)
        itens = localizar_lista_itens(page)
        total = itens.count()
        print(f"Encontradas {total} mensagens para esta empresa.")

        if total == 0:
            escrever_linha_msg(cnpj, empresa, "", "Sem mensagens", "")
            # voltar para serviços após processar essa empresa
            voltar_para_servicos(page)
            return

        for i in range(total):
            item = itens.nth(i)
            try:
                item.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass

            remetente, titulo, data = extrair_msg_do_item(item)

            escrever_linha_msg(
                (cnpj or "").strip(),
                (empresa or "").strip(),
                (remetente or "").strip(),
                (titulo or "").strip(),
                (data or "").strip()
            )

        # voltar para serviços após processar essa empresa
        voltar_para_servicos(page)

    except Exception as e:
        print(f"Ocorreu um erro ao processar as mensagens da empresa {empresa}: {e}")

if __name__ == "__main__":
    with sync_playwright() as p:
        browser = chrome_9222(p, PORT)   # conecta ou inicia Chrome
        context = browser.contexts[0]    # contexto persistente
        page = context.new_page()
        
        page.goto("https://det.sit.trabalho.gov.br/login?r=%2Fservicos")
        page.wait_for_load_state("domcontentloaded")
        page.click("text=Entrar com")
        page.wait_for_url("https://det.sit.trabalho.gov.br/servicos")

        empresas = ler_lista_empresas(CSV_PATH)
        for empresa_info in empresas:
            cnpj = empresa_info["cnpj"]
            razao_social = empresa_info["empresa"]
            processar_empresa(page, cnpj, razao_social)