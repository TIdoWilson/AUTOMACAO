import os
import re
import unicodedata
from collections import defaultdict
from datetime import datetime, timedelta

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# =====================================================
# CONFIGURAÇÕES GERAIS
# =====================================================

# Credenciais Tareffa (usuário de robôs)
TAREFFA_EMAIL = "robos@wilsonlopes.com.br"   # e-mail usado para login no Tareffa
TAREFFA_SENHA = "Wlopes.2025$"               # senha usada para login no Tareffa

# Nomes de arquivos usados pelo robô (sem caminho)
# Todos serão salvos na MESMA PASTA do script.
NOME_ARQUIVO_CSV = "Exportação de Serviços.csv"
NOME_ARQUIVO_SAIDA = "tareffa_inconsistencias_auto_sem_balao_v9.3.xlsx"
NOME_ARQUIVO_CONTROLE = "Serviços que precisam ser checados.xlsx"
NOME_ARQUIVO_SERVICOS_INCONS = "Serviços que precisam de atenção.xlsx"
NOME_ARQUIVO_DIAS_PROCESSADOS = "Dias conferidos com data e hora.xlsx"

# Ícone do balão azul (comunicação criada)
SELETOR_BALAO_AZUL = "button svg[data-icon='comment-dots']"

# Parâmetros de scroll/coleta na lista
SCROLL_PAUSE_MS = 400
COLETA_INTERVALO = 5
CHECAR_FIM_INTERVALO = 5
REACHED_LIMIT = 20
MAX_LOOPS_SCROLL = 400

# Modo detalhado de prints no terminal (não afeta o log em arquivo)
DEBUG_MODE = True
MAX_DEBUG_LINHAS = 15  # quantas linhas por departamento/dia mostrar no terminal

# Se False, NÃO salva no Excel as linhas Bx.Autom sem balão que
# não casarem com nenhum serviço do CSV (Servico_casado_com_CSV=False)
SALVAR_LINHAS_SEM_VINCULO = False

# Pasta base = mesma pasta do script .py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Caminho do arquivo de log (sempre na mesma pasta do script)
LOG_ARQUIVO = os.path.join(BASE_DIR, "tareffa_log.txt")


# =====================================================
# LOG
# =====================================================

def debug(msg: str):
    """Print no terminal só se DEBUG_MODE estiver ativo."""
    if DEBUG_MODE:
        print(msg)


def escrever_log(msg: str):
    """
    Registra uma linha de log de forma leve.
    Formato: 2025-12-02T08:19:48|TEXTO_COM_UNDERSCORES
    """
    try:
        ts = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        linha = f"{ts}|{msg.strip().replace(' ', '_')}"
        with open(LOG_ARQUIVO, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        # Nunca deixar o log quebrar o robô
        pass


# =====================================================
# FUNÇÕES AUXILIARES PARA TEXTO
# =====================================================

def limpar_cauda(texto: str) -> str:
    if not isinstance(texto, str):
        return ""
    texto = texto.rstrip()
    while texto:
        ch = texto[-1]
        cat = unicodedata.category(ch)
        if ch.isspace() or cat.startswith("Z") or cat == "Cf":
            texto = texto[:-1]
        else:
            break
    return texto


def normalizar_espacos_internos(texto: str) -> str:
    texto = limpar_cauda(texto)
    return " ".join(texto.split())


def normalizar_nome(texto: str) -> str:
    texto = limpar_cauda(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return texto.upper()


# =====================================================
# CARREGAMENTO DOS SERVIÇOS (CSV / CONTROLE)
# =====================================================

def carregar_servicos_lista(caminho_csv: str, caminho_controle: str):
    """
    Lê o CSV exportado do Tareffa e/ou o arquivo de controle (Excel).
    Retorna:
      - lista de serviços *ativos para checagem* (dicts com id, nome, nome_norm, departamento)
      - DataFrame de controle (df_controle) completo, com todos os serviços.

    Serviços que já têm inconsistência registrada (Status contendo 'COM INCONSISTENCIA'
    ou Qtd_inconsistencias > 0) NÃO entram mais na lista de checagem para rodadas futuras.
    """

    def criar_controle_a_partir_do_csv(motivo: str):
        """Gera df_controle novo a partir do CSV quando o Excel não existe ou está corrompido."""
        print(motivo)
        if not os.path.exists(caminho_csv):
            raise FileNotFoundError(
                f"{motivo}\n"
                f"Mas o CSV base também não foi encontrado.\n"
                f"CSV esperado: {caminho_csv}"
            )

        print(f"Gerando arquivo de controle a partir do CSV: {caminho_csv}")

        df = pd.read_csv(
            caminho_csv,
            sep=";",
            encoding="latin-1",
            dtype=str,
        )

        df["Nome"] = df["Nome"].fillna("").astype(str)
        df["Departamento"] = df["Departamento"].fillna("").astype(str)
        df["Comunica cliente"] = df["Comunica cliente"].fillna("").astype(str)

        df["comunica_bool"] = df["Comunica cliente"].str.strip().str.upper().isin(["SIM", "S"])
        df["nome_norm"] = df["Nome"].apply(normalizar_nome)

        df_sim = df[df["comunica_bool"]].copy()

        if df_sim.empty:
            print("Nenhum serviço com 'Comunica cliente = SIM' encontrado no CSV.")
            df_ctrl_local = pd.DataFrame(columns=[
                "ID", "Nome", "nome_norm", "Departamento",
                "Status", "Qtd_inconsistencias", "Ultima_verificacao",
                "Exemplo_linha", "Periodo_inicio", "Periodo_fim"
            ])
        else:
            df_ctrl_local = df_sim[["Nome", "nome_norm", "Departamento"]].copy()
            df_ctrl_local.insert(0, "ID", range(1, len(df_ctrl_local) + 1))
            df_ctrl_local["Status"] = ""
            df_ctrl_local["Qtd_inconsistencias"] = ""
            df_ctrl_local["Ultima_verificacao"] = ""
            df_ctrl_local["Exemplo_linha"] = ""
            df_ctrl_local["Periodo_inicio"] = ""
            df_ctrl_local["Periodo_fim"] = ""
            df_ctrl_local["ID"] = df_ctrl_local["ID"].astype(str)

            print(
                f"Total no CSV: {len(df)} linhas, "
                f"{len(df_sim)} com 'Comunica cliente = SIM'."
            )

        # Salva controle bem-formatado
        df_ctrl_local.to_excel(caminho_controle, index=False)
        print(f"Arquivo de controle criado/recuperado: {caminho_controle}")
        return df_ctrl_local

    if os.path.exists(caminho_controle):
        print(f"Usando arquivo de controle existente: {caminho_controle}")
        try:
            df_ctrl = pd.read_excel(caminho_controle, dtype=str)
        except Exception as e:
            # Arquivo não é um xlsx válido (ex: renomeado, corrompido, etc.)
            print("\nATENÇÃO: não consegui ler o arquivo de controle como Excel válido.")
            print("Motivo técnico:", e)
            escrever_log("ARQUIVO_CONTROLE_CORROMPIDO_OU_INVALIDO")

            # Tenta renomear o arquivo ruim para não perder totalmente
            try:
                base, ext = os.path.splitext(caminho_controle)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup = f"{base}_CORROMPIDO_{ts}{ext or '.bak'}"
                os.rename(caminho_controle, backup)
                print(f"Arquivo de controle antigo foi renomeado para: {backup}")
            except Exception as e_ren:
                print("Não consegui renomear o arquivo de controle antigo:", e_ren)

            # Recria a partir do CSV
            df_ctrl = criar_controle_a_partir_do_csv(
                "Arquivo de controle existente estava corrompido/inválido."
            )
    else:
        # Não existe ainda -> cria do zero
        df_ctrl = criar_controle_a_partir_do_csv(
            "Arquivo de controle não encontrado."
        )

    # Se por algum motivo vier vazio aqui, apenas devolve
    if df_ctrl.empty:
        print("Arquivo de controle está vazio; nenhum serviço cadastrado.")
        return [], df_ctrl

    # Garante colunas mínimas
    colunas_min = [
        "ID",
        "Nome",
        "nome_norm",
        "Departamento",
        "Status",
        "Qtd_inconsistencias",
        "Ultima_verificacao",
        "Exemplo_linha",
    ]
    for col in colunas_min:
        if col not in df_ctrl.columns:
            df_ctrl[col] = ""

    for col in ["Periodo_inicio", "Periodo_fim"]:
        if col not in df_ctrl.columns:
            df_ctrl[col] = ""

    if "ID" not in df_ctrl.columns:
        df_ctrl.insert(0, "ID", range(1, len(df_ctrl) + 1))
    df_ctrl["ID"] = df_ctrl["ID"].astype(str)

    # Normalizações finais
    for col in ["Nome", "Departamento"]:
        df_ctrl[col] = df_ctrl[col].fillna("").astype(str).apply(limpar_cauda)
    df_ctrl["nome_norm"] = df_ctrl["Nome"].apply(normalizar_nome)

    # -------------------------------------------------
    # NOVO: só entram na lista de checagem serviços que
    # ainda NÃO têm inconsistência registrada.
    # -------------------------------------------------
    df_ctrl["Status"] = df_ctrl["Status"].fillna("").astype(str)
    df_ctrl["Qtd_inconsistencias"] = df_ctrl["Qtd_inconsistencias"].fillna("").astype(str)

    mask_status_incons = df_ctrl["Status"].str.upper().str.contains("COM INCONSISTENCIA")
    qtd_num = pd.to_numeric(df_ctrl["Qtd_inconsistencias"], errors="coerce").fillna(0)
    mask_qtd_pos = qtd_num > 0

    # Pendentes = não COM INCONSISTENCIA e sem qtd > 0
    mask_pendente = ~(mask_status_incons | mask_qtd_pos)

    df_iter = df_ctrl[mask_pendente].copy()

    if df_iter.empty:
        print(
            "Todos os serviços do controle já têm inconsistência registrada "
            "(Status='COM INCONSISTENCIA' ou Qtd_inconsistencias>0). "
            "Nenhum serviço resta para checagem nesta rodada."
        )
    else:
        print(
            f"Total de serviços cadastrados no controle: {len(df_ctrl)} "
            f"(ativos para checagem nesta rodada: {len(df_iter)})"
        )

    # Monta lista de serviços ativos para checagem
    servicos = []
    for _, row in df_iter.iterrows():
        servicos.append(
            {
                "id": str(row["ID"]),
                "nome": row["Nome"],
                "nome_norm": row["nome_norm"],
                "departamento": limpar_cauda(row["Departamento"]) or "SEM DEPARTAMENTO",
            }
        )

    return servicos, df_ctrl


# =====================================================
# DIAS JÁ PROCESSADOS
# =====================================================

def carregar_dias_processados(caminho_dias: str):
    if not os.path.exists(caminho_dias):
        df_dias = pd.DataFrame(columns=["Data", "Ultima_execucao"])
        return set(), df_dias

    df_dias = pd.read_excel(caminho_dias, dtype=str)
    if "Data" not in df_dias.columns:
        df_dias["Data"] = ""
    if "Ultima_execucao" not in df_dias.columns:
        df_dias["Ultima_execucao"] = ""

    df_dias["Data"] = df_dias["Data"].fillna("").astype(str).str.strip()
    dias_set = set(df_dias["Data"].tolist())
    return dias_set, df_dias


def salvar_dias_processados(df_dias: pd.DataFrame, caminho_dias: str):
    df_dias.to_excel(caminho_dias, index=False)
    print(f"Dias já processados salvos/atualizados: {caminho_dias}")


# =====================================================
# LOGIN + PERFIL ROBOS
# =====================================================

def escolher_perfil_robos(page) -> bool:
    print("Tentando selecionar a conta/perfil 'ROBOS'...")

    try:
        loc_text = page.get_by_text("ROBOS", exact=False)
        qtd = loc_text.count()
        print(f"Elementos com texto 'ROBOS' encontrados: {qtd}")

        for idx in range(qtd):
            el = loc_text.nth(idx)
            card = el.locator(
                "xpath=ancestor::*[self::form or self::button or self::div or self::mat-card][1]"
            )
            if card.count() == 0:
                card = el

            overlay = card.locator(
                'button[type="submit"][style*="z-index: 10"]'
            )
            if overlay.count() == 0:
                overlay = card.locator("button[type='submit']")

            if overlay.count() > 0:
                print(f"Clicando botão overlay de ROBOS (card {idx})...")
                overlay.first.click()
                page.wait_for_timeout(2000)
                return True

            print(f"Clicando card com texto ROBOS (card {idx})...")
            card.first.click()
            page.wait_for_timeout(2000)
            return True

    except Exception as e:
        print("Erro ao tentar clicar baseado em texto 'ROBOS':", e)
        escrever_log("ERRO_AO_TENTAR_CLICAR_BASEADO_EM_TEXTO_ROBOS")

    try:
        overlay_all = page.locator(
            'button[type="submit"][style*="z-index: 10"]'
        )
        qtd_btn = overlay_all.count()
        print(f"Botões overlay encontrados pelo CSS: {qtd_btn}")
        if qtd_btn > 0:
            print("Fallback: clicando primeiro botão overlay...")
            overlay_all.first.click()
            page.wait_for_timeout(2000)
            return True
    except Exception as e:
        print("Erro no fallback clicando overlay:", e)
        escrever_log("ERRO_NO_FALLBACK_ROBOS")

    print("Não consegui selecionar a conta/perfil 'ROBOS'.")
    return False


def tentar_login_tareffa(page, max_tentativas: int = 3) -> bool:
    for tentativa in range(1, max_tentativas + 1):
        print(f"Tentativa de login {tentativa}/{max_tentativas}...")

        if page.locator("#username").count() == 0:
            print("Campo de usuário não está mais presente, assumindo que já logou.")
            return True

        try:
            user_input = page.locator("#username").first
            pass_input = page.locator("#password").first
        except Exception as e:
            print("Erro ao localizar campos de login:", e)
            escrever_log("ERRO_LOCALIZAR_CAMPOS_LOGIN")
            return False

        try:
            user_input.fill(TAREFFA_EMAIL)
            pass_input.fill(TAREFFA_SENHA)
        except Exception as e:
            print("Erro ao preencher usuário/senha:", e)
            escrever_log("ERRO_PREENCHER_LOGIN")
            return False

        try:
            botao_login = page.locator(
                'button[type="submit"][style*="z-index: 10"]'
            )
            if botao_login.count() == 0:
                botao_login = page.locator("button[type='submit']")
            if botao_login.count() > 0:
                botao_login.first.click()
            else:
                page.get_by_role("button").first.click()
        except Exception as e:
            print("Erro ao clicar no botão de login:", e)
            escrever_log("ERRO_CLIQUE_BOTAO_LOGIN")

        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeoutError:
            print("Timeout em networkidle após login (pode ser normal).")
        except Exception as e:
            print("Erro ao aguardar carregamento após login:", e)

        page.wait_for_timeout(2000)
        url_atual = page.url
        print(f"URL após tentativa de login: {url_atual}")

        if "servicos_programados" in url_atual:
            print("Login concluído: já estamos em /servicos_programados.")
            return True

        if "oauthchooseaccount" in url_atual or "ottimizza-oauth-server" in url_atual:
            print("Login efetuado: agora falta escolher a conta/perfil.")
            return True

        if page.locator("#username").count() > 0:
            print("Ainda na tela de login, tentando novamente se houver tentativas restantes.")
            continue

        print("Campo de login sumiu, assumindo que a autenticação foi concluída.")
        return True

    print("Não consegui logar após várias tentativas.")
    escrever_log("LOGIN_FALHOU_VARIAS_TENTATIVAS")
    return False


def acessar_servicos_programados(page):
    url_servicos = "https://web.tareffa.com.br/servicos_programados"

    for tentativa in range(15):
        print(f"Tentativa {tentativa+1}/15 para acessar /servicos_programados...")

        try:
            page.goto(url_servicos, wait_until="domcontentloaded")
        except Exception as e:
            print(f"Erro em page.goto({url_servicos}): {e}")

        page.wait_for_timeout(1000)
        url_atual = page.url
        print("URL atual:", url_atual)

        if "web.tareffa.com.br/servicos_programados" in url_atual:
            print("Estamos em /servicos_programados.")
            return

        if page.locator("#username").count() > 0:
            print("Tela de login detectada, iniciando rotina de login com retries...")
            tentar_login_tareffa(page)
            page.wait_for_timeout(1500)
            continue

        if "oauthchooseaccount" in url_atual or "ottimizza-oauth-server" in url_atual:
            print("Tela de escolha de conta/perfil detectada.")
            escolher_perfil_robos(page)
            page.wait_for_timeout(1500)
            continue

        page.wait_for_timeout(1000)

    print("Não consegui garantir /servicos_programados após várias tentativas.")
    escrever_log("NAO_CONSEGUI_ACESSAR_SERVICOS_PROGRAMADOS")


def garantir_pagina_servicos(page):
    """
    Garante que estamos na tela /servicos_programados.
    Se não estiver, tenta voltar pra ela.
    """
    try:
        if "servicos_programados" not in page.url:
            print("URL atual não é /servicos_programados; tentando voltar para a tela de serviços...")
            escrever_log("URL_FORA_DE_SERVICOS_PROGRAMADOS_TENTANDO_RECUPERAR")
            acessar_servicos_programados(page)
    except Exception as e:
        print("Erro ao garantir página de serviços:", e)
        escrever_log("ERRO_GARANTIR_PAGINA_SERVICOS")


# =====================================================
# FILTRO DE DATAS
# =====================================================

def interpretar_data_usuario(texto: str):
    txt = texto.strip().replace(" ", "")
    if not txt:
        raise ValueError("Data vazia")

    try:
        return datetime.strptime(txt, "%d/%m/%Y").date()
    except ValueError:
        pass

    if len(txt) == 8 and txt.isdigit():
        dia = int(txt[:2])
        mes = int(txt[2:4])
        ano = int(txt[4:])
        return datetime(ano, mes, dia).date()

    raise ValueError("Formato de data inválido")


def aplicar_filtro_datas(page, data_inicio: str, data_fim: str):
    print(f"Aplicando filtro de Data Programada: Início={data_inicio}, Fim={data_fim}")

    try:
        campo_ini = page.locator("#mat-input-3")
        campo_fim = page.locator("#mat-input-4")

        if campo_ini.count() == 0 or campo_fim.count() == 0:
            print("Campos de data não encontrados (#mat-input-3 / #mat-input-4). Seguindo sem filtro de data.")
            escrever_log("CAMPOS_DATA_NAO_ENCONTRADOS")
            return

        campo_ini = campo_ini.first
        campo_fim = campo_fim.first

        campo_ini.click()
        campo_ini.fill("")
        campo_ini.fill(data_inicio)
        campo_ini.press("Enter")
        page.wait_for_timeout(500)

        campo_fim.click()
        campo_fim.fill("")
        campo_fim.fill(data_fim)
        campo_fim.press("Enter")
        page.wait_for_timeout(1500)

        print("Filtro de datas aplicado.")
    except Exception as e:
        print("Erro ao aplicar filtro de datas:", e)
        escrever_log("ERRO_APLICAR_FILTRO_DATAS")


# =====================================================
# STATUS
# =====================================================

def selecionar_status_bx_autom(page):
    print("Tentando aplicar filtro de Status = 'Bx. Autom'...")

    try:
        trigger = page.locator("#mat-select-value-1")
        if trigger.count() == 0:
            trigger = page.get_by_text("Status", exact=False).locator(
                "xpath=ancestor::*[1]//*[contains(@class,'mat-mdc-select-trigger')]"
            )

        if trigger.count() == 0:
            print("Não encontrei trigger de Status; seguindo sem filtro.")
            escrever_log("NAO_ENCONTREI_TRIGGER_STATUS")
            return

        trigger.first.click()
        page.wait_for_timeout(400)
        print("Dropdown de Status aberto.")
    except Exception as e:
        print("Não consegui abrir o dropdown de Status; erro:", e)
        escrever_log("ERRO_ABRIR_DROPDOWN_STATUS")
        return

    selecionado = False
    candidatos_opcao = [
        page.locator("span.mdc-list-item__primary-text:has-text('Bx. Autom')"),
        page.get_by_text("Bx. Autom", exact=False),
    ]

    for idx, loc in enumerate(candidatos_opcao, start=1):
        try:
            if loc.count() == 0:
                continue
            print(f"Tentando clicar na opção 'Bx. Autom' com candidato {idx}...")
            loc.first.click()
            selecionado = True
            print("Opção 'Bx. Autom' clicada.")
            break
        except Exception as e:
            print(f"Erro ao clicar na opção com candidato {idx}: {e}")

    if not selecionado:
        print("Não consegui selecionar 'Bx. Autom'; filtro ficará em 'Todos'.")
        escrever_log("NAO_CONSEGUI_SELECIONAR_BX_AUTOM")
        return

    page.wait_for_timeout(1500)
    print("Filtro de Status ajustado para 'Bx. Autom'.")


# =====================================================
# FILTRO AVANÇADO (Departamento)
# =====================================================

def abrir_filtro_avancado(page) -> bool:
    print("Tentando abrir Filtro Avançado...")

    candidatos = []

    try:
        loc = page.locator("button[mattooltip*='Filtro'][mattooltip*='avanç']")
        if loc.count() > 0:
            candidatos.append(("button[mattooltip*='Filtro*avanç']", loc))
    except Exception:
        pass

    try:
        loc = page.locator("button[ng-reflect-message*='Filtros avançados']")
        if loc.count() > 0:
            candidatos.append(("button[ng-reflect-message*='Filtros avançados']", loc))
    except Exception:
        pass

    try:
        loc = page.locator("button[aria-label*='Filtros avançados']")
        if loc.count() > 0:
            candidatos.append(("button[aria-label*='Filtros avançados']", loc))
    except Exception:
        pass

    try:
        loc = page.get_by_role("button", name="Filtros avançados", exact=False)
        if loc.count() > 0:
            candidatos.append(("role=button name~'Filtros avançados'", loc))
    except Exception:
        pass

    try:
        loc = page.get_by_role("button", name="Filtro avançado", exact=False)
        if loc.count() > 0:
            candidatos.append(("role=button name~'Filtro avançado'", loc))
    except Exception:
        pass

    if not candidatos:
        print("Não encontrei nenhum botão de 'Filtro(s) avançado(s)'; seguirei sem ele.")
        escrever_log("BOTAO_FILTRO_AVANCADO_NAO_ENCONTRADO")
        return False

    for idx, (descr, loc) in enumerate(candidatos, start=1):
        try:
            print(f"Clicando no filtro avançado ({descr}, candidato {idx})...")
            loc.first.click()
            page.wait_for_timeout(800)
            return True
        except Exception as e:
            print(f"Erro ao clicar candidato {idx} ({descr}): {e}")

    print("Não consegui abrir o Filtro Avançado com nenhum botão candidato.")
    escrever_log("ERRO_ABRIR_FILTRO_AVANCADO")
    return False


def aplicar_filtro_avancado_departamento(page, departamento: str):
    depto_txt = normalizar_espacos_internos(departamento or "")
    if not depto_txt or depto_txt == "SEM DEPARTAMENTO":
        print("Departamento 'SEM DEPARTAMENTO'; não aplicarei filtro avançado.")
        return

    if not abrir_filtro_avancado(page):
        return

    print(f"Aplicando Filtro Avançado para Departamento = {depto_txt!r}...")

    campo_depto = None

    try:
        loc = page.locator("#mat-input-7")
        if loc.count() > 0:
            campo_depto = loc.first
            print("Campo 'Departamento' encontrado via id #mat-input-7.")
    except Exception:
        pass

    if campo_depto is None:
        try:
            campo = page.locator("mat-form-field:has-text('Departamento') input")
            if campo.count() > 0:
                campo_depto = campo.first
                print("Campo 'Departamento' encontrado via mat-form-field.")
        except Exception:
            pass

    if campo_depto is None:
        try:
            label = page.get_by_text("Departamento", exact=False).first
            wrapper = label.locator(
                "xpath=ancestor::*[self::mat-form-field or self::div][1]"
            )
            campo = wrapper.locator("input")
            if campo.count() > 0:
                campo_depto = campo.first
                print("Campo 'Departamento' encontrado próximo ao texto.")
        except Exception:
            pass

    if campo_depto is None:
        print("Campo 'Departamento' dentro do Filtro Avançado não encontrado.")
    else:
        campo_depto.click()
        campo_depto.fill(depto_txt)
        page.wait_for_timeout(600)
        campo_depto.press("Enter")

    aplicou = False
    candidatos_botoes = [
        page.get_by_role("button", name="Aplicar", exact=False),
        page.get_by_role("button", name="Filtrar", exact=False),
        page.get_by_role("button", name="Aplicar filtro", exact=False),
        page.get_by_text("Aplicar", exact=False),
    ]
    for idx, loc in enumerate(candidatos_botoes, start=1):
        try:
            if loc.count() > 0:
                print(f"Clicando botão de aplicar (candidato {idx})...")
                loc.first.click()
                aplicou = True
                break
        except Exception as e:
            print(f"Erro ao clicar botão de aplicar {idx}: {e}")

    if not aplicou:
        print("Não consegui clicar em 'Aplicar' no Filtro Avançado.")
        escrever_log("ERRO_APLICAR_FILTRO_AVANCADO")

    page.wait_for_timeout(1500)


def limpar_filtro_avancado_departamento(page):
    try:
        print("Limpando filtro avançado de Departamento...")
        if not abrir_filtro_avancado(page):
            print("Não foi possível abrir o Filtro Avançado para limpar.")
            escrever_log("ERRO_ABRIR_FILTRO_AVANCADO_PARA_LIMPAR")
            return

        campo_depto = None

        try:
            loc = page.locator("#mat-input-7")
            if loc.count() > 0:
                campo_depto = loc.first
                print("Campo 'Departamento' (limpeza) via #mat-input-7.")
        except Exception:
            pass

        if campo_depto is None:
            try:
                loc = page.locator("mat-form-field:has-text('Departamento') input")
                if loc.count() > 0:
                    campo_depto = loc.first
                    print("Campo 'Departamento' (limpeza) via mat-form-field.")
            except Exception:
                pass

        if campo_depto is not None:
            campo_depto.click()
            campo_depto.fill("")
            campo_depto.press("Enter")
        else:
            print("Campo 'Departamento' não encontrado para limpeza.")

        aplicou = False
        candidatos_botoes = [
            page.get_by_role("button", name="Aplicar", exact=False),
            page.get_by_role("button", name="Filtrar", exact=False),
            page.get_by_role("button", name="Aplicar filtro", exact=False),
            page.get_by_text("Aplicar", exact=False),
        ]
        for idx, loc in enumerate(candidatos_botoes, start=1):
            try:
                if loc.count() > 0:
                    print(f"Clicando botão de aplicar (limpeza, candidato {idx})...")
                    loc.first.click()
                    aplicou = True
                    break
            except Exception as e:
                print(f"Erro ao clicar botão de aplicar (limpeza) {idx}: {e}")

        if not aplicou:
            print("Não consegui clicar em 'Aplicar' ao limpar Filtro Avançado.")
            escrever_log("ERRO_APLICAR_FILTRO_AVANCADO_LIMPEZA")

        page.wait_for_timeout(800)

    except Exception as e:
        print("Erro inesperado ao limpar Filtro Avançado:", e)
        escrever_log("ERRO_LIMPEZA_FILTRO_AVANCADO_EXCEPTION")


# =====================================================
# FIM DA LISTA (mensagens)
# =====================================================

def chegou_ao_fim_da_lista(page) -> bool:
    padroes_texto = [
        "Não foram encontrados serviços programados",
        "Nao foram encontrados servicos programados",
        "Você chegou ao fim",
        "voce chegou ao fim",
        "fim da lista",
        "fim da fila",
        "fim dos resultados",
        "não há mais registros",
        "nao ha mais registros",
        "nenhum serviço encontrado",
        "nenhum servico encontrado",
        "nenhum serviço programado",
        "nenhum servico programado",
        "nenhum registro encontrado",
        "nenhum resultado encontrado",
    ]

    try:
        for texto in padroes_texto:
            loc = page.get_by_text(texto, exact=False)
            if loc.count() > 0:
                print(f"Mensagem de fim detectada: {texto!r}")
                return True
    except Exception as e:
        print("Erro ao procurar mensagem de fim:", e)
        escrever_log("ERRO_PROCURAR_MENSAGEM_FIM")

    return False


# =====================================================
# SCROLL (lista interna)
# =====================================================

def _scroll_lista_uma_vez(page, subir_primeiro: bool = False):
    """
    Faz scroll especificamente no container da LISTA (scroll interno),
    nunca priorizando o scroll da página.

    subir_primeiro=True: primeiro sobe um pouco, depois desce (para
    forçar recarregamento da lista de tempos em tempos).
    """
    js = """
    (subirPrimeiro) => {
        const selectors = [
            '#card-content .cdk-virtual-scroll-viewport',
            '#card-content [cdk-virtual-scroll-viewport]',
            '.cdk-virtual-scroll-viewport',
            'div[cdk-virtual-scroll-viewport]',
            '#card-content'
        ];

        let target = null;
        let selectorUsed = null;
        let bestOverflow = -1;

        // Escolhe o container com MAIOR overflow (scrollHeight - clientHeight)
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (!el) continue;
            const ch = el.clientHeight || 0;
            const sh = el.scrollHeight || 0;
            const overflow = sh - ch;
            if (overflow > bestOverflow) {
                bestOverflow = overflow;
                target = el;
                selectorUsed = sel;
            }
        }

        if (!target) {
            target = document.scrollingElement || document.documentElement || document.body;
            selectorUsed = 'document';
        }

        if (!target) {
            return {js_ok:false, selector:null, prev:0, newv:0, max:0, reached:false, scrollable:false};
        }

        const clientH = target.clientHeight || window.innerHeight || 0;
        const scrollH = target.scrollHeight || 0;
        const max = scrollH > clientH ? (scrollH - clientH) : 0;

        // Se não tem overflow, não há scroll real
        let scrollable = max > 0;

        let prev = target.scrollTop || 0;

        // De vez em quando, sobe um pouco antes de descer
        if (subirPrimeiro && scrollable) {
            const upStep = clientH * 1.2 || 800;
            const novoTopo = Math.max(0, (target.scrollTop || 0) - upStep);
            target.scrollTop = novoTopo;
            prev = target.scrollTop || novoTopo;
        }

        if (!scrollable) {
            return {
                js_ok: true,
                selector: selectorUsed,
                prev: prev,
                newv: prev,
                max: max,
                reached: true,
                scrollable: false
            };
        }

        let step = clientH * 3;
        if (!step || step < 1500) step = 1500;

        let next = prev + step;
        if (next > max) next = max;

        target.scrollTop = next;
        const real = target.scrollTop || next;
        const reached = real >= max - 2;

        return {
            js_ok: true,
            selector: selectorUsed,
            prev: prev,
            newv: real,
            max: max,
            reached: reached,
            scrollable: true
        };
    }
    """

    try:
        info = page.evaluate(js, subir_primeiro)
    except Exception as e:
        print("Erro no JS de scroll:", e)
        escrever_log("ERRO_JS_SCROLL")
        info = {
            "js_ok": False,
            "selector": None,
            "prev": 0,
            "newv": 0,
            "max": 0,
            "reached": False,
            "scrollable": False,
        }

    selector = info.get("selector") or "desconhecido"
    prev = info.get("prev", 0) or 0
    newv = info.get("newv", info.get("new", 0)) or 0
    max_scroll = info.get("max", 0) or 0
    reached = bool(info.get("reached", False))
    scrollable = bool(info.get("scrollable", False))

    # Se é rolável, mas o topo não mudou, usa wheel no container interno
    if scrollable and newv == prev:
        try:
            if selector not in (None, "desconhecido", "document"):
                loc = page.locator(selector)
                if loc.count() > 0:
                    box = loc.first.bounding_box()
                    if box:
                        cx = box["x"] + box["width"] / 2
                        cy = box["y"] + box["height"] / 2
                        page.mouse.move(cx, cy)
                        page.mouse.wheel(0, box["height"] or 1500)
            else:
                vs = page.viewport_size
                cx = (vs["width"] / 2) if vs else 640
                cy = (vs["height"] / 2) if vs else 360
                page.mouse.move(cx, cy)
                page.mouse.wheel(0, 1500)
        except Exception as e:
            print("Fallback wheel falhou:", e)
            escrever_log("ERRO_FALLBACK_SCROLL")

    debug(
        f"scroll: seletor={selector}, prev={prev}, "
        f"novo={newv}, max={max_scroll}, reached={reached}, scrollable={scrollable}"
    )

    return {
        "selector": selector,
        "prev": prev,
        "new": newv,
        "max": max_scroll,
        "reached": reached,
        "scrollable": scrollable,
    }


# =====================================================
# PRÉ-SCROLL / COLETA (POR DEPARTAMENTO/DIA)
# =====================================================

def carregar_todos_registros_lista(page, max_loops: int = MAX_LOOPS_SCROLL):
    """
    Varre a lista inteira ANCORANDO em chips 'Bx. Autom' e ignorando linhas que tenham balão azul.
    É por DEPARTAMENTO/DIA (sem filtro de Serviço).
    Retorna:
      - total_chips_aprox (numero de linhas Bx.Autom sem balão)
      - linhas_inconsistentes (lista de dicts com texto/celulas)
      - fim_confirmado (True se foi possível garantir o fim da lista)
    """
    print("Pré-scroll: carregando TODOS os registros desse departamento/dia...")

    reached_seq = 0
    linhas_inconsistentes = []
    textos_vistos = set()
    fim_confirmado = False

    # Heurística de "sem novidade" para evitar softlock
    ultimo_total = 0
    loops_sem_novidade = 0
    LIMITE_SEM_NOVIDADE = 10

    # JS REFINADO: só linhas reais de serviço, com dígitos, dentro do container,
    # ignorando cabeçalho / menu / "Status\nBx. Autom" etc.
    js_coleta = """
    (selectorBalao) => {
        const xpathExpr =
          "//*[contains(normalize-space(.), 'Bx. Autom') or " +
          "contains(normalize-space(.), 'Bx.Autom') or " +
          "contains(normalize-space(.), 'BX. AUTOM')]";

        const containerSelectors = [
          '#card-content .cdk-virtual-scroll-viewport',
          '#card-content ul.timeline',
          '#card-content',
          '.cdk-virtual-scroll-viewport',
          'div[cdk-virtual-scroll-viewport]',
          'ul.timeline'
        ];

        let container = null;
        for (const sel of containerSelectors) {
            const el = document.querySelector(sel);
            if (el) {
                container = el;
                break;
            }
        }
        if (!container) {
            container = document.body;
        }

        const result = document.evaluate(
          xpathExpr,
          container,
          null,
          XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,
          null
        );

        function linhaDentroDoContainer(row) {
            if (!row) return false;
            if (
              !row.closest('#card-content') &&
              !row.closest('.cdk-virtual-scroll-viewport') &&
              !row.closest('div[cdk-virtual-scroll-viewport]') &&
              !row.closest('ul.timeline')
            ) {
              return false;
            }
            if (row.closest('thead')) return false;
            const cls = row.classList ? Array.from(row.classList).join(' ') : '';
            if (/mat-mdc-header-row|mat-header-row|mdc-data-table__header-row/i.test(cls)) return false;
            return true;
        }

        function escolherLinha(chip) {
            if (!(chip instanceof HTMLElement)) return null;

            let el = chip;
            let depth = 0;

            while (el && depth < 25) {
                try {
                    if (el.tagName === 'LI' && el.closest('ul.timeline')) {
                        if (linhaDentroDoContainer(el)) return el;
                    }

                    if (el.getAttribute && el.getAttribute('role') === 'row') {
                        if (linhaDentroDoContainer(el)) return el;
                    }

                    if (el.tagName === 'TR') {
                        if (linhaDentroDoContainer(el)) return el;
                    }

                    if (el.classList && el.classList.length) {
                        const cls = Array.from(el.classList).join(' ');
                        if (/(mat-mdc-row|mdc-data-table__row|mat-row)/i.test(cls)) {
                            if (linhaDentroDoContainer(el)) return el;
                        }
                    }
                } catch (e) {
                    // ignora e continua subindo
                }

                el = el.parentElement;
                depth++;
            }

            el = chip.closest('li') || chip.closest('tr, [role="row"], .mat-mdc-row, .mdc-data-table__row');
            if (el && linhaDentroDoContainer(el)) return el;

            return null;
        }

        const saida = [];
        const vistosLocal = new Set();

        for (let i = 0; i < result.snapshotLength; i++) {
            const chip = result.snapshotItem(i);
            if (!(chip instanceof HTMLElement)) continue;

            const row = escolherLinha(chip);
            if (!row) continue;

            if (selectorBalao && row.querySelector(selectorBalao)) {
                continue;
            }

            const text = (row.innerText || '').trim();
            if (!text) continue;

            const tl = text.toLowerCase();

            // PULA qualquer linha SEM dígito (sem data, sem 0/1, sem código etc.)
            const hasDigit = /[0-9]/.test(text);
            if (!hasDigit) continue;

            // PULA cabeçalhos lógicos e menus
            if (tl.includes('bom dia, robos') || tl.includes('bom dia, robôs')) continue;

            // Barra de filtros principal
            if (
              tl.includes('serviço') &&
              tl.includes('empresa') &&
              (tl.includes('responsavel') || tl.includes('responsável')) &&
              tl.includes('data programada início')
            ) {
              continue;
            }

            // Cabeçalho "Status / Bx. Autom" ou variantes
            const tlimp = tl.replace(/\\s+/g, ' ').trim();
            if (tlimp === 'status bx. autom' || tlimp === 'bx. autom') continue;

            if (vistosLocal.has(text)) continue;
            vistosLocal.add(text);

            let cells = [];
            const cellNodes = row.querySelectorAll(
                "td, .mat-mdc-cell, .mdc-data-table__cell, [role='cell']"
            );
            for (const c of cellNodes) {
                const t = (c.innerText || '').trim();
                if (t) cells.push(t);
            }

            if (cells.length === 0) {
                const partes = text.split(/\\r?\\n/);
                for (let p of partes) {
                    p = (p || '').trim();
                    if (!p) continue;
                    if (/^bx\\.?\\s*autom$/i.test(p)) continue;
                    if (/^status$/i.test(p)) continue;
                    cells.push(p);
                }
            }

            saida.push({
                texto: text,
                celulas: cells
            });
        }

        return saida;
    }
    """

    for loop in range(1, max_loops + 1):
        fazer_coleta = (loop == 1) or (loop % COLETA_INTERVALO == 0)
        debug(f"Pré-scroll loop={loop} (coleta={'SIM' if fazer_coleta else 'NÃO'})")

        if fazer_coleta:
            try:
                novas_linhas = page.evaluate(js_coleta, SELETOR_BALAO_AZUL)
            except Exception as e:
                print(f"Erro ao executar JS de coleta no loop {loop}: {e}")
                escrever_log("ERRO_JS_COLETA")
                novas_linhas = []

            for info in novas_linhas:
                texto = (info.get("texto") or "").strip()
                if texto and texto not in textos_vistos:
                    textos_vistos.add(texto)
                    linhas_inconsistentes.append(info)

            # Heurística: várias iterações sem novas linhas -> considera fim
            total_atual = len(textos_vistos)
            if total_atual == ultimo_total:
                loops_sem_novidade += 1
            else:
                loops_sem_novidade = 0
                ultimo_total = total_atual

            if loops_sem_novidade >= LIMITE_SEM_NOVIDADE:
                print("Várias iterações sem novas linhas Bx.Autom sem balão; encerrando pré-scroll (fim heurístico).")
                escrever_log("FIM_POR_SEM_NOVIDADE_BX_AUTOM")
                fim_confirmado = True
                break

        # a cada 15 loops, sobe um pouco antes de descer para “mexer” a lista
        subir_primeiro = (loop % 15 == 0)
        info_scroll = _scroll_lista_uma_vez(page, subir_primeiro=subir_primeiro)
        prev_scroll = info_scroll.get("prev", 0) or 0
        new_scroll = info_scroll.get("new", 0) or 0
        scrollable = info_scroll.get("scrollable", True)

        if not scrollable:
            print("Conteúdo da lista não é rolável (lista pequena). Encerrando pré-scroll (fim CONFIRMADO).")
            fim_confirmado = True
            break

        if info_scroll.get("reached") and new_scroll == prev_scroll:
            reached_seq += 1
        else:
            reached_seq = 0

        page.wait_for_timeout(SCROLL_PAUSE_MS)

        if (loop % CHECAR_FIM_INTERVALO == 0):
            if chegou_ao_fim_da_lista(page):
                print("Mensagem clara de fim encontrada; encerrando pré-scroll (fim CONFIRMADO).")
                fim_confirmado = True
                break

        if reached_seq >= REACHED_LIMIT:
            print("Scroll parece travado há vários loops. Encerrando pré-scroll (fim NÃO confirmado).")
            escrever_log("SCROLL_PARECE_TRAVADO_FIM_NAO_CONFIRMADO")
            break

    if not fim_confirmado and chegou_ao_fim_da_lista(page):
        fim_confirmado = True

    total_chips_aprox = len(linhas_inconsistentes)

    print(
        f"Pré-scroll finalizado; total de linhas inconsistentes coletadas="
        f"{len(linhas_inconsistentes)} (chips Bx.Autom sem balão). "
        f"Fim_confirmado={fim_confirmado}. Voltando ao topo..."
    )

    try:
        page.evaluate(
            """
            () => {
                const cand =
                    document.querySelector('#card-content .cdk-virtual-scroll-viewport') ||
                    document.querySelector('#card-content') ||
                    document.querySelector('.cdk-virtual-scroll-viewport') ||
                    document.querySelector('div[cdk-virtual-scroll-viewport]') ||
                    document.scrollingElement ||
                    document.documentElement ||
                    document.body;
                if (cand && cand.scrollTo) {
                    cand.scrollTo(0, 0);
                }
            }
            """
        )
    except Exception:
        pass
    page.wait_for_timeout(300)

    return total_chips_aprox, linhas_inconsistentes, fim_confirmado


# =====================================================
# EXTRAÇÃO AUXILIAR (LINHA DO SITE)
# =====================================================

def extrair_campos_auxiliares(texto_linha: str, celulas: list):
    # Garante que tudo seja string
    celulas_str = [str(c) if c is not None else "" for c in (celulas or [])]

    combined = " | ".join(celulas_str)
    if texto_linha:
        combined = combined + " | " + str(texto_linha)

    # Datas DD/MM/AAAA
    datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", combined)
    data_principal = datas[0] if datas else ""
    outras_datas = "; ".join(datas[1:]) if len(datas) > 1 else ""

    # Competências MM/AAAA
    competencias = re.findall(r"\b\d{2}/\d{4}\b", combined)
    competencia = competencias[0] if competencias else ""

    # Documento – primeiro tenta formatado, depois só dígitos (11 ou 14)
    cnpjs = re.findall(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", combined)
    cpfs = re.findall(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", combined)

    if cnpjs:
        documento = cnpjs[0]
    elif cpfs:
        documento = cpfs[0]
    else:
        # CNPJ/CPF só com dígitos (11 ou 14 dígitos)
        numeros_docs = re.findall(r"\b\d{14}\b|\b\d{11}\b", combined)
        documento = numeros_docs[0] if numeros_docs else ""

    # Responsável – qualquer letra entre parênteses (U, T, C, R, etc.)
    responsavel = ""
    for c in celulas_str:
        if re.search(r"\([A-Z]\)", c):
            responsavel = c.strip()
            break

    # Título: maior pedaço de texto relevante
    candidatos_titulo = []
    for c in celulas_str:
        s = c.strip()
        if not s:
            continue
        if re.search(r"\bBx\.?\s*Autom\b", s, re.I):
            continue
        if re.search(r"\bBx\.?\s*Manual\b", s, re.I):
            continue
        if re.fullmatch(r"\d+\s*/\s*\d+", s):
            continue
        if s in datas or s in competencias:
            continue
        candidatos_titulo.append(s)

    if candidatos_titulo:
        titulo = max(candidatos_titulo, key=len)
    else:
        titulo = celulas_str[0].strip() if celulas_str else ""

    primeira_celula = celulas_str[0] if celulas_str else ""

    return {
        "data_principal": data_principal,
        "outras_datas": outras_datas,
        "competencia": competencia,
        "documento": documento,
        "responsavel": responsavel,
        "titulo": titulo,
        "primeira_celula": primeira_celula,
    }


# =====================================================
# SALVAR EXCEL DE INCONSISTÊNCIAS (APENAS 1 ARQUIVO)
# =====================================================

def salvar_excel_inconsistencias(inconsistencias, caminho_saida: str):
    """
    Salva:
      - Arquivo principal (tareffa_inconsistencias_auto_sem_balao_v9.3.xlsx) BEM ENXUTO,
        focado em campos que ajudam a localizar no Tareffa.

    Lógica:
      - Tenta priorizar apenas linhas com Servico_casado_com_CSV = True.
      - SE NÃO TIVER NENHUMA linha casada, usa TODAS as linhas (para não ficar 15 dias sem nada).
    """
    if not inconsistencias:
        print("Nenhuma inconsistência detalhada para salvar ainda.")
        return

    df_full = pd.DataFrame(inconsistencias)

    # Garante colunas mínimas usadas na montagem
    for col in [
        "Dia_verificado",
        "Departamento (CSV)",
        "Servico (CSV)",
        "Linha completa (site)",
        "Data principal (linha Tareffa)",
        "Competencia detectada",
        "Documento detectado (CNPJ/CPF)",
        "Responsavel (linha Tareffa)",
        "Titulo detectado (linha Tareffa)",
    ]:
        if col not in df_full.columns:
            df_full[col] = ""

    # TENTAR priorizar só os serviços casados com CSV,
    # MAS se não tiver nenhum, usa TODAS as linhas (para não zerar o arquivo).
    df_main = df_full.copy()
    if "Servico_casado_com_CSV" in df_full.columns:
        mask_casado = df_full["Servico_casado_com_CSV"] == True
        if mask_casado.any():
            df_main = df_full[mask_casado].copy()
        else:
            print(
                "Nenhuma inconsitência casada com serviço do CSV; "
                "usando TODAS as linhas para não gerar planilha vazia."
            )

    if df_main.empty:
        print("Nenhuma inconsistência para salvar no arquivo principal.")
        pd.DataFrame(columns=[
            "Dia_verificado",
            "Departamento (CSV)",
            "Servico (CSV)",
            "Data principal (linha Tareffa)",
            "Competencia detectada",
            "Documento detectado (CNPJ/CPF)",
            "Titulo detectado (linha Tareffa)",
            "Responsavel (linha Tareffa)",
            "Chave_localizacao_Tareffa",
        ]).to_excel(caminho_saida, index=False)
        print(f"Arquivo de inconsistências (conciso) salvo/atualizado (vazio): {caminho_saida}")
        return

    # Ajusta o título para ficar só "número - nome da empresa"
    def extrair_empresa_de_titulo(t):
        if not isinstance(t, str):
            return t
        partes = [p.strip() for p in t.split(" - ") if p.strip()]
        # Ex.: "FGTS -  - 55 - DELRODO ..." -> partes = ["FGTS", "55", "DELRODO ..."]
        if len(partes) >= 2:
            num = partes[-2]
            nome = partes[-1]
            return f"{num} - {nome}"
        return t

    df_main["Titulo detectado (linha Tareffa)"] = df_main["Titulo detectado (linha Tareffa)"].apply(
        extrair_empresa_de_titulo
    )

    # Cria uma chave de busca mais amigável
    def montar_chave(row):
        partes = []
        if row["Data principal (linha Tareffa)"]:
            partes.append(str(row["Data principal (linha Tareffa)"]))
        if row["Competencia detectada"]:
            partes.append(str(row["Competencia detectada"]))
        if row["Documento detectado (CNPJ/CPF)"]:
            partes.append(str(row["Documento detectado (CNPJ/CPF)"]))
        if row["Titulo detectado (linha Tareffa)"]:
            partes.append(str(row["Titulo detectado (linha Tareffa)"]))
        if row["Responsavel (linha Tareffa)"]:
            partes.append(str(row["Responsavel (linha Tareffa)"]))
        if not partes and row["Linha completa (site)"]:
            partes.append(str(row["Linha completa (site)"])[:200])
        return " | ".join(partes)

    df_main["Chave_localizacao_Tareffa"] = df_main.apply(montar_chave, axis=1)

    colunas_main = [
        "Dia_verificado",
        "Departamento (CSV)",
        "Servico (CSV)",
        "Data principal (linha Tareffa)",
        "Competencia detectada",
        "Documento detectado (CNPJ/CPF)",
        "Titulo detectado (linha Tareffa)",
        "Responsavel (linha Tareffa)",
        "Chave_localizacao_Tareffa",
    ]

    df_main = df_main[colunas_main]

    # Remove duplicados por dia/depto/serviço/chave
    df_main = df_main.drop_duplicates(
        subset=[
            "Dia_verificado",
            "Departamento (CSV)",
            "Servico (CSV)",
            "Chave_localizacao_Tareffa",
        ]
    )

    df_main.to_excel(caminho_saida, index=False)
    print(f"Arquivo de inconsistências (conciso) salvo/atualizado: {caminho_saida}")


# =====================================================
# CONFIRMAÇÃO MODO MINIMALISTA
# =====================================================

def perguntar_modo_minimalista() -> bool:
    resp = input(
        "\nAtivar modo minimalista (rodar sem janela / headless)? [S/N]: "
    ).strip().lower()

    if not resp:
        return False

    primeira = resp[0]
    return primeira in ("s", "y")


# =====================================================
# MAIN
# =====================================================

def main():
    print("Iniciando robô Tareffa – v9.3 (sem filtro de Serviço, casando tudo pelo nome, com fallback)...\n")
    print(f"DEBUG_MODE: {'ATIVADO' if DEBUG_MODE else 'DESATIVADO'}")
    print(f"SELETOR_BALAO_AZUL: {SELETOR_BALAO_AZUL}")
    print(f"Arquivo de log: {LOG_ARQUIVO}\n")

    modo_minimalista = perguntar_modo_minimalista()
    print(f"MODO MINIMALISTA (headless): {'ATIVADO' if modo_minimalista else 'DESATIVADO'}\n")

    print("Antes de começar, informe o intervalo de datas a verificar.")
    print("O robô processa DIA A DIA dentro do intervalo.")
    print("Você pode digitar as datas como:")
    print("  • 01/11/2025  (com barras)")
    print("  • 01112025    (sem barras)\n")

    while True:
        entrada_ini = input("Data de INÍCIO (DD/MM/AAAA ou DDMMAAAA): ").strip()
        entrada_fim = input("Data de FIM   (DD/MM/AAAA ou DDMMAAAA): ").strip()

        try:
            dt_ini = interpretar_data_usuario(entrada_ini)
            dt_fim = interpretar_data_usuario(entrada_fim)

            if dt_ini > dt_fim:
                dt_ini, dt_fim = dt_fim, dt_ini

            break
        except ValueError:
            print("\nDatas inválidas. Use DD/MM/AAAA ou DDMMAAAA (ex: 01/12/2025 ou 01122025). Tente novamente.\n")

    # Caminhos dos arquivos (sempre na mesma pasta do script)
    caminho_csv = os.path.join(BASE_DIR, NOME_ARQUIVO_CSV)
    caminho_saida = os.path.join(BASE_DIR, NOME_ARQUIVO_SAIDA)
    caminho_controle = os.path.join(BASE_DIR, NOME_ARQUIVO_CONTROLE)
    caminho_lista_incons = os.path.join(BASE_DIR, NOME_ARQUIVO_SERVICOS_INCONS)
    caminho_dias_processados = os.path.join(BASE_DIR, NOME_ARQUIVO_DIAS_PROCESSADOS)

    servicos, df_controle = carregar_servicos_lista(caminho_csv, caminho_controle)

    dias_processados_set, df_dias = carregar_dias_processados(caminho_dias_processados)
    if dias_processados_set:
        print(f"Dias já checados em execuções anteriores: {sorted(dias_processados_set)}")

    def salvar_controle_parcial():
        df_controle.to_excel(caminho_controle, index=False)
        print(f"Arquivo de controle salvo: {caminho_controle}")

    salvar_cada_n_servicos = 20
    servicos_desde_ultimo_save = 0

    if not servicos:
        print("Nenhum serviço cadastrado (ou todos já com inconsistência registrada). Nada a fazer.")
        return

    # Agrupamento: Departamento -> nome_norm -> lista de serviços (IDs, nomes...)
    grupos_por_depto_nome = defaultdict(lambda: defaultdict(list))
    for svc in servicos:
        grupos_por_depto_nome[svc["departamento"]][svc["nome_norm"]].append(svc)

    total_grupos = sum(len(mapas) for mapas in grupos_por_depto_nome.values())
    print(f"Total de grupos distintos (Departamento + Nome de serviço) ATIVOS nesta rodada: {total_grupos}")

    todas_inconsistencias = []

    def extrair_texto_celulas(info):
        if not isinstance(info, dict):
            return str(info or ""), []
        texto = info.get("texto") or info.get("linha") or info.get("fullText") or ""
        celulas = info.get("celulas") or []
        if celulas is None:
            celulas = []
        return str(texto), list(map(str, celulas))

    # Lista de dias do intervalo
    dias = []
    dia_atual = dt_ini
    while dia_atual <= dt_fim:
        dias.append(dia_atual)
        dia_atual = dia_atual + timedelta(days=1)

    total_dias = len(dias)

    with sync_playwright() as p:
        browser = p.firefox.launch(
            headless=modo_minimalista,
            slow_mo=0,
        )
        context = browser.new_context(
            viewport={"width": 1280, "height": 720},
        )
        page = context.new_page()

        try:
            acessar_servicos_programados(page)
            selecionar_status_bx_autom(page)

            for idx_dia, dia in enumerate(dias, start=1):
                data_str = dia.strftime("%d/%m/%Y")

                print("\n" + "#" * 50)
                print(f"Dia {idx_dia}/{total_dias}: {data_str}")
                print("#" * 50)

                if data_str in dias_processados_set:
                    print(f"Dia {data_str} já estava marcado como verificado anteriormente. Pulando.")
                    escrever_log(f"DIA_{data_str}_JA_PROCESSADO_PULANDO")
                    continue

                # Garantir que estamos na tela correta antes de mexer em filtros
                garantir_pagina_servicos(page)

                aplicar_filtro_datas(page, data_str, data_str)

                data_verificacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                dia_concluido = True

                grupo_global_idx = 0
                total_linhas_bx_sem_balao_dia = 0
                total_servicos_com_incons_dia = 0

                for departamento, mapa_servicos in grupos_por_depto_nome.items():
                    print("\n" + "-" * 50)
                    print(
                        f"Departamento: {departamento!r} "
                        f"({len(mapa_servicos)} serviços distintos neste grupo) – Dia {data_str}"
                    )
                    print("-" * 50)

                    # Garantir novamente que estamos na tela certa
                    garantir_pagina_servicos(page)

                    try:
                        aplicar_filtro_avancado_departamento(page, departamento)
                    except Exception as e:
                        print(f"Erro ao aplicar filtro avançado para depto {departamento!r}: {e}")
                        escrever_log(f"ERRO_APLICAR_FILTRO_AVANCADO_DEPTO_{departamento.replace(' ','_')}")

                    try:
                        selecionar_status_bx_autom(page)
                    except Exception as e:
                        print(f"Erro ao garantir Status 'Bx. Autom' para o grupo: {e}")
                        escrever_log(f"ERRO_REAPLICAR_BX_AUTOM_DEPTO_{departamento.replace(' ','_')}")

                    # Carrega todas as linhas Bx.Autom sem balão para ESTE departamento/dia
                    _, linhas_incons_depto, fim_confirmado = carregar_todos_registros_lista(page)

                    if not fim_confirmado:
                        dia_concluido = False
                        print("Fim da lista NÃO confirmado para este departamento/dia; não vou marcar nada como OK definitivo.")

                    # Monta cache de linhas com texto normalizado
                    linhas_enriquecidas = []
                    for info in linhas_incons_depto:
                        texto_linha, celulas = extrair_texto_celulas(info)
                        linha_norm = normalizar_nome(texto_linha)
                        linhas_enriquecidas.append(
                            {
                                "info": info,
                                "texto": texto_linha,
                                "celulas": celulas,
                                "linha_norm": linha_norm,
                            }
                        )

                    total_linhas_bx_sem_balao_dia += len(linhas_enriquecidas)

                    debug(f"[DEBUG] Total de linhas Bx.Autom sem balão neste departamento/dia: {len(linhas_enriquecidas)}")
                    if DEBUG_MODE and linhas_enriquecidas:
                        for i, lin in enumerate(linhas_enriquecidas[:MAX_DEBUG_LINHAS], start=1):
                            debug(
                                f"  [DEBUG] LinhaDepto {i}: "
                                f"texto='{lin['texto'][:160].replace(chr(10),' ')}'"
                            )
                        if len(linhas_enriquecidas) > MAX_DEBUG_LINHAS:
                            debug(f"  [DEBUG] ... (+{len(linhas_enriquecidas) - MAX_DEBUG_LINHAS} linhas não exibidas)")

                    escrever_log(
                        f"[RESUMO_DEPARTAMENTO]_DIA={data_str}_DEPTO={departamento}_"
                        f"LINHAS_BX_AUTOM_SEM_BALAO={len(linhas_enriquecidas)}_"
                        f"FIM_CONFIRMADO={str(fim_confirmado).upper()}"
                    )

                    # Flags para saber quais linhas foram associadas a algum serviço
                    usados_por_algum_servico = [False] * len(linhas_enriquecidas)

                    # Para CADA NOME de serviço do departamento, contamos quantas linhas contêm aquele nome_norm
                    for nome_norm, lista_servicos in mapa_servicos.items():
                        grupo_global_idx += 1
                        svc_ref = lista_servicos[0]
                        nome_servico = svc_ref["nome"]
                        ids_repetidos = [s["id"] for s in lista_servicos]

                        print("\n" + "=" * 40)
                        print(
                            f"[Dia {data_str}] Grupo {grupo_global_idx}/{total_grupos} – Serviço '{nome_servico}' "
                            f"(Departamento: {departamento!r}, repetido {len(lista_servicos)} vez(es) no controle, IDs={ids_repetidos})"
                        )
                        print("=" * 40)

                        indices_para_servico = [
                            idx for idx, lin in enumerate(linhas_enriquecidas)
                            if nome_norm and nome_norm in lin["linha_norm"]
                        ]
                        linhas_para_servico = [linhas_enriquecidas[idx] for idx in indices_para_servico]

                        for idx in indices_para_servico:
                            usados_por_algum_servico[idx] = True

                        debug(
                            f"[DEBUG] Linhas desse departamento/dia que contêm o serviço '{nome_servico}' "
                            f"(nome_norm='{nome_norm[:60]}'): {len(linhas_para_servico)}"
                        )

                        qtd_incons = len(linhas_para_servico)

                        # Para cada ID desse serviço, aplica o resultado
                        for svc in lista_servicos:
                            svc_id = svc["id"]
                            mask = df_controle["ID"].astype(str) == str(svc_id)
                            if not mask.any():
                                print(f"ID {svc_id} não encontrado no DataFrame de controle.")
                                continue

                            prev_qtd = pd.to_numeric(
                                df_controle.loc[mask, "Qtd_inconsistencias"],
                                errors="coerce"
                            ).fillna(0).astype(int)

                            if qtd_incons == 0:
                                if not fim_confirmado:
                                    print(f"(ID {svc_id}) Lista deste departamento/dia NÃO foi confirmada até o fim; não vou marcar como OK. Mantendo Status anterior.")
                                    df_controle.loc[mask, "Ultima_verificacao"] = data_verificacao
                                else:
                                    print(f"(ID {svc_id}) Nenhuma inconsistência (Bx. Autom sem balão) para este serviço neste dia.")
                                    status_atual = df_controle.loc[mask, "Status"].astype(str).str.upper()
                                    if not (status_atual == "COM INCONSISTENCIA").any():
                                        df_controle.loc[mask, "Status"] = f"OK ({data_str} a {data_str})"
                                        df_controle.loc[mask, "Periodo_inicio"] = data_str
                                        df_controle.loc[mask, "Periodo_fim"] = data_str
                                    df_controle.loc[mask, "Qtd_inconsistencias"] = prev_qtd.astype(str)
                                    df_controle.loc[mask, "Ultima_verificacao"] = data_verificacao
                            else:
                                print(f"(ID {svc_id}) {qtd_incons} inconsistência(s) (Bx. Autom sem balão) para este serviço neste dia.")
                                total_servicos_com_incons_dia += 1

                                novo_total_qtd = prev_qtd + qtd_incons
                                df_controle.loc[mask, "Qtd_inconsistencias"] = novo_total_qtd.astype(str)
                                df_controle.loc[mask, "Status"] = "COM INCONSISTENCIA"
                                df_controle.loc[mask, "Ultima_verificacao"] = data_verificacao
                                df_controle.loc[mask, "Periodo_inicio"] = data_str
                                df_controle.loc[mask, "Periodo_fim"] = data_str

                                # Gera linhas detalhadas de inconsistência (casadas com serviço do CSV)
                                for lin in linhas_para_servico:
                                    texto_linha = lin["texto"]
                                    celulas = lin["celulas"]
                                    extras = extrair_campos_auxiliares(texto_linha, celulas)

                                    row_data = {
                                        "Departamento (CSV)": departamento,
                                        "Servico (CSV)": nome_servico,
                                        "Servico normalizado": nome_norm,
                                        "ID_controle": svc_id,
                                        "Servico_casado_com_CSV": True,
                                        "Tem balão azul?": False,
                                        "Linha completa (site)": texto_linha,
                                        "Celulas (raw)": " | ".join(celulas),
                                        "Data principal (linha Tareffa)": extras["data_principal"],
                                        "Outras datas (linha Tareffa)": extras["outras_datas"],
                                        "Competencia detectada": extras["competencia"],
                                        "Documento detectado (CNPJ/CPF)": extras["documento"],
                                        "Responsavel (linha Tareffa)": extras["responsavel"],
                                        "Titulo detectado (linha Tareffa)": extras["titulo"],
                                        "Primeira celula visivel": extras["primeira_celula"],
                                        "Data/hora verificação": data_verificacao,
                                        "Dia_verificado": data_str,
                                    }
                                    for i in range(6):
                                        key = f"Celula_{i+1}"
                                        row_data[key] = celulas[i] if i < len(celulas) else ""
                                    todas_inconsistencias.append(row_data)

                                # Exemplo de linha (pega a última)
                                texto_exemplo = linhas_para_servico[-1]["texto"]
                                df_controle.loc[mask, "Exemplo_linha"] = texto_exemplo

                            servicos_desde_ultimo_save += 1
                            if servicos_desde_ultimo_save >= salvar_cada_n_servicos:
                                salvar_controle_parcial()
                                servicos_desde_ultimo_save = 0

                    # Fallback: linhas sem balão que NÃO casaram com nenhum serviço do CSV
                    linhas_sem_vinculo = [
                        linhas_enriquecidas[idx]
                        for idx, usado in enumerate(usados_por_algum_servico)
                        if not usado
                    ]

                    if linhas_sem_vinculo:
                        print(
                            f"Foram encontradas {len(linhas_sem_vinculo)} linha(s) Bx.Autom sem balão "
                            f"neste departamento/dia que não casaram com nenhum serviço do CSV."
                        )
                        escrever_log(
                            f"LINHAS_SEM_VINCULO_CSV_DIA={data_str}_DEPTO={departamento}_QTDE={len(linhas_sem_vinculo)}"
                        )

                        if SALVAR_LINHAS_SEM_VINCULO:
                            for lin in linhas_sem_vinculo:
                                texto_linha = lin["texto"]
                                celulas = lin["celulas"]
                                extras = extrair_campos_auxiliares(texto_linha, celulas)

                                row_data = {
                                    "Departamento (CSV)": departamento,
                                    "Servico (CSV)": "",
                                    "Servico normalizado": "",
                                    "ID_controle": "",
                                    "Servico_casado_com_CSV": False,
                                    "Tem balão azul?": False,
                                    "Linha completa (site)": texto_linha,
                                    "Celulas (raw)": " | ".join(celulas),
                                    "Data principal (linha Tareffa)": extras["data_principal"],
                                    "Outras datas (linha Tareffa)": extras["outras_datas"],
                                    "Competencia detectada": extras["competencia"],
                                    "Documento detectado (CNPJ/CPF)": extras["documento"],
                                    "Responsavel (linha Tareffa)": extras["responsavel"],
                                    "Titulo detectado (linha Tareffa)": extras["titulo"],
                                    "Primeira celula visivel": extras["primeira_celula"],
                                    "Data/hora verificação": data_verificacao,
                                    "Dia_verificado": data_str,
                                }
                                for i in range(6):
                                    key = f"Celula_{i+1}"
                                    row_data[key] = celulas[i] if i < len(celulas) else ""
                                todas_inconsistencias.append(row_data)

                    print(
                        f"\nEncerrando grupo do departamento {departamento!r} para o dia {data_str}, "
                        "limpando filtro de departamento..."
                    )

                    try:
                        limpar_filtro_avancado_departamento(page)
                    except Exception as e:
                        print(f"Erro ao limpar filtro avançado entre departamentos: {e}")
                        escrever_log(f"ERRO_LIMPAR_FILTRO_AVANCADO_DEPTO_{departamento.replace(' ','_')}")

                    try:
                        selecionar_status_bx_autom(page)
                    except Exception as e:
                        print(f"Erro ao re-aplicar Status 'Bx. Autom' entre departamentos: {e}")
                        escrever_log(f"ERRO_REAPLICAR_BX_AUTOM_POS_LIMPEZA_DEPTO_{departamento.replace(' ','_')}")

                # Resumo do DIA (para log)
                escrever_log(
                    f"[RESUMO_DIA]_DIA={data_str}_SERVICOS_COM_INCONS={total_servicos_com_incons_dia}_"
                    f"LINHAS_BX_AUTOM_SEM_BALAO={total_linhas_bx_sem_balao_dia}_"
                    f"DIA_CONCLUIDO={str(dia_concluido).upper()}"
                )

                print(f"\nSalvando arquivos após o dia {data_str}...")

                salvar_controle_parcial()
                salvar_excel_inconsistencias(todas_inconsistencias, caminho_saida)

                df_incons_parcial = df_controle[df_controle["Status"] == "COM INCONSISTENCIA"].copy()
                if not df_incons_parcial.empty:
                    df_incons_parcial.to_excel(caminho_lista_incons, index=False)
                    print(f"Lista parcial de serviços com inconsistência salva em: {caminho_lista_incons}")
                else:
                    print("Até agora nenhum serviço está marcado como 'COM INCONSISTENCIA'.")

                now_exec = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                if dia_concluido:
                    if data_str in dias_processados_set:
                        df_dias.loc[df_dias["Data"] == data_str, "Ultima_execucao"] = now_exec
                    else:
                        df_dias = pd.concat(
                            [df_dias, pd.DataFrame([{"Data": data_str, "Ultima_execucao": now_exec}])],
                            ignore_index=True,
                        )
                        dias_processados_set.add(data_str)

                    salvar_dias_processados(df_dias, caminho_dias_processados)
                    print(f"Dia {data_str} marcado como VERIFICADO com sucesso.")
                else:
                    print(
                        f"Dia {data_str} NÃO foi marcado como verificado porque houve pelo menos "
                        "um departamento sem fim confirmado (possível travamento). "
                        "Esse dia será tentado novamente em futuras execuções."
                    )
                    escrever_log(
                        f"DIA_{data_str}_NAO_MARCADO_COMO_VERIFICADO_FIM_NAO_CONFIRMADO"
                    )

        finally:
            browser.close()

    df_controle.to_excel(caminho_controle, index=False)
    print(f"Arquivo de controle FINAL atualizado: {caminho_controle}")

    df_incons = df_controle[df_controle["Status"] == "COM INCONSISTENCIA"].copy()
    if not df_incons.empty:
        df_incons.to_excel(caminho_lista_incons, index=False)
        print(f"Lista FINAL de serviços com inconsistência salva em: {caminho_lista_incons}")
    else:
        print("Nenhum serviço ficou marcado como 'COM INCONSISTENCIA' neste intervalo.")

    salvar_excel_inconsistencias(todas_inconsistencias, caminho_saida)
    salvar_dias_processados(df_dias, caminho_dias_processados)


if __name__ == "__main__":
    main()
