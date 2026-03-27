# 2 - Fromatador XLSX

# =========================
# Configuracoes
# =========================

import os
import shutil
import sys
import argparse

try:
    import win32com.client as win32
except Exception as exc:
    print("ERRO: pywin32 nao encontrado.")
    print("Instale: pip install pywin32")
    print(f"Detalhe: {exc}")
    raise


BASE_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\FAZEDOR DE AEF"
PASTA_ARQUIVOS = os.path.join(BASE_DIR, "Arquivos")
CAMINHO_EMPRESAS = os.path.join(BASE_DIR, "empresas.txt")
CAMINHO_EMPRESAS_ALTERNATIVO = os.path.join(BASE_DIR, "7 empresas.txt")

CAMINHO_MODELO = os.path.join(BASE_DIR, "PLANILHA AEF NOVEMBRO.xlsx")
NOME_SAIDA_PREFIXO = "final"

PLANILHA_DESTINO = "balancete"
PLANILHA_ORIGEM = ""  # vazio = unica planilha do balancete baixado

COPIAR_BALANCETE_PARA_PASTA_SCRIPT = False
REMOVER_DC_BALANCETE = True
COLUNA_DC = "H"
PLANILHAS_VALIDACAO = ["DRE", "Passivo", "Ativo"]
NOME_RELATORIO_ERRO = "Erros formatação.txt"

# =========================
# Utilitarios
# =========================


def carregar_empresas() -> list[str]:
    caminho = CAMINHO_EMPRESAS
    if not os.path.isfile(caminho) and os.path.isfile(CAMINHO_EMPRESAS_ALTERNATIVO):
        caminho = CAMINHO_EMPRESAS_ALTERNATIVO

    if not os.path.isfile(caminho):
        print(f"ERRO: arquivo de empresas nao encontrado: {CAMINHO_EMPRESAS}")
        sys.exit(1)

    with open(caminho, "r", encoding="utf-8") as arquivo:
        empresas = [_normalizar_empresa(linha) for linha in arquivo.readlines() if linha.strip()]

    if not empresas:
        print("ERRO: lista de empresas vazia.")
        sys.exit(1)

    return empresas


def _normalizar_empresa(texto: str) -> str:
    return texto.strip().lstrip("\ufeff")


def localizar_balancete(pasta_empresa: str, empresa: str) -> str:
    candidatos = [
        os.path.join(pasta_empresa, f"Balancete_{empresa}.xlsx"),
        os.path.join(pasta_empresa, f"Balancete_{empresa}.xls"),
    ]
    for c in candidatos:
        if os.path.isfile(c):
            return c

    if os.path.isdir(pasta_empresa):
        arquivos = [f for f in os.listdir(pasta_empresa) if f.lower().endswith((".xls", ".xlsx"))]
        if len(arquivos) == 1:
            return os.path.join(pasta_empresa, arquivos[0])
        if arquivos:
            print(f"ERRO: multiplos arquivos encontrados em {pasta_empresa}: {arquivos}")
            sys.exit(1)

    encontrados = buscar_balancete_global(empresa)
    if len(encontrados) == 1:
        return encontrados[0]
    if len(encontrados) > 1:
        print(f"ERRO: multiplos balancetes encontrados para {empresa}: {encontrados}")
        sys.exit(1)

    print(f"ERRO: balancete nao encontrado para a empresa {empresa}.")
    sys.exit(1)


def copiar_modelo(destino: str) -> None:
    if not os.path.isfile(CAMINHO_MODELO):
        print(f"ERRO: arquivo modelo nao encontrado: {CAMINHO_MODELO}")
        sys.exit(1)
    shutil.copy2(CAMINHO_MODELO, destino)


def copiar_balancete_para_pasta_script(caminho_balancete: str, empresa: str) -> None:
    if not COPIAR_BALANCETE_PARA_PASTA_SCRIPT:
        return

    _, ext = os.path.splitext(caminho_balancete)
    destino = os.path.join(BASE_DIR, f"Balancete_{empresa}{ext}")
    shutil.copy2(caminho_balancete, destino)


def limpar_dc_balancete(caminho_balancete: str) -> None:
    if not REMOVER_DC_BALANCETE:
        return

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(caminho_balancete)
        if wb.ReadOnly:
            print("ERRO: balancete aberto em outra instancia. Feche o arquivo e tente novamente.")
            sys.exit(1)

        if wb.Worksheets.Count != 1:
            print("ERRO: balancete possui mais de uma planilha.")
            sys.exit(1)
        ws = wb.Worksheets(1)

        usado = ws.UsedRange
        linhas = usado.Rows.Count
        inicio = usado.Row
        col = ord(COLUNA_DC.upper()) - ord("A") + 1
        rng = ws.Range(ws.Cells(inicio, col), ws.Cells(inicio + linhas - 1, col))
        valores = rng.Value2

        if not isinstance(valores, tuple):
            valores = ((valores,),)

        novos_valores: list[tuple] = []
        for linha in valores:
            atual = linha[0]
            if isinstance(atual, str):
                novo = atual.replace("D", "").replace("C", "")
            else:
                novo = atual
            novos_valores.append((novo,))

        rng.Value2 = novos_valores
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(True)
        excel.Quit()


def buscar_balancete_global(empresa: str) -> list[str]:
    encontrados: list[str] = []
    alvo = f"balancete_{empresa}".lower()
    for raiz, _, arquivos in os.walk(PASTA_ARQUIVOS):
        for nome in arquivos:
            if not nome.lower().endswith((".xls", ".xlsx")):
                continue
            nome_normalizado = nome.lower().replace("\ufeff", "")
            base, _ = os.path.splitext(nome_normalizado)
            if base == alvo:
                encontrados.append(os.path.join(raiz, nome))
    return encontrados


def colar_planilha(
    caminho_origem: str,
    caminho_destino: str,
    planilha_destino: str,
    planilha_origem: str,
) -> None:
    xl_paste_values = -4163
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb_origem = None
    wb_destino = None
    try:
        wb_origem = excel.Workbooks.Open(caminho_origem)
        wb_destino = excel.Workbooks.Open(caminho_destino)
        if wb_destino.ReadOnly:
            print("ERRO: arquivo destino aberto em outra instancia. Feche o arquivo e tente novamente.")
            sys.exit(1)

        if planilha_origem:
            ws_origem = wb_origem.Worksheets(planilha_origem)
        else:
            if wb_origem.Worksheets.Count != 1:
                print("ERRO: balancete possui mais de uma planilha.")
                sys.exit(1)
            ws_origem = wb_origem.Worksheets(1)

        nomes_destino = [wb_destino.Worksheets(i).Name for i in range(1, wb_destino.Worksheets.Count + 1)]
        if planilha_destino not in nomes_destino:
            print(f"ERRO: planilha destino '{planilha_destino}' nao encontrada no modelo.")
            sys.exit(1)
        if nomes_destino[0] != planilha_destino:
            print(f"AVISO: planilha destino '{planilha_destino}' nao e a primeira no modelo.")

        ws_destino = wb_destino.Worksheets(planilha_destino)
        usado = ws_origem.UsedRange
        dados = usado.Value2
        if dados is None:
            print("ERRO: balancete vazio.")
            sys.exit(1)

        origem_a1 = ws_origem.Cells(1, 1).Value2

        if not isinstance(dados, tuple):
            dados = ((dados,),)
        elif dados and not isinstance(dados[0], tuple):
            dados = (dados,)

        linhas = len(dados)
        colunas = len(dados[0]) if linhas > 0 else 0

        if linhas == 0 or colunas == 0:
            print("ERRO: balancete vazio.")
            sys.exit(1)

        ws_destino.Range("A1").Resize(linhas, colunas).Value2 = dados

        destino_a1 = ws_destino.Cells(1, 1).Value2
        if origem_a1 not in [None, ""] and destino_a1 in [None, ""]:
            usado.Copy()
            ws_destino.Range("A1").PasteSpecial(Paste=xl_paste_values)
            excel.CutCopyMode = False

        gerar_relatorio_erros_formatacao(
            wb_destino=wb_destino,
            ws_balancete=ws_destino,
            caminho_destino=caminho_destino,
        )

        wb_destino.Save()
    finally:
        if wb_origem is not None:
            wb_origem.Close(False)
        if wb_destino is not None:
            wb_destino.Close(True)
        excel.Quit()


def _normalizar_codigo(texto: str) -> str:
    return texto.strip().lstrip("\ufeff")


def _somente_digitos(texto: str) -> str:
    return "".join(ch for ch in texto if ch.isdigit())


def _codigo_numeros_pontos(texto: str) -> bool:
    if not texto:
        return False
    if not any(ch.isdigit() for ch in texto):
        return False
    return all((ch.isdigit() or ch == ".") for ch in texto)


def _prefixos_grupo(codigo: str) -> list[str]:
    partes = [p for p in codigo.split(".") if p]
    prefixos: list[str] = []
    for i in range(1, len(partes)):
        prefixos.append(".".join(partes[:i]))
    return prefixos


def _valor_para_texto(valor: object) -> str:
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return ""
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor)
    return str(valor)


def _valor_para_codigo(valor: object) -> str:
    return _normalizar_codigo(_valor_para_texto(valor))


def _tem_conteudo(valor: object) -> bool:
    if valor is None:
        return False
    if isinstance(valor, str) and not valor.strip():
        return False
    return True


def extrair_codigos_planilhas_validacao(wb_destino) -> set[str]:
    nomes = [wb_destino.Worksheets(i).Name for i in range(1, wb_destino.Worksheets.Count + 1)]
    faltando = [p for p in PLANILHAS_VALIDACAO if p not in nomes]
    if faltando:
        print(f"ERRO: planilhas nao encontradas: {faltando}")
        sys.exit(1)

    codigos: set[str] = set()
    for nome in PLANILHAS_VALIDACAO:
        ws = wb_destino.Worksheets(nome)
        usado = ws.UsedRange
        if usado is None:
            continue
        dados = usado.Value2
        if dados is None:
            continue
        if not isinstance(dados, tuple):
            dados = ((dados,),)
        elif dados and not isinstance(dados[0], tuple):
            dados = (dados,)

        for linha in dados:
            if not isinstance(linha, tuple):
                linha = (linha,)
            for valor in linha:
                codigo = _valor_para_codigo(valor)
                if codigo:
                    codigos.add(codigo)

    return codigos


def gerar_relatorio_erros_formatacao(wb_destino, ws_balancete, caminho_destino: str) -> None:
    codigos_validos = extrair_codigos_planilhas_validacao(wb_destino)
    codigos_prefixo7: set[str] = set()
    codigos_grupo: set[str] = set()
    for codigo in codigos_validos:
        digitos = _somente_digitos(codigo)
        if len(digitos) == 7:
            codigos_prefixo7.add(digitos)
        codigo_base = codigo.strip(".")
        if _codigo_numeros_pontos(codigo_base):
            codigos_grupo.add(codigo_base)

    usado = ws_balancete.UsedRange
    inicio = usado.Row
    linhas = usado.Rows.Count
    if linhas <= 0:
        return

    dados = ws_balancete.Range(
        ws_balancete.Cells(inicio, 1), ws_balancete.Cells(inicio + linhas - 1, 3)
    ).Value2
    if not isinstance(dados, tuple):
        dados = (dados,)

    faltando: list[tuple[str, str, str]] = []
    for linha in dados:
        if not isinstance(linha, tuple):
            linha = (linha,)
        valor_a = linha[0] if len(linha) > 0 else None
        valor_b = linha[1] if len(linha) > 1 else None
        valor_c = linha[2] if len(linha) > 2 else None

        codigo = _valor_para_codigo(valor_a)
        if not codigo:
            continue
        if not _tem_conteudo(valor_b):
            continue
        if codigo in codigos_validos:
            continue
        digitos_codigo = _somente_digitos(codigo)
        if len(digitos_codigo) >= 7 and digitos_codigo[:7] in codigos_prefixo7:
            continue
        codigo_base = codigo.strip(".")
        prefixos = _prefixos_grupo(codigo_base)
        if any(p in codigos_grupo for p in prefixos):
            continue

        faltando.append(
            (
                codigo,
                _valor_para_texto(valor_b),
                _valor_para_texto(valor_c),
            )
        )

    caminho_relatorio = os.path.join(os.path.dirname(caminho_destino), NOME_RELATORIO_ERRO)
    if not faltando:
        if os.path.isfile(caminho_relatorio):
            os.remove(caminho_relatorio)
        return

    with open(caminho_relatorio, "w", encoding="utf-8") as arquivo:
        for codigo, col_b, col_c in faltando:
            arquivo.write(f"{codigo}\t{col_b}\t{col_c}\n")

    print(f"Relatorio de erros gerado: {caminho_relatorio}")


def processar_empresa(empresa: str) -> None:
    pasta_empresa = os.path.join(PASTA_ARQUIVOS, empresa)
    os.makedirs(pasta_empresa, exist_ok=True)

    caminho_balancete = localizar_balancete(pasta_empresa, empresa)
    _, ext = os.path.splitext(caminho_balancete)
    caminho_padrao = os.path.join(pasta_empresa, f"Balancete_{empresa}{ext}")
    if os.path.normcase(caminho_balancete) != os.path.normcase(caminho_padrao):
        shutil.copy2(caminho_balancete, caminho_padrao)
        caminho_balancete = caminho_padrao
    copiar_balancete_para_pasta_script(caminho_balancete, empresa)
    limpar_dc_balancete(caminho_balancete)

    caminho_saida = os.path.join(pasta_empresa, f"{NOME_SAIDA_PREFIXO}_{empresa}.xlsx")
    copiar_modelo(caminho_saida)

    colar_planilha(
        caminho_origem=caminho_balancete,
        caminho_destino=caminho_saida,
        planilha_destino=PLANILHA_DESTINO,
        planilha_origem=PLANILHA_ORIGEM,
    )


def conferir_arquivo(caminho_final: str) -> None:
    if not os.path.isfile(caminho_final):
        print(f"ERRO: arquivo final nao encontrado: {caminho_final}")
        return

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(caminho_final)
        if wb.ReadOnly:
            print("ERRO: arquivo final aberto em outra instancia. Feche o arquivo e tente novamente.")
            return

        nomes = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        if PLANILHA_DESTINO not in nomes:
            print(f"ERRO: planilha '{PLANILHA_DESTINO}' nao encontrada no arquivo final.")
            return

        ws_balancete = wb.Worksheets(PLANILHA_DESTINO)
        gerar_relatorio_erros_formatacao(
            wb_destino=wb,
            ws_balancete=ws_balancete,
            caminho_destino=caminho_final,
        )
    finally:
        if wb is not None:
            wb.Close(False)
        excel.Quit()


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--conferir",
        action="store_true",
        help="Apenas confere o arquivo final ja formatado.",
    )
    parser.add_argument(
        "--arquivo",
        help="Caminho do arquivo final para conferencia.",
    )
    parser.add_argument(
        "--empresa",
        help="Numero da empresa para localizar o final_<empresa>.xlsx na pasta Arquivos.",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    if args.conferir:
        if args.arquivo:
            conferir_arquivo(args.arquivo)
            return

        if args.empresa:
            empresas = [args.empresa]
        else:
            empresas = carregar_empresas()

        for empresa in empresas:
            caminho_final = os.path.join(
                PASTA_ARQUIVOS,
                empresa,
                f"{NOME_SAIDA_PREFIXO}_{empresa}.xlsx",
            )
            print(f"Conferindo empresa: {empresa}")
            conferir_arquivo(caminho_final)
        return

    empresas = carregar_empresas()
    for empresa in empresas:
        print(f"Processando empresa: {empresa}")
        processar_empresa(empresa)


if __name__ == "__main__":
    main()
