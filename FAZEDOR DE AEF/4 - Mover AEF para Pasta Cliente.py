r"""
4 - Mover AEF para Pasta Cliente

Objetivo:
- Mover o arquivo bruto baixado no Script 1 (Balancete_*.xls/xlsx)
- Mover o arquivo formatado no Script 2 (BALANCETE AEF *.xlsx)
- Destino padrao: W:\PASTA CLIENTES\<CLIENTE>\AEF\YYYY\MM
"""

import argparse
import os
import shutil
import sys
from datetime import datetime


# =========================
# Configuracoes
# =========================

BASE_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\FAZEDOR DE AEF"
PASTA_ARQUIVOS = os.path.join(BASE_DIR, "Arquivos")

EMPRESA_PADRAO = "22"
DESTINO_CLIENTE_PADRAO = r"W:\PASTA CLIENTES\LOBO MOTOS LTDA\AEF"

# Destinos por empresa (base da AEF; ano/mes e resolvido automaticamente).
DESTINOS_POR_EMPRESA = {
    "22": r"W:\PASTA CLIENTES\LOBO MOTOS LTDA\AEF",
    "3": r"W:\PASTA CLIENTES\MOTOACAO MOTOCICLETAS E NAUTICA LTDA\AEF",
    "143-1": r"W:\PASTA CLIENTES\TIBAGI MOTOS LTDA\AEF",
    "143-2": r"W:\PASTA CLIENTES\TIBAGI MOTOS LTDA\AEF",
    "201": r"W:\PASTA CLIENTES\RIO BRANCO VEÍCULOS LTDA\AEF",
}

# Sufixo adicional para diferenciar filiais/grupos no nome do arquivo.
SUFIXO_ARQUIVO_POR_EMPRESA = {
    "143-1": "-1",
    "143-2": "-2",
}

PADRAO_BRUTO_XLSX = "Balancete_{empresa}.xlsx"
PADRAO_BRUTO_XLS = "Balancete_{empresa}.xls"
PADRAO_FORMATADO = "BALANCETE AEF {empresa}.xlsx"

NOME_BRUTO_DESTINO = "AEF BRUTO {data}{ext}"
NOME_FORMATADO_DESTINO = "AEF FORMATADO {data}{ext}"
FORMATO_DATA_NOME = "%d-%m-%Y"


# =========================
# Utilitarios
# =========================


def _normalizar_empresa(texto: str) -> str:
    return (texto or "").strip().lstrip("\ufeff")


def _montar_destino_mes(base_cliente: str, data_ref: datetime) -> str:
    ano = data_ref.strftime("%Y")
    mes = data_ref.strftime("%m")

    base_norm = os.path.normpath(base_cliente)
    ultimo = os.path.basename(base_norm)
    # Se a base ja vier como ...\AEF\YYYY, adiciona so o mes.
    if len(ultimo) == 4 and ultimo.isdigit():
        return os.path.join(base_norm, mes)
    return os.path.join(base_norm, ano, mes)


def _resolver_origem_bruto(pasta_empresa: str, empresa: str) -> str:
    candidatos = [
        os.path.join(pasta_empresa, PADRAO_BRUTO_XLSX.format(empresa=empresa)),
        os.path.join(pasta_empresa, PADRAO_BRUTO_XLS.format(empresa=empresa)),
    ]
    for caminho in candidatos:
        if os.path.isfile(caminho):
            return caminho
    raise FileNotFoundError(
        f"Arquivo bruto nao encontrado para empresa {empresa}. "
        f"Esperado: {os.path.basename(candidatos[0])} ou {os.path.basename(candidatos[1])}."
    )


def _resolver_origem_formatado(pasta_empresa: str, empresa: str) -> str:
    caminho = os.path.join(pasta_empresa, PADRAO_FORMATADO.format(empresa=empresa))
    if os.path.isfile(caminho):
        return caminho
    raise FileNotFoundError(
        f"Arquivo formatado nao encontrado para empresa {empresa}. "
        f"Esperado: {os.path.basename(caminho)}."
    )


def _mover_com_nome(caminho_origem: str, pasta_destino: str, nome_destino: str) -> str:
    os.makedirs(pasta_destino, exist_ok=True)
    caminho_final = os.path.join(pasta_destino, nome_destino)
    if os.path.exists(caminho_final):
        os.remove(caminho_final)
    shutil.move(caminho_origem, caminho_final)
    return caminho_final


# =========================
# Main
# =========================


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Move arquivos AEF (bruto e formatado) para pasta do cliente.")
    p.add_argument(
        "--empresa",
        default=EMPRESA_PADRAO,
        help=f"Codigo da empresa (padrao: {EMPRESA_PADRAO}).",
    )
    p.add_argument(
        "--destino-cliente",
        default="",
        help=(
            "Pasta base do cliente (sem ano/mes). "
            "Se vazio, usa o mapa interno por codigo da empresa."
        ),
    )
    p.add_argument(
        "--somente-listar",
        action="store_true",
        help="Apenas exibe origem/destino sem mover arquivos.",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()
    empresa = _normalizar_empresa(args.empresa)
    if not empresa:
        print("ERRO: --empresa vazio.")
        return 1

    pasta_empresa = os.path.join(PASTA_ARQUIVOS, empresa)
    if not os.path.isdir(pasta_empresa):
        print(f"ERRO: pasta da empresa nao encontrada: {pasta_empresa}")
        return 1

    data_ref = datetime.now()
    data_nome = data_ref.strftime(FORMATO_DATA_NOME)
    destino_base_cfg = (args.destino_cliente or "").strip()
    if not destino_base_cfg:
        destino_base_cfg = DESTINOS_POR_EMPRESA.get(empresa, "")
    if not destino_base_cfg:
        print(
            "ERRO: destino nao configurado para a empresa. "
            "Informe --destino-cliente ou adicione no mapa DESTINOS_POR_EMPRESA."
        )
        return 1

    pasta_destino = _montar_destino_mes(destino_base_cfg, data_ref)

    try:
        origem_bruto = _resolver_origem_bruto(pasta_empresa, empresa)
        origem_formatado = _resolver_origem_formatado(pasta_empresa, empresa)
    except FileNotFoundError as exc:
        print(f"ERRO: {exc}")
        return 1

    ext_bruto = os.path.splitext(origem_bruto)[1]
    ext_formatado = os.path.splitext(origem_formatado)[1]
    sufixo = SUFIXO_ARQUIVO_POR_EMPRESA.get(empresa, "")
    nome_bruto = NOME_BRUTO_DESTINO.format(data=data_nome, ext=sufixo + ext_bruto)
    nome_formatado = NOME_FORMATADO_DESTINO.format(data=data_nome, ext=sufixo + ext_formatado)

    print(f"Empresa: {empresa}")
    print(f"Origem bruto: {origem_bruto}")
    print(f"Origem formatado: {origem_formatado}")
    print(f"Destino base: {pasta_destino}")
    print(f"Nome bruto destino: {nome_bruto}")
    print(f"Nome formatado destino: {nome_formatado}")

    if args.somente_listar:
        print("Modo somente-listar: nenhuma movimentacao executada.")
        return 0

    destino_bruto = _mover_com_nome(origem_bruto, pasta_destino, nome_bruto)
    destino_formatado = _mover_com_nome(origem_formatado, pasta_destino, nome_formatado)

    print("OK: arquivos movidos com sucesso.")
    print(f"Bruto -> {destino_bruto}")
    print(f"Formatado -> {destino_formatado}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
