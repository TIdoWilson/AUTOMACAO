import os
import sys
import subprocess
from datetime import date

import pandas as pd


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_CNPJS_BASE = os.path.join(BASE_DIR, "cnpjs.xlsx")

DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
LIMPOS_DIR = os.path.join(DOWNLOAD_DIR, "arquivos limpos")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(LIMPOS_DIR, exist_ok=True)

# partes de entrada/saida
CNPJS_PARTES = [
    os.path.join(BASE_DIR, "cnpjs_parte1.xlsx"),
    os.path.join(BASE_DIR, "cnpjs_parte2.xlsx"),
    os.path.join(BASE_DIR, "cnpjs_parte3.xlsx"),
]

REG_PARTES = [
    os.path.join(DOWNLOAD_DIR, "registros_icr_parte1.xlsx"),
    os.path.join(DOWNLOAD_DIR, "registros_icr_parte2.xlsx"),
    os.path.join(DOWNLOAD_DIR, "registros_icr_parte3.xlsx"),
]

REG_FINAL = os.path.join(DOWNLOAD_DIR, "registros_icr.xlsx")
HIST_EXECUCOES = os.path.join(DOWNLOAD_DIR, "historico_execucoes.xlsx")


def codigo_para_nome_arquivo(codigo: str) -> str:
    import re
    return re.sub(r"[^\w\-]", "_", str(codigo))


def dividir_cnpjs_em_3():
    df = pd.read_excel(EXCEL_CNPJS_BASE, dtype=str)
    lista = df.iloc[:, 0].dropna().astype(str).tolist()
    total = len(lista)
    print(f"Total de CNPJs no arquivo base: {total}")

    if total == 0:
        raise RuntimeError("Nenhum CNPJ encontrado em cnpjs.xlsx.")

    k = 3
    base = total // k
    resto = total % k

    partes = []
    inicio = 0
    for i in range(k):
        tamanho = base + (1 if i < resto else 0)
        fim = inicio + tamanho
        partes.append(lista[inicio:fim])
        inicio = fim

    for idx, parte in enumerate(partes, start=1):
        caminho = CNPJS_PARTES[idx - 1]
        df_parte = pd.DataFrame({"cnpj": parte})
        df_parte.to_excel(caminho, index=False)
        print(f"Parte {idx}: {len(parte)} CNPJs -> {caminho}")

    return partes


def rodar_workers():
    procs = []
    for i in range(3):
        entrada = CNPJS_PARTES[i]
        saida = REG_PARTES[i]
        worker_id = i + 1

        cmd = [
            sys.executable,
            os.path.join(BASE_DIR, "worker_mte.py"),
            "--entrada", entrada,
            "--saida", saida,
            "--id", str(worker_id),
        ]
        print(f"Iniciando worker {worker_id}: {cmd}")
        p = subprocess.Popen(cmd)
        procs.append(p)

    for i, p in enumerate(procs, start=1):
        ret = p.wait()
        print(f"Worker {i} terminou com codigo de retorno {ret}")


def merge_registros():
    dfs = []
    for caminho in REG_PARTES:
        if os.path.exists(caminho):
            try:
                df = pd.read_excel(caminho, dtype=str)
                if not df.empty:
                    dfs.append(df)
            except Exception:
                pass

    if not dfs:
        print("Nenhum registros_icr_parteX.xlsx encontrado ou todos vazios.")
        df_final = pd.DataFrame(columns=["codigo", "sindicatos", "data_inicio", "data_fim"])
        df_final.to_excel(REG_FINAL, index=False)
        return df_final

    df_final = pd.concat(dfs, ignore_index=True)
    for col in ["codigo", "sindicatos", "data_inicio", "data_fim"]:
        if col not in df_final.columns:
            df_final[col] = ""
    for col in df_final.columns:
        df_final[col] = df_final[col].astype(str)

    df_final.drop_duplicates(subset=["codigo"], keep="first", inplace=True)
    df_final.to_excel(REG_FINAL, index=False)
    print(f"Merge concluido. Registros unicos: {len(df_final)}")
    return df_final


def sincronizar_registros_e_pasta(df_registros: pd.DataFrame):
    try:
        arquivos = os.listdir(LIMPOS_DIR)
    except FileNotFoundError:
        arquivos = []

    codigos_pasta = set()
    for nome in arquivos:
        caminho = os.path.join(LIMPOS_DIR, nome)
        if not os.path.isfile(caminho):
            continue
        base, ext = os.path.splitext(nome)
        if not base or not ext:
            continue
        codigos_pasta.add(base)

    if df_registros.empty:
        print("Nenhum registro para sincronizar com a pasta de arquivos.")
        return df_registros

    codigos_reg = df_registros["codigo"].astype(str)
    chaves_reg = codigos_reg.map(codigo_para_nome_arquivo)

    mask_keep = chaves_reg.isin(codigos_pasta)
    removidos_registros = (~mask_keep).sum()
    df_registros_ok = df_registros[mask_keep].reset_index(drop=True)

    chaves_validas = set(chaves_reg[mask_keep])

    removidos_arquivos = 0
    for nome in arquivos:
        caminho = os.path.join(LIMPOS_DIR, nome)
        if not os.path.isfile(caminho):
            continue
        base, ext = os.path.splitext(nome)
        if not base or not ext:
            continue
        if base not in chaves_validas:
            try:
                os.remove(caminho)
                removidos_arquivos += 1
            except OSError:
                pass

    df_registros_ok.to_excel(REG_FINAL, index=False)

    print(
        f"Sincronizacao final: {removidos_registros} registro(s) sem arquivo removido(s), "
        f"{removidos_arquivos} arquivo(s) sem registro apagado(s)."
    )

    return df_registros_ok


def registrar_execucao(df_registros_final: pd.DataFrame):
    data_exec = date.today().isoformat()
    total_codigos = len(df_registros_final)

    try:
        arquivos = [f for f in os.listdir(LIMPOS_DIR)
                    if os.path.isfile(os.path.join(LIMPOS_DIR, f))]
        total_arquivos = len(arquivos)
    except FileNotFoundError:
        total_arquivos = 0

    linha = {
        "data_execucao": data_exec,
        "total_registros": total_codigos,
        "total_arquivos": total_arquivos,
    }

    if os.path.exists(HIST_EXECUCOES):
        df_hist = pd.read_excel(HIST_EXECUCOES, dtype=str)
    else:
        df_hist = pd.DataFrame(columns=["data_execucao", "total_registros", "total_arquivos"])

    df_hist = pd.concat([df_hist, pd.DataFrame([linha])], ignore_index=True)
    df_hist.to_excel(HIST_EXECUCOES, index=False)

    print(
        f"Execucao registrada em {HIST_EXECUCOES}: data={data_exec}, "
        f"registros={total_codigos}, arquivos={total_arquivos}"
    )


def main():
    print("Dividindo cnpjs.xlsx em 3 partes...")
    dividir_cnpjs_em_3()

    print("\nIniciando workers em paralelo...")
    rodar_workers()

    print("\nFazendo merge dos registros das 3 partes...")
    df_final = merge_registros()

    print("\nSincronizando registros com a pasta de arquivos...")
    df_final = sincronizar_registros_e_pasta(df_final)

    print("\nRegistrando data da execucao...")
    registrar_execucao(df_final)

    print("\nProcesso completo concluido.")


if __name__ == "__main__":
    main()

print("\nExecutando pós-processamento ChatGPT...")
subprocess.run([
    sys.executable,
    os.path.join(BASE_DIR, "posprocessar_chatgpt.py")
])
