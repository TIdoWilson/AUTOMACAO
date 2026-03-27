import fitz  # PyMuPDF
from tkinter import Tk, filedialog
import re


def escolher_pdf():
    root = Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    root.destroy()
    return caminho


def rgb01_para_255(c):
    return tuple(int(round(x * 255)) for x in c)


def extrair_valores(texto):
    """
    Procura valores em formato brasileiro, cobrindo:
      126,00C
      640,00D
      -126,00
      1.234,56
      -1.234,56D

    Regra base:
      - Sufixo 'C' => tipo_base = 'C'
      - Sufixo 'D' => tipo_base = 'D'
      - Sinal '-'  => tipo_base = 'D'
      - Sem C/D e sem '-' => tipo_base = 'N' (neutro)

    Retorna lista de dicts: {valor: float, tipo_base: 'C'/'D'/'N'}
    """
    if not texto:
        return []

    padrao = r'([+-]?)\s*(\d{1,3}(?:\.\d{3})*,\d{2})\s*([CD])?'
    resultados = []

    for m in re.finditer(padrao, texto):
        sinal = m.group(1) or ""
        num_str = m.group(2)
        letra = m.group(3)  # 'C', 'D' ou None

        limpa = num_str.replace('.', '').replace(',', '.')
        try:
            valor = float(limpa)
        except ValueError:
            continue

        # Decide tipo_base
        if letra == "C":
            tipo_base = "C"
        elif letra == "D":
            tipo_base = "D"
        elif sinal == "-":
            tipo_base = "D"
        else:
            tipo_base = "N"  # neutro (nem C/D nem '-')

        resultados.append({"valor": valor, "tipo_base": tipo_base})

    return resultados


def analisar_anotacoes(caminho_pdf):
    doc = fitz.open(caminho_pdf)

    # Tipos de marcação de texto (text markup) que queremos considerar
    text_markup_types = {"Highlight", "Underline", "Squiggly", "StrikeOut"}

    # dicionário: cor (tuple 0–1) -> infos
    # info = { "qtd": int, "tipos": set, "paginas": set, "annots": [ {pagina, tipo, texto} ] }
    por_cor = {}

    print("\n=== Anotações de texto encontradas no PDF ===\n")

    for num_pagina in range(len(doc)):
        pagina = doc[num_pagina]
        for anot in pagina.annots() or []:
            tipo_codigo, tipo_nome = anot.type  # ex.: (8, 'Highlight')

            # Pega cor
            cor = None
            try:
                cores = anot.colors  # dict com 'stroke'/'fill'
                if cores:
                    if cores.get("stroke") is not None:
                        cor = cores["stroke"]
                    elif cores.get("fill") is not None:
                        cor = cores["fill"]
            except Exception:
                pass

            if cor is None:
                try:
                    cor = anot.color
                except Exception:
                    pass

            # Tenta pegar o texto dentro da anotação (útil para extrato)
            try:
                texto = pagina.get_textbox(anot.rect).strip()
            except Exception:
                texto = ""

            if cor is not None:
                # normaliza para tuple (e arredonda um pouco pra evitar ruído de float)
                cor = tuple(round(float(x), 4) for x in cor)
                cor_255 = rgb01_para_255(cor)
                print(
                    f"Página {num_pagina + 1:>3} | "
                    f"Tipo: {tipo_nome:<10} | "
                    f"Cor 0–1: {cor} | Cor 0–255 aprox.: {cor_255} | "
                    f"Texto: {texto}"
                )
            else:
                print(
                    f"Página {num_pagina + 1:>3} | "
                    f"Tipo: {tipo_nome:<10} | "
                    "Cor: (nenhuma/indefinida)"
                )

            # Só contabiliza os text markups com cor
            if tipo_nome in text_markup_types and cor is not None:
                info = por_cor.setdefault(
                    cor,
                    {"qtd": 0, "tipos": set(), "paginas": set(), "annots": []},
                )
                info["qtd"] += 1
                info["tipos"].add(tipo_nome)
                info["paginas"].add(num_pagina + 1)
                info["annots"].append(
                    {
                        "pagina": num_pagina + 1,
                        "tipo": tipo_nome,
                        "texto": texto,
                    }
                )

    doc.close()
    return por_cor


def mostrar_resumo(por_cor):
    if not por_cor:
        print("\nNenhuma anotação de texto (Highlight/Underline/Squiggly/StrikeOut) encontrada.")
        return []

    print("\n=== Resumo de destaques por cor ===")
    total_geral = 0

    cores_ordenadas = sorted(por_cor.keys())

    for idx, cor in enumerate(cores_ordenadas, start=1):
        info = por_cor[cor]
        rgb_255 = rgb01_para_255(cor)
        qtd = info["qtd"]
        tipos = ", ".join(sorted(info["tipos"]))
        paginas = ", ".join(str(p) for p in sorted(info["paginas"]))
        total_geral += qtd

        print("-" * 70)
        print(f"[{idx}] Cor 0–255: {rgb_255}")
        print(f"    Cor 0–1:   {cor}")
        print(f"    Tipos:     {tipos}")
        print(f"    Quantidade de anotações: {qtd}")
        print(f"    Páginas:   {paginas}")

    print("-" * 70)
    print(f"\nTotal geral de anotações de texto no PDF: {total_geral}\n")

    return cores_ordenadas


def perguntar_tratamento_neutros(descricao_cor):
    """
    Pergunta ao usuário como tratar valores neutros (tipo_base='N') para uma cor.
    Retorna 'C', 'D' ou 'I' (ignorar).
    """
    print(f"\nPara a cor {descricao_cor}, como tratar valores SEM C/D e SEM sinal '-'?")
    print("  C - Tratar como CRÉDITO")
    print("  D - Tratar como DÉBITO")
    print("  I - IGNORAR esses valores")
    resp = input("Escolha (C/D/I, padrão C): ").strip().upper()
    if resp not in ("C", "D", "I"):
        resp = "C"
    return resp


def somar_valores_para_cor(info_cor, trat_neutros):
    """
    Recebe o dict info_cor (para UMA cor) e soma:
      - créditos (C)
      - débitos (D)

    trat_neutros: 'C', 'D' ou 'I' (ignorar) define
    o destino dos valores com tipo_base 'N'.

    Retorna (total_credito, total_debito, lista_detalhada)
    onde cada item da lista tem 'tipo_final' (C/D/-).
    """
    valores_encontrados = []
    total_credito = 0.0
    total_debito = 0.0

    for ann in info_cor["annots"]:
        texto = ann["texto"]
        valores = extrair_valores(texto)
        for v in valores:
            valor = v["valor"]
            tipo_base = v["tipo_base"]  # 'C', 'D' ou 'N'

            if tipo_base == "C":
                tipo_final = "C"
            elif tipo_base == "D":
                tipo_final = "D"
            else:  # 'N'
                if trat_neutros == "C":
                    tipo_final = "C"
                elif trat_neutros == "D":
                    tipo_final = "D"
                else:  # 'I'
                    tipo_final = "-"

            item = {
                "pagina": ann["pagina"],
                "texto": texto,
                "valor": valor,
                "tipo_base": tipo_base,
                "tipo_final": tipo_final,
            }
            valores_encontrados.append(item)

            if tipo_final == "C":
                total_credito += valor
            elif tipo_final == "D":
                total_debito += valor
            # tipo_final '-' é ignorado

    return total_credito, total_debito, valores_encontrados


def somar_valores_todas_cores(por_cor, tratamentos_por_cor):
    """
    Soma valores de TODAS as cores, usando o tratamento
    de neutros definido para cada cor.
    Retorna (total_credito, total_debito, lista_detalhada).
    """
    total_credito = 0.0
    total_debito = 0.0
    detalhes = []

    for cor, info_cor in por_cor.items():
        trat_neutros = tratamentos_por_cor.get(cor, "C")
        c, d, lista = somar_valores_para_cor(info_cor, trat_neutros)
        total_credito += c
        total_debito += d
        # adiciona a informação da cor em cada item para referência
        for item in lista:
            item_com_cor = dict(item)
            item_com_cor["cor"] = rgb01_para_255(cor)
            detalhes.append(item_com_cor)

    return total_credito, total_debito, detalhes


def imprimir_resultado_soma(total_credito, total_debito, valores_encontrados):
    if not valores_encontrados:
        print("Nenhum valor numérico no formato esperado foi encontrado nos destaques selecionados.")
        return

    print("Valores encontrados:")
    for item in valores_encontrados:
        tipo_final = item.get("tipo_final", "-")
        tipo_base = item.get("tipo_base", "-")
        cor_str = ""
        if "cor" in item:
            cor_str = f" | Cor 0–255: {item['cor']}"
        print(
            f"  Página {item['pagina']}: base={tipo_base} -> final={tipo_final} | "
            f"texto = '{item['texto']}' -> valor = {item['valor']:.2f}{cor_str}"
        )

    print("\nResumo:")
    print(f"  Total de créditos (C): {total_credito:.2f}")
    print(f"  Total de débitos (D):  {total_debito:.2f}")
    print(f"  Resultado (C - D):     {total_credito - total_debito:.2f}\n")


def main():
    caminho_pdf = escolher_pdf()
    if not caminho_pdf:
        print("Nenhum PDF selecionado. Saindo...")
        return

    print(f"\nArquivo selecionado:\n  {caminho_pdf}\n")

    por_cor = analisar_anotacoes(caminho_pdf)
    cores_ordenadas = mostrar_resumo(por_cor)

    if not cores_ordenadas:
        return

    # Pergunta se o usuário quer somar valores
    resp = input("Deseja somar valores numéricos dos textos destacados? (s/N): ").strip().lower()
    if resp != "s":
        # Comportamento antigo: apenas contagem por cor
        return

    # Escolher modo: uma cor específica ou todas
    print("\nComo deseja somar?")
    print("  1 - Somar apenas uma cor (ex.: só verde, só amarelo etc.)")
    print("  2 - Somar todas as cores juntas")
    modo = input("Escolha 1 ou 2 (padrão 1): ").strip()
    if modo not in ("1", "2"):
        modo = "1"

    if modo == "1":
        # Escolhe a cor pelo índice mostrado no resumo
        while True:
            escolha = input("Digite o número da cor que deseja somar (ex.: 1): ").strip()
            if not escolha.isdigit():
                print("Digite um número válido.")
                continue
            idx = int(escolha)
            if not (1 <= idx <= len(cores_ordenadas)):
                print("Número fora do intervalo. Tente novamente.")
                continue
            break

        cor_escolhida = cores_ordenadas[idx - 1]
        rgb_255 = rgb01_para_255(cor_escolhida)
        desc_cor = f"[{idx}] RGB 0–255 {rgb_255}"

        # Pergunta como tratar neutros para ESSA cor
        trat_neutros = perguntar_tratamento_neutros(desc_cor)

        print(f"\nSomando valores para a cor {idx} (RGB 0–255: {rgb_255})...\n")

        info_cor = por_cor[cor_escolhida]
        total_credito, total_debito, lista = somar_valores_para_cor(info_cor, trat_neutros)
        imprimir_resultado_soma(total_credito, total_debito, lista)

    else:
        # Somar todas as cores: perguntar tratamento de neutros para cada uma
        tratamentos_por_cor = {}
        for idx, cor in enumerate(cores_ordenadas, start=1):
            rgb_255 = rgb01_para_255(cor)
            desc_cor = f"[{idx}] RGB 0–255 {rgb_255}"
            trat = perguntar_tratamento_neutros(desc_cor)
            tratamentos_por_cor[cor] = trat

        print("\nSomando valores de TODAS as cores...\n")
        total_credito, total_debito, lista = somar_valores_todas_cores(por_cor, tratamentos_por_cor)
        imprimir_resultado_soma(total_credito, total_debito, lista)


if __name__ == "__main__":
    main()
