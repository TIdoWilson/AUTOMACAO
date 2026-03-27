from pathlib import Path
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation

# Para escolher pasta/arquivo em janela
import tkinter as tk
from tkinter import filedialog, messagebox

# Namespace NFCom
NS = {"n": "http://www.portalfiscal.inf.br/nfcom"}

# Configuração básica do contribuinte (ajuste se precisar)
IND_OPER_PADRAO = "1"  # 0 = Entrada, 1 = Saída
IND_EMIT_PADRAO = "0"  # 0 = Emissão própria, 1 = Terceiros


def to_decimal(text):
    """Converte texto para Decimal (aceita ponto ou vírgula)."""
    if text is None or text == "":
        return None
    s = str(text).strip().replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def format_decimal(val, dec=2):
    """
    Formata Decimal no padrão brasileiro:
    - Usa vírgula como separador
    - Mantém '0' como '0'
    """
    if val is None:
        return ""
    if not isinstance(val, Decimal):
        val = Decimal(str(val))
    if val == 0:
        return "0"

    quant = Decimal("1").scaleb(-dec)  # 10^-dec
    val = val.quantize(quant)

    s = f"{val:f}"  # sem notação científica
    if "." in s:
        int_part, frac = s.split(".", 1)
    else:
        int_part, frac = s, ""
    if dec > 0:
        frac = (frac + "0" * dec)[:dec]
        s = int_part + "," + frac
    else:
        s = int_part
    return s


def zero_pad(value, length):
    """Preenche com zeros à esquerda até o tamanho desejado."""
    if value is None:
        return ""
    s = str(value).strip()
    return s.zfill(length)


def iso_to_ddmmaaaa(iso):
    """Converte data ISO 2025-11-01T... para 01112025."""
    if not iso:
        return ""
    date_part = iso.split("T", 1)[0]  # YYYY-MM-DD
    y, m, d = date_part.split("-")
    return f"{d}{m}{y}"


def map_cod_sit(cstat):
    """
    Mapeia cStat da autorização NFCom para COD_SIT do SPED (tabela 4.1.2, simplificado).
    Ajuste se precisar de mais situações.
    """
    if cstat is None:
        return "00"
    cstat = str(cstat)
    if cstat in {"100", "150"}:
        return "00"  # documento regular
    if cstat in {"101", "151"}:
        return "02"  # documento cancelado
    if cstat in {"135", "136"}:
        return "06"  # documento complementar
    return "00"      # fallback


def process_nfcom_xml(xml_content):
    """Gera D700, D730 e D731 para um único XML NFCom."""
    root = ET.fromstring(xml_content)

    ide = root.find(".//n:ide", NS)
    total = root.find(".//n:total", NS)
    dest = root.find(".//n:dest", NS)
    ender_dest = dest.find("n:enderDest", NS) if dest is not None else None
    icms_tot = total.find("n:ICMSTot", NS) if total is not None else None
    prot = root.find(".//n:protNFCom", NS)
    inf_prot = prot.find("n:infProt", NS) if prot is not None else None

    # Campos básicos do cabeçalho (D700)
    mod = (
        ide.find("n:mod", NS).text
        if ide is not None and ide.find("n:mod", NS) is not None
        else ""
    )
    serie = (
        ide.find("n:serie", NS).text
        if ide is not None and ide.find("n:serie", NS) is not None
        else ""
    )
    n_nf = (
        ide.find("n:nNF", NS).text
        if ide is not None and ide.find("n:nNF", NS) is not None
        else ""
    )
    dh_emi = (
        ide.find("n:dhEmi", NS).text
        if ide is not None and ide.find("n:dhEmi", NS) is not None
        else ""
    )
    fin_nfcom = (
        ide.find("n:finNFCom", NS).text
        if ide is not None and ide.find("n:finNFCom", NS) is not None
        else ""
    )
    tp_fat = (
        ide.find("n:tpFat", NS).text
        if ide is not None and ide.find("n:tpFat", NS) is not None
        else ""
    )

    # Município do destinatário (D700.COD_MUN_DEST)
    cmun_dest = None
    if ender_dest is not None:
        cmun_el = ender_dest.find("n:cMun", NS)
        if cmun_el is not None:
            cmun_dest = cmun_el.text

    if not cmun_dest and ide is not None:
        cmun_fg = ide.find("n:cMunFG", NS)
        if cmun_fg is not None:
            cmun_dest = cmun_fg.text

    # Chave NFCom + situação
    ch_nfcom = None
    c_stat = None
    if inf_prot is not None:
        ch_el = inf_prot.find("n:chNFCom", NS)
        if ch_el is not None:
            ch_nfcom = ch_el.text
        cstat_el = inf_prot.find("n:cStat", NS)
        if cstat_el is not None:
            c_stat = cstat_el.text

    # Totais
    v_nf = (
        to_decimal(total.find("n:vNF", NS).text)
        if total is not None and total.find("n:vNF", NS) is not None
        else None
    )
    v_desc = (
        to_decimal(total.find("n:vDesc", NS).text)
        if total is not None and total.find("n:vDesc", NS) is not None
        else None
    )
    v_bc = (
        to_decimal(icms_tot.find("n:vBC", NS).text)
        if icms_tot is not None and icms_tot.find("n:vBC", NS) is not None
        else None
    )
    v_icms = (
        to_decimal(icms_tot.find("n:vICMS", NS).text)
        if icms_tot is not None and icms_tot.find("n:vICMS", NS) is not None
        else None
    )
    v_fcp_tot = (
        to_decimal(icms_tot.find("n:vFCP", NS).text)
        if icms_tot is not None and icms_tot.find("n:vFCP", NS) is not None
        else Decimal("0")
    )

    # --- Agregação dos itens para D730/D731 (por CST, CFOP, alíquota) ---
    agregados = {}

    for det in root.findall(".//n:det", NS):
        prod = det.find("n:prod", NS)
        imposto = det.find("n:imposto", NS)
        if prod is None or imposto is None:
            continue

        # CFOP + valor da operação (vProd)
        cfop_el = prod.find("n:CFOP", NS)
        cfop = cfop_el.text.strip() if cfop_el is not None and cfop_el.text else ""
        v_prod = (
            to_decimal(prod.find("n:vProd", NS).text)
            if prod.find("n:vProd", NS) is not None
            else Decimal("0")
        )

        # Primeiro grupo ICMS* encontrado (ICMS00, ICMS20, etc.)
        icms_grp = None
        for child in imposto:
            if child.tag.startswith(f"{{{NS['n']}}}ICMS"):
                icms_grp = child
                break
        if icms_grp is None:
            continue

        cst_el = icms_grp.find("n:CST", NS)
        cst = cst_el.text.strip() if cst_el is not None and cst_el.text else ""

        p_icms = (
            to_decimal(icms_grp.find("n:pICMS", NS).text)
            if icms_grp.find("n:pICMS", NS) is not None
            else None
        )
        v_bc_item = (
            to_decimal(icms_grp.find("n:vBC", NS).text)
            if icms_grp.find("n:vBC", NS) is not None
            else Decimal("0")
        )
        v_icms_item = (
            to_decimal(icms_grp.find("n:vICMS", NS).text)
            if icms_grp.find("n:vICMS", NS) is not None
            else Decimal("0")
        )
        v_fcp_item = (
            to_decimal(icms_grp.find("n:vFCP", NS).text)
            if icms_grp.find("n:vFCP", NS) is not None
            else Decimal("0")
        )

        chave = (cst, cfop, p_icms)
        if chave not in agregados:
            agregados[chave] = {
                "vl_opr": Decimal("0"),
                "vl_bc": Decimal("0"),
                "vl_icms": Decimal("0"),
                "vl_red_bc": Decimal("0"),
                "vl_fcp": Decimal("0"),
            }

        rec = agregados[chave]
        rec["vl_opr"] += v_prod or Decimal("0")
        rec["vl_bc"] += v_bc_item or Decimal("0")
        rec["vl_icms"] += v_icms_item or Decimal("0")
        rec["vl_fcp"] += v_fcp_item or Decimal("0")

    # Se por algum motivo não agregou nada, usa os totais da nota
    if not agregados and v_bc is not None and v_icms is not None:
        chave = ("00", "", None)
        agregados[chave] = {
            "vl_opr": v_nf or Decimal("0"),
            "vl_bc": v_bc or Decimal("0"),
            "vl_icms": v_icms or Decimal("0"),
            "vl_red_bc": Decimal("0"),
            "vl_fcp": v_fcp_tot or Decimal("0"),
        }

    # --------- D700 ---------
    d700_campos = [
        "D700",                      # REG
        IND_OPER_PADRAO,            # IND_OPER
        IND_EMIT_PADRAO,            # IND_EMIT
        "",                         # COD_PART (usado só em entradas)
        mod,                        # COD_MOD (62)
        map_cod_sit(c_stat),        # COD_SIT
        serie,                      # SER
        n_nf,                       # NUM_DOC
        iso_to_ddmmaaaa(dh_emi),    # DT_DOC
        "",                         # DT_E_S
        format_decimal(v_nf, 2),    # VL_DOC
        format_decimal(v_desc, 2) if v_desc is not None else "0",  # VL_DESC
        format_decimal(v_bc, 2),    # VL_SERV (aqui usei BC do ICMS)
        "0",                        # VL_SERV_NT
        "0",                        # VL_TERC
        "0",                        # VL_DA
        format_decimal(v_bc, 2),    # VL_BC_ICMS
        format_decimal(v_icms, 2),  # VL_ICMS
        "",                         # COD_INF
        "",                         # VL_PIS
        "",                         # VL_COFINS
        ch_nfcom or "",             # CHV_DOCe
        fin_nfcom or "",            # FIN_DOCe
        tp_fat or "",               # TIP_FAT
        "", "", "", "", "", "",     # campos de referência vazios
        cmun_dest or "",            # COD_MUN_DEST
        "",                         # DED
    ]
    d700_linha = "|" + "|".join(d700_campos) + "|"

    # --------- D730 / D731 ---------
    d730_linhas = []
    d731_linhas = []

    for (cst, cfop, p_icms), vals in sorted(
        agregados.items(), key=lambda x: (x[0][1], x[0][0])
    ):
        d730_campos = [
            "D730",
            zero_pad(cst, 3),                   # CST_ICMS
            zero_pad(cfop, 4),                  # CFOP
            format_decimal(p_icms, 2),          # ALIQ_ICMS
            format_decimal(vals["vl_opr"], 2),  # VL_OPR
            format_decimal(vals["vl_bc"], 2),   # VL_BC_ICMS
            format_decimal(vals["vl_icms"], 2), # VL_ICMS
            format_decimal(vals["vl_red_bc"], 2),  # VL_RED_BC
            "",                                 # COD_OBS
        ]
        d730_linhas.append("|" + "|".join(d730_campos) + "|")

        d731_campos = [
            "D731",
            format_decimal(vals["vl_fcp"], 2),  # VL_FCP_OP
        ]
        d731_linhas.append("|" + "|".join(d731_campos) + "|")

    return d700_linha, d730_linhas, d731_linhas


def processar_pasta(pasta_xml, arquivo_saida):
    """Percorre a pasta, gera D700/D730/D731 para todos os XML e grava em um TXT."""
    pasta = Path(pasta_xml)
    linhas = []

    for xml_path in sorted(pasta.glob("*.xml")):
        with open(xml_path, "r", encoding="utf-8") as f:
            conteudo = f.read()
        d700, d730_list, d731_list = process_nfcom_xml(conteudo)
        linhas.append(d700)
        linhas.extend(d730_list)
        linhas.extend(d731_list)

    if not linhas:
        raise RuntimeError("Nenhum XML processado na pasta selecionada.")

    with open(arquivo_saida, "w", encoding="utf-8") as f_out:
        f_out.write("\n".join(linhas))


if __name__ == "__main__":
    # Inicializa Tkinter (sem abrir janela principal)
    root = tk.Tk()
    root.withdraw()

    # Selecionar pasta com XML
    pasta_xml = filedialog.askdirectory(
        title="Selecione a pasta com os XML das NFCom"
    )
    if not pasta_xml:
        print("Nenhuma pasta selecionada. Encerrando.")
        raise SystemExit

    # Selecionar arquivo de saída (TXT)
    arquivo_saida = filedialog.asksaveasfilename(
        title="Selecione o arquivo TXT de saída",
        defaultextension=".txt",
        filetypes=[("Arquivo texto", "*.txt"), ("Todos os arquivos", "*.*")],
    )
    if not arquivo_saida:
        print("Nenhum arquivo de saída selecionado. Encerrando.")
        raise SystemExit

    try:
        processar_pasta(pasta_xml, arquivo_saida)
        messagebox.showinfo(
            "Concluído",
            f"Arquivo gerado com sucesso:\n{arquivo_saida}",
        )
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")
        raise
