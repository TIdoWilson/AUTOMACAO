import os
import re
from pathlib import Path

# Pastas de origem e destino
PASTA_ORIGEM = r"C:\Users\Usuario\Desktop\CGT TOSCAN"
PASTA_DESTINO = r"C:\Users\Usuario\Desktop\CGT TOSCAN\LALUR"

# Cria a pasta de destino se não existir
Path(PASTA_DESTINO).mkdir(parents=True, exist_ok=True)

# Função para limpar nomes inválidos em arquivos do Windows
def sanitizar_nome(nome: str) -> str:
    # remove caracteres não permitidos e aparas espaços/pontos finais
    nome = re.sub(r'[<>:"/\\|?*]', '_', nome)
    nome = nome.strip().rstrip(".")
    # limita a ~150 chars para evitar problemas com comprimento de caminho
    return nome[:150] if len(nome) > 150 else nome

def listar_arquivos_excel(pasta: str):
    exts = (".xlsx", ".xlsm")
    for item in Path(pasta).iterdir():
        if item.is_file() and item.suffix.lower() in exts:
            yield item.resolve()

def exportar_abas_para_pdf(caminhos_arquivos):
    import win32com.client as win32

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for caminho in caminhos_arquivos:
            wb = None
            try:
                wb = excel.Workbooks.Open(str(caminho), ReadOnly=True)
                nome_wb = sanitizar_nome(caminho.stem)

                # Itera apenas por Worksheets (abas de planilha)
                for sh in wb.Worksheets:
                    try:
                        # Define área de impressão A1:I32
                        sh.PageSetup.PrintArea = "$A$1:$I$32"

                        # Ajuste para caber em 1 página (opcional, facilita leitura)
                        sh.PageSetup.Zoom = False
                        sh.PageSetup.FitToPagesWide = 1
                        sh.PageSetup.FitToPagesTall = 1

                        # Monta o nome do PDF: "<Arquivo> - <Aba>.pdf"
                        nome_aba = sanitizar_nome(sh.Name)
                        nome_pdf = f"{nome_wb} - {nome_aba}.pdf"
                        caminho_pdf = str(Path(PASTA_DESTINO) / nome_pdf)

                        # Exporta a aba para PDF respeitando a área definida
                        sh.ExportAsFixedFormat(
                            Type=0,  # xlTypePDF
                            Filename=caminho_pdf,
                            Quality=0,  # xlQualityStandard
                            IncludeDocProperties=True,
                            IgnorePrintAreas=False,
                            OpenAfterPublish=False
                        )
                        print(f"Gerado: {caminho_pdf}")
                    except Exception as e:
                        print(f"[ERRO] Falha ao exportar aba '{sh.Name}' de '{caminho.name}': {e}")

            except Exception as e:
                print(f"[ERRO] Não foi possível abrir '{caminho}': {e}")
            finally:
                if wb is not None:
                    # Fecha sem salvar alterações
                    wb.Close(SaveChanges=False)

    finally:
        # Encerra o Excel
        excel.Quit()

if __name__ == "__main__":
    arquivos = list(listar_arquivos_excel(PASTA_ORIGEM))
    if not arquivos:
        print("Nenhum arquivo .xlsx ou .xlsm encontrado na pasta de origem.")
    else:
        exportar_abas_para_pdf(arquivos)
        print("Concluído.")
