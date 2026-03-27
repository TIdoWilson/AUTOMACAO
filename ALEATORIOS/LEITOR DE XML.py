import xml.etree.ElementTree as ET
from tkinter import Tk, filedialog

def selecionar_arquivo():
    Tk().withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo XML",
        filetypes=[("Arquivos XML", "*.xml")]
    )
    return caminho

def ler_nr_documento(caminho_xml):
    # Namespace do XML
    ns = {'ns': 'https://www.esnfs.com.br/xsd'}

    tree = ET.parse(caminho_xml)
    root = tree.getroot()

    lista = []

    # Encontra todas as tags <nrDocumento> dentro do namespace
    for tag in root.findall('.//ns:nrDocumento', ns):
        if tag.text:
            lista.append(tag.text.strip())

    return lista


if __name__ == "__main__":
    arquivo = selecionar_arquivo()

    if not arquivo:
        print("Nenhum arquivo selecionado.")
        exit()

    documentos = ler_nr_documento(arquivo)

    print("\n=== LISTA DE <nrDocumento> ===")
    for d in documentos:
        print(d)

    print("\nTotal encontrado:", len(documentos))
