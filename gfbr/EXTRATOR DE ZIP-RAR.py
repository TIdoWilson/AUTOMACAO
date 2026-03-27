#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import zipfile
from pathlib import Path
from io import BytesIO
from tkinter import Tk, filedialog  # <-- adicionado

# Tente importar rarfile se existir
try:
    import rarfile
    RAR_AVAILABLE = True
except Exception:
    RAR_AVAILABLE = False

MAX_DEPTH = 5

def ask_base_dir() -> Path:
    """Abre janela para o usuário selecionar o diretório."""
    root = Tk()
    root.withdraw()  # esconde a janela principal
    print("Selecione a pasta onde estão os arquivos .zip/.rar:")
    base = filedialog.askdirectory(title="Selecione a pasta com os arquivos .zip/.rar")
    root.destroy()

    if not base:
        print('Nenhum diretório selecionado. Saindo.')
        sys.exit(1)

    p = Path(base)
    if not p.exists() or not p.is_dir():
        print(f'Caminho inválido: {p}')
        sys.exit(1)
    return p

def ensure_ARQUIVOS_dir(base_dir: Path) -> Path:
    ARQUIVOS = base_dir / 'ARQUIVOS'
    ARQUIVOS.mkdir(exist_ok=True)
    return ARQUIVOS

def is_archive_filename(name: str) -> bool:
    low = name.lower()
    return low.endswith('.zip') or low.endswith('.rar')

def unique_name_for_size(dest_dir: Path, base_name: str, size: int, used_sizes_for_name: dict) -> Path:
    name = base_name
    stem = Path(base_name).stem
    suffix = Path(base_name).suffix

    sizes = used_sizes_for_name.setdefault(base_name.lower(), [])
    if size in sizes:
        return dest_dir / name
    else:
        if len(sizes) == 0 and not (dest_dir / name).exists():
            return dest_dir / name
        else:
            idx = 2
            while True:
                cand = f"{stem} ({idx}){suffix}"
                cand_path = dest_dir / cand
                if cand_path.exists():
                    if cand_path.stat().st_size == size:
                        return cand_path
                    idx += 1
                else:
                    return cand_path

def add_record_for_written(filepath: Path, base_name: str, size: int, used_sizes_for_name: dict):
    sizes = used_sizes_for_name.setdefault(base_name.lower(), [])
    if size not in sizes:
        sizes.append(size)

def save_file(member_name: str, data: bytes, dest_dir: Path, used_sizes_for_name: dict):
    base_name = Path(member_name).name
    size = len(data)
    out_path = unique_name_for_size(dest_dir, base_name, size, used_sizes_for_name)
    if out_path.exists() and out_path.stat().st_size == size:
        return
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(data)
    add_record_for_written(out_path, base_name, size, used_sizes_for_name)
    print(f'  [+] {out_path.name}  ({size} bytes)')

def extract_zipfile(zf: zipfile.ZipFile, dest_dir: Path, used_sizes_for_name: dict, depth: int):
    for member in zf.infolist():
        if member.is_dir():
            continue
        member_name = Path(member.filename).name
        if not member_name or member_name in ('.', '..'):
            continue
        try:
            data = zf.read(member)
        except RuntimeError as e:
            print(f'  [ERRO] Zip protegido por senha ou inválido: {member.filename} ({e})')
            continue
        except Exception as e:
            print(f'  [ERRO] Falha ao ler "{member.filename}": {e}')
            continue

        if depth < MAX_DEPTH and is_archive_filename(member_name):
            print(f'    [*] Encontrado compactado interno: {member_name} (profundidade {depth+1})')
            process_archive_bytes(member_name, data, dest_dir, used_sizes_for_name, depth+1)
        else:
            save_file(member_name, data, dest_dir, used_sizes_for_name)

def extract_rarfile(rf, dest_dir: Path, used_sizes_for_name: dict, depth: int):
    for member in rf.infolist():
        if getattr(member, 'is_dir', lambda: False)():
            continue
        member_name = Path(member.filename).name
        if not member_name or member_name in ('.', '..'):
            continue
        try:
            data = rf.read(member)
        except rarfile.NeedFirstVolume:
            print(f'  [ERRO] Parte inicial de multi-volume RAR ausente: {member.filename}')
            continue
        except rarfile.BadRarFile as e:
            print(f'  [ERRO] RAR inválido: {member.filename} ({e})')
            continue
        except Exception as e:
            print(f'  [ERRO] Falha ao ler "{member.filename}": {e}')
            continue

        if depth < MAX_DEPTH and is_archive_filename(member_name):
            print(f'    [*] Encontrado compactado interno: {member_name} (profundidade {depth+1})')
            process_archive_bytes(member_name, data, dest_dir, used_sizes_for_name, depth+1)
        else:
            save_file(member_name, data, dest_dir, used_sizes_for_name)

def process_zip_path(zip_path: Path, dest_dir: Path, used_sizes_for_name: dict, depth: int = 0):
    print(f'\n[ZIP] {zip_path.name}')
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            extract_zipfile(zf, dest_dir, used_sizes_for_name, depth)
    except zipfile.BadZipFile:
        print('  [ERRO] Arquivo .zip corrompido ou inválido.')
    except Exception as e:
        print(f'  [ERRO] {e}')

def process_rar_path(rar_path: Path, dest_dir: Path, used_sizes_for_name: dict, depth: int = 0):
    print(f'\n[RAR] {rar_path.name}')
    if not RAR_AVAILABLE:
        print('  [AVISO] Suporte a .rar indisponível. Instale: pip install rarfile e o utilitário "unrar" ou "bsdtar".')
        return
    try:
        with rarfile.RarFile(rar_path, 'r') as rf:
            extract_rarfile(rf, dest_dir, used_sizes_for_name, depth)
    except rarfile.Error as e:
        print(f'  [ERRO] {e}')
    except Exception as e:
        print(f'  [ERRO] {e}')

def process_archive_bytes(name: str, data: bytes, dest_dir: Path, used_sizes_for_name: dict, depth: int):
    ext = Path(name).suffix.lower()
    bio = BytesIO(data)
    if ext == '.zip':
        try:
            with zipfile.ZipFile(bio, 'r') as zf:
                extract_zipfile(zf, dest_dir, used_sizes_for_name, depth)
        except zipfile.BadZipFile:
            print(f'  [ERRO] Compactado interno .zip corrompido: {name}')
    elif ext == '.rar':
        if not RAR_AVAILABLE:
            print(f'  [AVISO] Encontrado RAR interno "{name}", mas suporte .rar não está disponível.')
            return
        try:
            with rarfile.RarFile(fileobj=bio) as rf:
                extract_rarfile(rf, dest_dir, used_sizes_for_name, depth)
        except rarfile.Error as e:
            print(f'  [ERRO] RAR interno inválido "{name}": {e}')

def main():
    base_dir = ask_base_dir()
    dest_dir = ensure_ARQUIVOS_dir(base_dir)
    used_sizes_for_name = {}

    for existing in dest_dir.glob('*'):
        if existing.is_file():
            base_name = existing.name
            size = existing.stat().st_size
            add_record_for_written(existing, base_name, size, used_sizes_for_name)

    archives = []
    for p in base_dir.iterdir():
        if p.is_file() and p.suffix.lower() in ('.zip', '.rar'):
            archives.append(p)

    if not archives:
        print('Nenhum .zip ou .rar encontrado na pasta selecionada.')
        return

    print(f'\nEncontrados {len(archives)} arquivo(s) compactado(s). Extraindo (recursivo, até {MAX_DEPTH} níveis) para: {dest_dir}')
    for a in archives:
        if a.suffix.lower() == '.zip':
            process_zip_path(a, dest_dir, used_sizes_for_name, depth=0)
        elif a.suffix.lower() == '.rar':
            process_rar_path(a, dest_dir, used_sizes_for_name, depth=0)

    print('\nConcluído.')

if __name__ == '__main__':
    main()
