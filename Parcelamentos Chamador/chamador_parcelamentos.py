#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Chamador local de consultores de parcelamentos.

- Le arquivo JSON exportado pelo portal (parcelamentos.import.json)
- Monta checklist por tipo/empresa
- Dispara consultores por tipo (batch) ou por item (per_item)
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

ROOT_DIR = Path(r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA")
DEFAULT_IMPORT_JSON = ROOT_DIR / "central-utils" / "data" / "parcelamentos" / "parcelamentos.import.json"
DEFAULT_CONFIG = ROOT_DIR / "python" / "Parcelamentos Chamador" / "consultores.parcelamentos.json"
DEFAULT_OUTPUT_DIR = ROOT_DIR / "python" / "Parcelamentos Chamador" / "saida"


@dataclass
class ChecklistItem:
    index: int
    company_name: str
    cnpj: str
    parcelamento_type: str
    parcelamento_number: str
    parcelamento_number_digits: str
    consultores: List[str]
    status: str
    raw: Dict[str, Any]


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().upper()
    text = (
        text.replace("Á", "A").replace("À", "A").replace("Ã", "A").replace("Â", "A")
        .replace("É", "E").replace("Ê", "E")
        .replace("Í", "I")
        .replace("Ó", "O").replace("Ô", "O").replace("Õ", "O")
        .replace("Ú", "U").replace("Ü", "U")
        .replace("Ç", "C")
    )
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_cnpj(value: Any) -> str:
    return re.sub(r"\D+", "", str(value or ""))


def normalize_digits(value: Any) -> str:
    return re.sub(r"\D+", "", str(value or ""))


def ensure_path(path_str: str, base: Optional[Path] = None) -> Path:
    p = Path(path_str)
    if p.is_absolute():
        return p
    if base:
        return (base / p).resolve()
    return p.resolve()


def load_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as f:
        return json.load(f)


def load_config(path: Path) -> Dict[str, Any]:
    config = load_json(path)
    config.setdefault("consultores", {})
    config.setdefault("typeMap", {})
    return config


def build_checklist(items: Iterable[Dict[str, Any]], type_map: Dict[str, List[str]]) -> List[ChecklistItem]:
    checklist: List[ChecklistItem] = []

    for idx, raw in enumerate(items, start=1):
        company_name = str(raw.get("companyName") or "").strip()
        cnpj = normalize_cnpj(raw.get("cnpj"))
        parcelamento_type = normalize_text(raw.get("parcelamentoType"))
        parcelamento_number = str(raw.get("parcelamentoNumber") or "").strip() or "S/N"
        parcelamento_number_digits = normalize_digits(parcelamento_number)

        consultores = list(type_map.get(parcelamento_type, []))
        if consultores:
            cnpj_ok = len(cnpj) in (11, 14)
            numero_ok = bool(parcelamento_number_digits)
            if cnpj_ok and numero_ok:
                status = "pendente"
            else:
                consultores = []
                status = "dados_insuficientes"
        else:
            status = "sem_consultor"

        checklist.append(
            ChecklistItem(
                index=idx,
                company_name=company_name,
                cnpj=cnpj,
                parcelamento_type=parcelamento_type,
                parcelamento_number=parcelamento_number,
                parcelamento_number_digits=parcelamento_number_digits,
                consultores=consultores,
                status=status,
                raw=raw,
            )
        )

    return checklist


def write_checklist_files(checklist: List[ChecklistItem], output_dir: Path) -> Dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    json_path = output_dir / f"checklist_parcelamentos_{stamp}.json"
    txt_path = output_dir / f"checklist_parcelamentos_{stamp}.txt"

    payload = {
        "generatedAt": datetime.now().isoformat(),
        "total": len(checklist),
        "items": [
            {
                "index": c.index,
                "companyName": c.company_name,
                "cnpj": c.cnpj,
                "parcelamentoType": c.parcelamento_type,
                "parcelamentoNumber": c.parcelamento_number,
                "parcelamentoNumberDigits": c.parcelamento_number_digits,
                "consultores": c.consultores,
                "status": c.status,
            }
            for c in checklist
        ],
    }

    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    grouped: Dict[str, List[ChecklistItem]] = defaultdict(list)
    for item in checklist:
        grouped[item.parcelamento_type].append(item)

    lines: List[str] = []
    lines.append(f"Checklist gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Total de registros: {len(checklist)}")
    lines.append("")

    for parcel_type in sorted(grouped.keys()):
        group_items = grouped[parcel_type]
        lines.append(f"[{parcel_type}] ({len(group_items)} registro(s))")
        for item in group_items:
            consultores = ", ".join(item.consultores) if item.consultores else "SEM CONSULTOR"
            lines.append(
                f"[ ] #{item.index:03d} | {item.company_name} | {item.cnpj} | "
                f"nro {item.parcelamento_number} | {consultores}"
            )
        lines.append("")

    txt_path.write_text("\n".join(lines), encoding="utf-8")

    return {"json": json_path, "txt": txt_path}


def get_current_month_checklist_files(output_dir: Path) -> Optional[Dict[str, Path]]:
    """
    Retorna o checklist mais recente do mes atual, se existir.
    """
    if not output_dir.exists():
        return None

    month_prefix = datetime.now().strftime("%Y%m")
    candidates = sorted(output_dir.glob(f"checklist_parcelamentos_{month_prefix}*.json"), reverse=True)
    if not candidates:
        return None

    json_path = candidates[0]
    stamp = json_path.stem.replace("checklist_parcelamentos_", "")
    txt_path = output_dir / f"checklist_parcelamentos_{stamp}.txt"
    if not txt_path.exists():
        return None

    return {"json": json_path, "txt": txt_path}


def build_placeholder_map(item: Optional[ChecklistItem], extra: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    data: Dict[str, str] = {}
    if item is not None:
        data = {
            "index": str(item.index),
            "company_name": item.company_name,
            "cnpj": item.cnpj,
            "parcelamento_type": item.parcelamento_type,
            "parcelamento_number": item.parcelamento_number,
            "parcelamento_number_digits": item.parcelamento_number_digits,
        }
    if extra:
        for key, value in extra.items():
            data[key] = str(value)
    return data


def format_command(template: List[str], placeholders: Dict[str, str]) -> List[str]:
    command: List[str] = []
    for token in template:
        value = str(token)
        for key, rep in placeholders.items():
            value = value.replace("{" + key + "}", rep)
        command.append(value)
    return command


def run_tasks(
    checklist: List[ChecklistItem],
    config: Dict[str, Any],
    output_paths: Dict[str, Path],
    output_dir: Path,
    continuar_em_erro: bool,
) -> int:
    consultores: Dict[str, Dict[str, Any]] = config.get("consultores", {})

    by_consultor: Dict[str, List[ChecklistItem]] = defaultdict(list)
    for item in checklist:
        for consultor_id in item.consultores:
            by_consultor[consultor_id].append(item)

    execution_log: List[Dict[str, Any]] = []
    exit_code = 0

    for consultor_id, items in by_consultor.items():
        spec = consultores.get(consultor_id)
        if not spec:
            execution_log.append(
                {
                    "consultorId": consultor_id,
                    "status": "ignorado",
                    "reason": "consultor nao encontrado na configuracao",
                    "items": len(items),
                }
            )
            continue

        if not spec.get("enabled", False):
            execution_log.append(
                {
                    "consultorId": consultor_id,
                    "name": spec.get("name", consultor_id),
                    "status": "ignorado",
                    "reason": "consultor desabilitado",
                    "items": len(items),
                }
            )
            continue

        mode = str(spec.get("mode") or "batch").strip().lower()
        command_template = spec.get("command") or []
        if not isinstance(command_template, list) or not command_template:
            execution_log.append(
                {
                    "consultorId": consultor_id,
                    "name": spec.get("name", consultor_id),
                    "status": "erro",
                    "reason": "command vazio",
                    "items": len(items),
                }
            )
            exit_code = 1
            if not continuar_em_erro:
                break
            continue

        workdir = ensure_path(str(spec.get("workdir") or ROOT_DIR), ROOT_DIR)

        if mode == "batch":
            batch_path = output_dir / f"lote_{consultor_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with batch_path.open("w", encoding="utf-8") as f:
                json.dump(
                    {
                        "consultorId": consultor_id,
                        "name": spec.get("name", consultor_id),
                        "total": len(items),
                        "items": [
                            {
                                "index": i.index,
                                "companyName": i.company_name,
                                "cnpj": i.cnpj,
                                "parcelamentoType": i.parcelamento_type,
                                "parcelamentoNumber": i.parcelamento_number,
                            }
                            for i in items
                        ],
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )

            placeholders = build_placeholder_map(
                None,
                {
                    "lote_json": str(batch_path),
                    "checklist_json": str(output_paths["json"]),
                    "checklist_txt": str(output_paths["txt"]),
                    "total_items": str(len(items)),
                },
            )

            command = format_command(command_template, placeholders)
            started = time.time()
            process = subprocess.run(command, cwd=str(workdir), check=False)
            elapsed = round(time.time() - started, 2)

            entry = {
                "consultorId": consultor_id,
                "name": spec.get("name", consultor_id),
                "mode": mode,
                "command": command,
                "workdir": str(workdir),
                "returnCode": process.returncode,
                "seconds": elapsed,
                "items": len(items),
            }
            execution_log.append(entry)

            if process.returncode != 0:
                exit_code = process.returncode or 1
                if not continuar_em_erro:
                    break

        elif mode == "per_item":
            for item in items:
                placeholders = build_placeholder_map(
                    item,
                    {
                        "checklist_json": str(output_paths["json"]),
                        "checklist_txt": str(output_paths["txt"]),
                    },
                )
                command = format_command(command_template, placeholders)

                started = time.time()
                process = subprocess.run(command, cwd=str(workdir), check=False)
                elapsed = round(time.time() - started, 2)

                entry = {
                    "consultorId": consultor_id,
                    "name": spec.get("name", consultor_id),
                    "mode": mode,
                    "index": item.index,
                    "companyName": item.company_name,
                    "cnpj": item.cnpj,
                    "parcelamentoType": item.parcelamento_type,
                    "parcelamentoNumber": item.parcelamento_number,
                    "command": command,
                    "workdir": str(workdir),
                    "returnCode": process.returncode,
                    "seconds": elapsed,
                }
                execution_log.append(entry)

                if process.returncode != 0:
                    exit_code = process.returncode or 1
                    if not continuar_em_erro:
                        break

            if exit_code != 0 and not continuar_em_erro:
                break

        else:
            execution_log.append(
                {
                    "consultorId": consultor_id,
                    "name": spec.get("name", consultor_id),
                    "status": "erro",
                    "reason": f"modo invalido: {mode}",
                    "items": len(items),
                }
            )
            exit_code = 1
            if not continuar_em_erro:
                break

    exec_log_path = output_dir / f"execucao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with exec_log_path.open("w", encoding="utf-8") as f:
        json.dump({"executions": execution_log}, f, ensure_ascii=False, indent=2)

    print(f"\nRelatorio de execucao: {exec_log_path}")
    return exit_code


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Le parcelamentos.import.json, monta checklist e chama consultores por tipo.",
    )
    parser.add_argument(
        "--arquivo-json",
        default=str(DEFAULT_IMPORT_JSON),
        help="Caminho do parcelamentos.import.json",
    )
    parser.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG),
        help="Arquivo de configuracao dos consultores (JSON)",
    )
    parser.add_argument(
        "--saida-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Pasta para checklist e logs",
    )
    parser.add_argument(
        "--tipos",
        default="",
        help="Filtra tipos de parcelamento (separados por virgula), ex: PGFN,ICMS",
    )
    parser.add_argument(
        "--executar",
        action="store_true",
        help="Executa os consultores configurados (sem esse flag, gera apenas checklist)",
    )
    parser.add_argument(
        "--continuar-em-erro",
        action="store_true",
        help="Continua a execucao mesmo se um consultor retornar erro",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    arquivo_json = ensure_path(args.arquivo_json)
    config_path = ensure_path(args.config)
    output_dir = ensure_path(args.saida_dir)

    if not arquivo_json.exists():
        print(f"ERRO: arquivo nao encontrado: {arquivo_json}")
        return 1

    if not config_path.exists():
        print(f"ERRO: config nao encontrado: {config_path}")
        return 1

    data = load_json(arquivo_json)
    config = load_config(config_path)

    items = data.get("items") or []
    if not isinstance(items, list):
        print("ERRO: chave 'items' invalida no JSON")
        return 1

    checklist = build_checklist(items, config.get("typeMap", {}))

    if args.tipos.strip():
        filtros = {normalize_text(x) for x in args.tipos.split(",") if x.strip()}
        checklist = [c for c in checklist if c.parcelamento_type in filtros]

    output_paths = get_current_month_checklist_files(output_dir)
    checklist_reutilizado = output_paths is not None
    if output_paths is None:
        output_paths = write_checklist_files(checklist, output_dir)

    by_type = Counter(c.parcelamento_type for c in checklist)
    by_status = Counter(c.status for c in checklist)
    consultor_counter = Counter()
    for item in checklist:
        for consultor in item.consultores:
            consultor_counter[consultor] += 1

    print("\nResumo")
    print(f"- Arquivo lido: {arquivo_json}")
    print(f"- Total de registros no checklist: {len(checklist)}")
    print(f"- Com consultor: {by_status.get('pendente', 0)}")
    print(f"- Com tipo mapeado mas sem dados: {by_status.get('dados_insuficientes', 0)}")
    print(f"- Sem consultor: {by_status.get('sem_consultor', 0)}")
    print(f"- Checklist do mes atual: {'reutilizado' if checklist_reutilizado else 'gerado agora'}")
    print(f"- Checklist JSON: {output_paths['json']}")
    print(f"- Checklist TXT: {output_paths['txt']}")

    print("\nTipos encontrados")
    for parcel_type, total in sorted(by_type.items(), key=lambda x: (-x[1], x[0])):
        print(f"- {parcel_type}: {total}")

    print("\nCarga por consultor")
    if consultor_counter:
        for consultor_id, total in sorted(consultor_counter.items(), key=lambda x: (-x[1], x[0])):
            print(f"- {consultor_id}: {total}")
    else:
        print("- Nenhum consultor mapeado")

    if not args.executar:
        print("\nModo checklist: nenhum consultor foi executado.")
        return 0

    print("\nModo execucao habilitado: iniciando consultores...")
    return run_tasks(
        checklist=checklist,
        config=config,
        output_paths=output_paths,
        output_dir=output_dir,
        continuar_em_erro=bool(args.continuar_em_erro),
    )


if __name__ == "__main__":
    sys.exit(main())
