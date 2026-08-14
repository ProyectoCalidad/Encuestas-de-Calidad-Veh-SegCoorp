"""
build_data.py — Convierte la BBDD mensual de Segurcoop (Excel) al data.json
que consume el panel (index.html).

Uso:
    python build_data.py BBDD_Segurcoop.xlsx

Genera (o sobrescribe) data.json en la misma carpeta que este script.
"""
import sys
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

SHEET_NAME = "Acumulado"

# Columnas de puntaje 1-5: a veces vienen con texto suelto (ej. "Muy satisfecho",
# "No", "Gracias") en vez del número — se descartan y quedan como nulos.
SCORE_COLS = ["ATENCION", "DEMORA", "PRESTADOR"]


def to_score(value):
    if pd.isna(value):
        return None
    s = str(value).strip()
    if re.fullmatch(r"[1-5]", s):
        return int(s)
    return None


def to_month_key(ts):
    if pd.isna(ts):
        return None
    return pd.Timestamp(ts).strftime("%Y-%m")


def build(input_path: Path, output_path: Path):
    df = pd.read_excel(input_path, sheet_name=SHEET_NAME)
    df.columns = [c.strip() for c in df.columns]

    casos = []
    for _, row in df.iterrows():
        casos.append({
            "mes": to_month_key(row.get("Mes")),
            "provincia": (row.get("Provincia") or None) if pd.notna(row.get("Provincia")) else None,
            "vehiculo": (row.get("Tipo de Vehículo") or None) if pd.notna(row.get("Tipo de Vehículo")) else None,
            "servicio": (row.get("Servicio") or None) if pd.notna(row.get("Servicio")) else None,
            "problema": (row.get("Problema") or None) if pd.notna(row.get("Problema")) else None,
            "atencion": to_score(row.get("ATENCION")),
            "demora": to_score(row.get("DEMORA")),
            "prestador": to_score(row.get("PRESTADOR")),
            "recomienda": row.get("Recomienda Ok") if pd.notna(row.get("Recomienda Ok")) else None,
            "respondio": row.get("RESPONDIO?") if pd.notna(row.get("RESPONDIO?")) else None,
        })

    meses = sorted({c["mes"] for c in casos if c["mes"]})

    data = {
        "generadoEl": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "meses": meses,
        "casos": casos,
    }

    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"OK -> {output_path} ({len(casos)} casos, meses: {', '.join(meses)})")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python build_data.py <ruta al Excel de la BBDD>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(__file__).parent / "data.json"
    build(input_path, output_path)
