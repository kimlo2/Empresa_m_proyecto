"""
Pipeline ETL para extraer, limpiar, anonimizar y migrar los registros operativos
de 2026 desde `REALACION EQUIPOS  ISAAC.xlsx` hacia SQLite (`app/msps.db`).

Cumplimiento Ley 1581 de 2012 (Habeas Data):
- Trunca y excluye todas las filas inferiores de nómina, salarios y deducciones.
- Filtra nombres propios de operarios y terceros en las observaciones.
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import openpyxl

from app.database import DB_PATH, SCHEMA_PATH

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_EXCEL_PATH = ROOT_DIR.parent / "REALACION EQUIPOS  ISAAC.xlsx"

# Mapeo de las 13 pestañas del Excel -> catálogo en SQLite (enlazado por placa)
CATALOGO_EXCEL: dict[str, dict[str, Any]] = {
    "CATERPILLAR  312D": {
        "marca": "Caterpillar",
        "modelo": "312D",
        "categoria": "Excavadora",
        "placa": "PENDIENTE-CAT312D",
        "capacidad": "0,5 – 0,7 m³",
        "descripcion": "Excavadora hidráulica de orugas Caterpillar 312D (12 toneladas).",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "HITACHI  75": {
        "marca": "Hitachi",
        "modelo": "ZAXIS 75",
        "categoria": "Excavadora",
        "placa": "MC758262",
        "capacidad": None,
        "descripcion": "Excavadora hidráulica de orugas de 8 toneladas con radio de giro reducido.",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "HITACHI 135": {
        "marca": "Hitachi",
        "modelo": "ZAXIS 135",
        "categoria": "Excavadora",
        "placa": "PENDIENTE-HIT135",
        "capacidad": "0,5 m³",
        "descripcion": "Excavadora hidráulica de orugas Hitachi ZAXIS 135 (13.5 toneladas).",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "JD 200G": {
        "marca": "John Deere",
        "modelo": "200G",
        "categoria": "Excavadora",
        "placa": "MC770043",
        "capacidad": "0,42 – 1,02 m³",
        "descripcion": "Excavadora hidráulica de orugas de 21 toneladas.",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "JD 310SL  2018": {
        "marca": "John Deere",
        "modelo": "310SL 2018",
        "categoria": "Retroexcavadora",
        "placa": "PENDIENTE-JD310-18",
        "capacidad": "0,2 – 1,0 m³",
        "descripcion": "Retroexcavadora cargadora John Deere 310SL modelo 2018.",
        "tarifa_balde": 140000.0,
        "tarifa_martillo": 200000.0,
        "tarifa_transporte": 900000.0,
    },
    "JD  310SL  2023": {
        "marca": "John Deere",
        "modelo": "310SL 2023",
        "categoria": "Retroexcavadora",
        "placa": "PENDIENTE-JD310-23",
        "capacidad": "0,2 – 1,0 m³",
        "descripcion": "Retroexcavadora cargadora John Deere 310SL modelo 2023.",
        "tarifa_balde": 140000.0,
        "tarifa_martillo": 200000.0,
        "tarifa_transporte": 900000.0,
    },
    "JD  310SL 2019": {
        "marca": "John Deere",
        "modelo": "310SL 2019",
        "categoria": "Retroexcavadora",
        "placa": "PENDIENTE-JD310-19",
        "capacidad": "0,2 – 1,0 m³",
        "descripcion": "Retroexcavadora cargadora John Deere 310SL modelo 2019.",
        "tarifa_balde": 140000.0,
        "tarifa_martillo": 200000.0,
        "tarifa_transporte": 900000.0,
    },
    "KOBELCO 28": {
        "marca": "Kobelco",
        "modelo": "SK28",
        "categoria": "Excavadora",
        "placa": "PENDIENTE-KOB28",
        "capacidad": "0,08 m³",
        "descripcion": "Miniexcavadora hidráulica Kobelco de 2.8 toneladas para espacios reducidos.",
        "tarifa_balde": 130000.0,
        "tarifa_martillo": 180000.0,
        "tarifa_transporte": 700000.0,
    },
    "KOBELCO 140": {
        "marca": "Kobelco",
        "modelo": "SK140LC",
        "categoria": "Excavadora",
        "placa": "MC662589",
        "capacidad": "0,5 – 0,7 m³",
        "descripcion": "Excavadora hidráulica de orugas de 14 toneladas.",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "LINK BELT 210X3": {
        "marca": "Link-Belt",
        "modelo": "210 X3",
        "categoria": "Excavadora",
        "placa": "MC109248",
        "capacidad": "hasta 1,0 m³",
        "descripcion": "Excavadora hidráulica de orugas de 22 toneladas con motor Isuzu.",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "KOMATSU  130": {
        "marca": "Komatsu",
        "modelo": "PC130",
        "categoria": "Excavadora",
        "placa": "PENDIENTE-PC130",
        "capacidad": "0,18 – 0,60 m³",
        "descripcion": "Excavadora hidráulica de orugas de 13 toneladas.",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
    "KOMATSU PC 60": {
        "marca": "Komatsu",
        "modelo": "PC60",
        "categoria": "Excavadora",
        "placa": "MC112258",
        "capacidad": "0,25 – 0,37 m³",
        "descripcion": "Miniexcavadora hidráulica de orugas de 6 toneladas.",
        "tarifa_balde": 135000.0,
        "tarifa_martillo": 185000.0,
        "tarifa_transporte": 800000.0,
    },
    "LINK BELT  130": {
        "marca": "Link-Belt",
        "modelo": "130 X3",
        "categoria": "Excavadora",
        "placa": "PENDIENTE-LB130",
        "capacidad": "0,5 m³",
        "descripcion": "Excavadora hidráulica de orugas Link-Belt de 13 toneladas.",
        "tarifa_balde": 160000.0,
        "tarifa_martillo": 220000.0,
        "tarifa_transporte": 1200000.0,
    },
}

# Términos propios/sensibles que nunca deben persistirse en `observacion` (Ley 1581)
NOMBRES_SENSIBLES = (
    "GILDARDO",
    "PEDRO",
    "EDUARDO",
    "MAYORGA",
    "FABIO",
    "WILLIAM",
    "ANDERSON",
    "ALFONSO",
    "JUAN PABLO",
    "PEPE",
    "MORA",
    "JOSE BARRERA",
    "TULIO",
    "CALDERON",
)


@dataclass
class BloqueMensual:
    header_row: int
    fecha_col: int
    end_col: int
    col_map: dict[str, int]


def _to_float(val: Any) -> Optional[float]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        f = float(val)
        return f if f >= 0 else None
    if isinstance(val, str):
        txt = val.strip().replace(",", ".")
        m = re.search(r"-?\d+(?:\.\d+)?", txt)
        if m:
            f = float(m.group(0))
            return f if f >= 0 else None
    return None


def _parse_fecha(val: Any) -> Optional[dt.date]:
    if isinstance(val, dt.datetime):
        d = val.date()
    elif isinstance(val, dt.date):
        d = val
    elif isinstance(val, (int, float)) and 43831 <= val <= 47000:
        d = (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(val))).date()
    else:
        return None

    # Filtra fechas atípicas de digitación (ej. 2021-10 aislada)
    if d.year not in (2023, 2026):
        return None
    return d


def _detectar_bloques(ws: openpyxl.worksheet.worksheet.Worksheet) -> list[BloqueMensual]:
    """Detecta todos los bloques horizontales mensuales buscando la cabecera `FECHA`."""
    posiciones: list[tuple[int, int]] = []
    for r in range(1, 8):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if isinstance(val, str) and val.strip().upper() == "FECHA":
                posiciones.append((r, c))

    # Ordenar por columna para delimitar el ancho de cada bloque hasta el siguiente FECHA
    posiciones.sort(key=lambda item: item[1])
    bloques: list[BloqueMensual] = []
    for idx, (r, c) in enumerate(posiciones):
        next_fecha_col = posiciones[idx + 1][1] if idx + 1 < len(posiciones) else ws.max_column + 1
        end_col = min(next_fecha_col - 1, c + 8)
        col_map: dict[str, int] = {"fecha": c}

        for cc in range(c + 1, end_col + 1):
            hdr = ws.cell(r, cc).value
            if not isinstance(hdr, str):
                continue
            h = hdr.strip().upper()
            if "HOROMETRO INICIAL" in h or "H- INICIAL" in h:
                col_map["horo_ini"] = cc
            elif "HOROMETRO FINAL" in h or "HOROMETO FINAL" in h or "H-FINAL" in h:
                col_map["horo_fin"] = cc
            elif h in ("HORAS", "# HORAS"):
                col_map["horas"] = cc
            elif "EXTRA" in h or "EXTAR" in h:
                col_map["extras"] = cc
            elif "BALDE" in h:
                col_map["balde"] = cc
            elif "MARTILLO" in h or "MARTLLO" in h:
                col_map["martillo"] = cc
            elif "STAMBY" in h or "STANDBY" in h:
                col_map["standby"] = cc
            elif "COMBUSTIBLE" in h or "COMBISTIBLE" in h:
                col_map["combustible"] = cc

        bloques.append(BloqueMensual(header_row=r, fecha_col=c, end_col=end_col, col_map=col_map))
    return bloques


def _es_texto_sensible(texto: str) -> bool:
    upper = texto.upper()
    return any(nombre in upper for nombre in NOMBRES_SENSIBLES)


def _diagnosticar_fila(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row_idx: int,
    bloque: BloqueMensual,
    horas: float,
    balde: float,
    martillo: float,
) -> tuple[str, Optional[str]]:
    """Diagnostica el estado operativo del día y limpia notas sensibles."""
    # Buscar palabras clave de novedad en todo el ancho del bloque (+1 columna adyacente)
    notas: list[str] = []
    scan_end = min(bloque.end_col + 1, ws.max_column)
    for cc in range(max(1, bloque.fecha_col - 1), scan_end + 1):
        val = ws.cell(row_idx, cc).value
        if isinstance(val, str):
            txt = val.strip()
            if txt and not _es_texto_sensible(txt):
                notas.append(txt)

    texto_unido = " | ".join(notas).upper()
    if "VARADO" in texto_unido:
        return "varada", "Novedad reportada en planilla: VARADOS"
    if "REPA" in texto_unido:
        return "reparacion", "Novedad reportada en planilla: REPA"
    if "STAMB" in texto_unido or "STANDB" in texto_unido:
        return "standby", "Equipo en espera operativa (STAMBY)"

    standby_col = bloque.col_map.get("standby")
    if standby_col:
        st_val = _to_float(ws.cell(row_idx, standby_col).value)
        if st_val and st_val > 0 and horas == 0:
            return "standby", f"Standby registrado ({st_val} h)"

    if horas > 0 or balde > 0 or martillo > 0:
        return "operativa", None
    return "inactiva", None


def _asegurar_maquina(conn: sqlite3.Connection, spec: dict[str, Any]) -> int:
    # 1. Asegurar categoría
    cat_row = conn.execute(
        "SELECT id_categoria FROM categorias WHERE nombre_categoria = ?",
        (spec["categoria"],),
    ).fetchone()
    if cat_row:
        id_cat = cat_row["id_categoria"]
    else:
        cur = conn.execute(
            "INSERT INTO categorias (nombre_categoria) VALUES (?)",
            (spec["categoria"],),
        )
        id_cat = cur.lastrowid

    # 2. Buscar máquina por placa
    maq_row = conn.execute(
        "SELECT id_maquinaria, costo_por_hora FROM maquinarias WHERE placa = ?",
        (spec["placa"],),
    ).fetchone()
    if maq_row:
        id_maq = maq_row["id_maquinaria"]
        if (maq_row["costo_por_hora"] or 0) == 0:
            conn.execute(
                "UPDATE maquinarias SET costo_por_hora = ? WHERE id_maquinaria = ?",
                (spec["tarifa_balde"], id_maq),
            )
    else:
        cur = conn.execute(
            """INSERT INTO maquinarias
               (marca, descripcion, capacidad, modelo, costo_por_hora, placa, disponible, id_categoria)
               VALUES (?, ?, ?, ?, ?, ?, 1, ?)""",
            (
                spec["marca"],
                spec["descripcion"],
                spec["capacidad"],
                spec["modelo"],
                spec["tarifa_balde"],
                spec["placa"],
                id_cat,
            ),
        )
        id_maq = cur.lastrowid

    # 3. Registrar tarifas diferenciadas por aditamento y transporte
    tarifas = [
        ("balde", spec["tarifa_balde"], "hora"),
        ("martillo", spec["tarifa_martillo"], "hora"),
        ("transporte", spec["tarifa_transporte"], "viaje"),
    ]
    for tipo, valor, unidad in tarifas:
        conn.execute(
            """INSERT INTO tarifas_aditamentos (id_maquinaria, tipo_servicio, tarifa_cop, unidad)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(id_maquinaria, tipo_servicio)
               DO UPDATE SET tarifa_cop = excluded.tarifa_cop, unidad = excluded.unidad""",
            (id_maq, tipo, valor, unidad),
        )
    return id_maq


def ejecutar_etl(excel_path: Path, db_path: Path, solo_2026: bool = True) -> dict[str, Any]:
    if not excel_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo Excel en: {excel_path}")

    db_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))

    resumen: list[dict[str, Any]] = []
    total_insertados = 0

    try:
        with conn:
            for sheet_name in wb.sheetnames:
                spec = CATALOGO_EXCEL.get(sheet_name)
                if not spec:
                    continue
                id_maq = _asegurar_maquina(conn, spec)
                ws = wb[sheet_name]
                bloques = _detectar_bloques(ws)

                registros_por_fecha: dict[str, tuple] = {}
                for b in bloques:
                    for r in range(b.header_row + 1, min(b.header_row + 33, ws.max_row + 1)):
                        fecha = _parse_fecha(ws.cell(r, b.fecha_col).value)
                        if fecha is None:
                            continue
                        if solo_2026 and fecha.year != 2026:
                            continue

                        h_ini = _to_float(ws.cell(r, b.col_map["horo_ini"]).value) if "horo_ini" in b.col_map else None
                        h_fin = _to_float(ws.cell(r, b.col_map["horo_fin"]).value) if "horo_fin" in b.col_map else None
                        horas = (_to_float(ws.cell(r, b.col_map["horas"]).value) or 0.0) if "horas" in b.col_map else 0.0
                        extras = (_to_float(ws.cell(r, b.col_map["extras"]).value) or 0.0) if "extras" in b.col_map else 0.0
                        balde = (_to_float(ws.cell(r, b.col_map["balde"]).value) or 0.0) if "balde" in b.col_map else 0.0
                        martillo = (_to_float(ws.cell(r, b.col_map["martillo"]).value) or 0.0) if "martillo" in b.col_map else 0.0
                        comb = (_to_float(ws.cell(r, b.col_map["combustible"]).value) or 0.0) if "combustible" in b.col_map else 0.0

                        # Si no hay columna HORAS pero sí horómetros válidos, calcular diferencia
                        if horas == 0.0 and h_ini is not None and h_fin is not None and h_fin > h_ini:
                            horas = round(h_fin - h_ini, 2)

                        estado, obs = _diagnosticar_fila(ws, r, b, horas, balde, martillo)
                        fecha_iso = fecha.isoformat()

                        registros_por_fecha[fecha_iso] = (
                            id_maq,
                            fecha_iso,
                            h_ini,
                            h_fin,
                            round(horas, 2),
                            round(extras, 2),
                            round(balde, 2),
                            round(martillo, 2),
                            round(comb, 2),
                            estado,
                            obs,
                            "excel_etl",
                        )

                for reg in registros_por_fecha.values():
                    conn.execute(
                        """INSERT INTO bitacora_horometros
                           (id_maquinaria, fecha, horometro_inicial, horometro_final,
                            horas_trabajadas, horas_extras, horas_balde, horas_martillo,
                            combustible_gal, estado_dia, observacion, origen)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                           ON CONFLICT(id_maquinaria, fecha, origen)
                           DO UPDATE SET
                             horometro_inicial = excluded.horometro_inicial,
                             horometro_final = excluded.horometro_final,
                             horas_trabajadas = excluded.horas_trabajadas,
                             horas_extras = excluded.horas_extras,
                             horas_balde = excluded.horas_balde,
                             horas_martillo = excluded.horas_martillo,
                             combustible_gal = excluded.combustible_gal,
                             estado_dia = excluded.estado_dia,
                             observacion = excluded.observacion""",
                        reg,
                    )

                horas_totales = sum(r[4] for r in registros_por_fecha.values())
                dias_op = sum(1 for r in registros_por_fecha.values() if r[9] == "operativa")
                dias_inac = sum(1 for r in registros_por_fecha.values() if r[9] == "inactiva")
                novedades = sum(1 for r in registros_por_fecha.values() if r[9] in ("varada", "reparacion", "standby"))
                total_insertados += len(registros_por_fecha)

                resumen.append(
                    {
                        "hoja": sheet_name,
                        "id_maquinaria": id_maq,
                        "placa": spec["placa"],
                        "registros": len(registros_por_fecha),
                        "dias_operativos": dias_op,
                        "dias_inactivos": dias_inac,
                        "novedades": novedades,
                        "horas_totales": round(horas_totales, 1),
                    }
                )
    finally:
        conn.close()

    return {"total_registros": total_insertados, "maquinas": resumen}


def main() -> None:
    parser = argparse.ArgumentParser(description="ETL Excel de Maquiservicio Pérez -> SQLite")
    parser.add_argument("--excel", type=Path, default=DEFAULT_EXCEL_PATH, help="Ruta al archivo Excel")
    parser.add_argument("--db", type=Path, default=DB_PATH, help="Ruta a msps.db")
    parser.add_argument("--incluir-2023", action="store_true", help="Incluye también los registros históricos de 2023")
    args = parser.parse_args()

    resultado = ejecutar_etl(args.excel, args.db, solo_2026=not args.incluir_2023)
    print(f"\n=== ETL COMPLETADO: {resultado['total_registros']} registros cargados en {args.db.name} ===")
    print(f"{'Hoja Excel':<20} {'Placa':<19} {'Reg':>4} {'Oper':>5} {'Inac':>5} {'Nov':>4} {'Horas':>8}")
    print("-" * 72)
    for m in resultado["maquinas"]:
        print(
            f"{m['hoja']:<20} {m['placa']:<19} {m['registros']:>4} "
            f"{m['dias_operativos']:>5} {m['dias_inactivos']:>5} {m['novedades']:>4} {m['horas_totales']:>8.1f}"
        )


if __name__ == "__main__":
    main()
