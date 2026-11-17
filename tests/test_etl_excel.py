"""Pruebas unitarias e idempotencia del pipeline ETL Excel -> SQLite."""
from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from tools.etl_excel_isaac import DEFAULT_EXCEL_PATH, NOMBRES_SENSIBLES, ejecutar_etl


@pytest.mark.skipif(
    not DEFAULT_EXCEL_PATH.exists(),
    reason="Archivo Excel operativo no disponible en el entorno.",
)
def test_etl_crea_esquema_carga_datos_y_es_idempotente(tmp_path: Path) -> None:
    db_path = tmp_path / "test_etl.db"

    res1 = ejecutar_etl(DEFAULT_EXCEL_PATH, db_path, solo_2026=True)
    assert res1["total_registros"] > 0

    conn = sqlite3.connect(db_path)
    try:
        roles = {row[0] for row in conn.execute("SELECT nombre FROM roles").fetchall()}
        assert "operario" in roles

        tablas = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        assert "tarifas_aditamentos" in tablas
        assert "bitacora_horometros" in tablas

        total_maq_1 = conn.execute("SELECT COUNT(*) FROM maquinarias").fetchone()[0]
        total_tarifas_1 = conn.execute("SELECT COUNT(*) FROM tarifas_aditamentos").fetchone()[0]
        total_bitacora_1 = conn.execute("SELECT COUNT(*) FROM bitacora_horometros").fetchone()[0]

        assert total_maq_1 >= 13
        assert total_tarifas_1 > 0
        assert total_bitacora_1 == res1["total_registros"]

        observaciones = [
            row[0]
            for row in conn.execute(
                "SELECT observacion FROM bitacora_horometros WHERE observacion IS NOT NULL"
            ).fetchall()
        ]
        for obs in observaciones:
            obs_upper = obs.upper()
            for nombre_sensible in NOMBRES_SENSIBLES:
                assert nombre_sensible not in obs_upper, (
                    f"Dato sensible detectado en observación: {nombre_sensible}"
                )
    finally:
        conn.close()

    # Segunda ejecución sobre la misma BD: debe ser idempotente (sin duplicar filas)
    res2 = ejecutar_etl(DEFAULT_EXCEL_PATH, db_path, solo_2026=True)
    assert res2["total_registros"] == res1["total_registros"]

    conn2 = sqlite3.connect(db_path)
    try:
        assert conn2.execute("SELECT COUNT(*) FROM maquinarias").fetchone()[0] == total_maq_1
        assert conn2.execute("SELECT COUNT(*) FROM tarifas_aditamentos").fetchone()[0] == total_tarifas_1
        assert conn2.execute("SELECT COUNT(*) FROM bitacora_horometros").fetchone()[0] == total_bitacora_1
    finally:
        conn2.close()
