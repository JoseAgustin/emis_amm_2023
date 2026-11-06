
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_emisiones_carreteras.py
==================================
Script para actualizar emisiones VEHICULARES en el inventario AMM 2023.
Lee FM_emis_vehiculo_muni_AMv2.xlsx y actualiza los archivos M_*.csv.

Fuente: Emisiones vehiculares por tipo de vehículo y GRIDCODE.
Cada vehículo tiene un SCC de escape (330) y opcionalmente uno evaporativo (33V).

Estructura de M_*.csv (formato Fortran):
  R1: ' GRIDCODE, emissions in gram per year,,,,,,,,,,,,,,,,,'
  R2: '      N, SCC_1, SCC_2, ..., SCC_N'
  R3+: '  GRIDCODE,  X.XXXXE+XX, ...,  X.XXXXE+XX, 6'

Características:
- 9 tipos de vehículo con SCC directo (sin proporciones)
- 13 SCCs: 8 escape (330) + 5 evaporativos (33V)
- NOx se descompone en NO y NO2 con relación NO/NO2 = 9.46
- M_VOC: TGHNoEva → SCC escape, TGHEvap → SCC evaporativo
- Protección Saltillo integrada (LONG <= -100.67)
- Conversión Mg/año → g/año (×10⁶)

Repositorio: emis_amm_2023
Autor: José Agustín García Reynoso — ICAyCC, UNAM
"""

import pandas as pd
import numpy as np
import os
import sys
import argparse
import shutil
from datetime import datetime

from proteccion_saltillo import cargar_localiza, proteger_saltillo

FACTOR_CONVERSION = 1e6  # Mg/year -> g/year
RATIO_NO_NO2 = 9.46

FRAC_NO2 = 1.0 / (1.0 + RATIO_NO_NO2)
FRAC_NO  = RATIO_NO_NO2 / (1.0 + RATIO_NO_NO2)

CAMPO_FINAL_DEFAULT = 6

# ============================================================
# MAPEO VEHÍCULO → SCC
# ============================================================
VEHICULOS = {
    "Autobús":              {"scc_escape": "2230075330", "scc_evap": "220107033V"},
    "Automóvil Particular": {"scc_escape": "2201001330", "scc_evap": "220100133V"},
    "Camioneta pasajeros":  {"scc_escape": "2201020330", "scc_evap": None},
    "Motocicleta":          {"scc_escape": "2201080330", "scc_evap": "220108033V"},
    "Pickup":               {"scc_escape": "2201040330", "scc_evap": "220102033V"},
    "Taxi":                 {"scc_escape": "2230001330", "scc_evap": "220100133V"},
    "Tractocamión":         {"scc_escape": "2230074330", "scc_evap": "220107033V"},
    "Veh<3.8 Ton":          {"scc_escape": "2230074330", "scc_evap": "220104033V"},
    "Veh>3.8 Ton":          {"scc_escape": "2230071330", "scc_evap": "220107033V"},
}

# SCCs ordenados (mismo orden que M_*.csv originales)
SCCS_ESCAPE = [
    "2201001330", "2201020330", "2201040330", "2201070330", "2201080330",
    "2230001330", "2230060330", "2230071330", "2230073330", "2230074330",
    "2230075330",
]
SCCS_EVAPORATIVOS = [
    "220100133V", "220102033V", "220104033V", "220107033V", "220108033V",
]
SCCS_TODOS = SCCS_ESCAPE + SCCS_EVAPORATIVOS

# ============================================================
# CONTAMINANTES → columna Excel → archivo CSV
# ============================================================
# Contaminantes regulares (escape): col_excel → SCC escape del vehículo
CONTAMINANTES_ESCAPE = {
    "PM10": {"col_excel": "SUM_PM10_CL", "csv": "M_PM10.csv"},
    "PM25": {"col_excel": "SUM_PM25_CL", "csv": "M_PM25.csv"},
    "SO2":  {"col_excel": "SUM_SO2_CL",  "csv": "M_SO2.csv"},
    "CO":   {"col_excel": "SUM_CO_CL",   "csv": "M_CO.csv"},
    "NH3":  {"col_excel": "SUM_NH3_CL",  "csv": "M_NH3.csv"},
    "CN":   {"col_excel": "SUM_CN_CL",   "csv": "M_CN.csv"},
    "CH4":  {"col_excel": "SUM_CH4_CL",  "csv": "M_CH4.csv"},
    "CO2":  {"col_excel": "SUM_CO2_CL",  "csv": "M_CO2.csv"},
}

# NOx se descompone en NO y NO2
CONTAMINANTES_NOX = {
    "NO":  {"col_excel": "SUM_NOx_CL", "csv": "M_NO.csv",  "fraccion": FRAC_NO},
    "NO2": {"col_excel": "SUM_NOx_CL", "csv": "M_NO2.csv", "fraccion": FRAC_NO2},
}

# VOC: TGHNoEva → SCC escape, TGHEvap → SCC evaporativo
CONTAMINANTE_VOC = {
    "csv": "M_VOC.csv",
    "col_no_evap": "SUM_TGHNoEva_CL",  # → SCC escape (330)
    "col_evap": "SUM_TGHEvap_CL",       # → SCC evaporativo (33V)
}


# ============================================================
# FUNCIONES DE LECTURA/ESCRITURA M_*.csv
# ============================================================
def leer_csv_vehicular(ruta):
    """
    Lee un archivo M_*.csv con formato Fortran.
    Retorna: encabezado_r1, sccs_list, datos {grid: {scc: valor}}, campo_final
    """
    with open(ruta, 'r') as f:
        lineas = f.readlines()

    # R1: encabezado
    encabezado_r1 = lineas[0].rstrip('\n')

    # R2: N, SCC_1, SCC_2, ..., SCC_N
    partes_r2 = lineas[1].strip().rstrip(',').split(',')
    n_sccs = int(partes_r2[0].strip())
    sccs_list = [p.strip() for p in partes_r2[1:]]

    # R3+: GRIDCODE, emis_1, ..., emis_N, campo_final
    datos = {}
    campo_final = CAMPO_FINAL_DEFAULT
    for linea in lineas[2:]:
        linea_limpia = linea.strip().rstrip(',')
        if not linea_limpia:
            continue
        partes = linea_limpia.split(',')
        grid = int(partes[0].strip())
        valores = {}
        for i, scc in enumerate(sccs_list):
            val_str = partes[i + 1].strip()
            try:
                valores[scc] = float(val_str)
            except ValueError:
                valores[scc] = 0.0
        # Campo final (huso UTM)
        if len(partes) > len(sccs_list) + 1:
            try:
                campo_final = int(partes[-1].strip())
            except ValueError:
                pass
        datos[grid] = valores

    return encabezado_r1, sccs_list, datos, campo_final


def escribir_csv_vehicular(ruta, encabezado_r1, sccs_list, datos, campo_final):
    """
    Escribe un archivo M_*.csv con formato Fortran.
    """
    n_sccs = len(sccs_list)

    with open(ruta, 'w') as f:
        # R1: encabezado con comas trailing
        n_comas = n_sccs + 1
        f.write(f" GRIDCODE, emissions in gram per year{',' * n_comas}\n")

        # R2: N, SCC_1, ..., SCC_N
        sccs_str = ', '.join(f' {scc}' for scc in sccs_list)
        f.write(f"      {n_sccs},{sccs_str}\n")

        # R3+: datos
        for grid in sorted(datos.keys()):
            valores = datos[grid]
            vals_str = ', '.join(f' {valores.get(scc, 0.0):.4E}' for scc in sccs_list)
            f.write(f"  {grid},{vals_str}, {campo_final}\n")


# ============================================================
# FUNCIÓN PRINCIPAL DE ACTUALIZACIÓN
# ============================================================
def actualizar_contaminante_escape(df_excel, csv_path, output_path, col_excel,
                                    grid_lon, factor=1.0, dry_run=False):
    """
    Actualiza un M_*.csv con emisiones de escape (SCC 330).
    Cada fila del Excel tiene un vehículo con su SCC directo.
    factor: multiplicador adicional (para NO/NO2).
    """
    csv_name = os.path.basename(csv_path)

    if not os.path.exists(csv_path):
        print(f"  ⚠️  {csv_name} no encontrado — omitiendo")
        return 0, 0

    if col_excel not in df_excel.columns:
        print(f"  ⚠️  Columna '{col_excel}' no encontrada — omitiendo")
        return 0, 0

    # Leer CSV original
    encabezado_r1, sccs_list, datos, campo_final = leer_csv_vehicular(csv_path)

    celdas_actualizadas = 0
    celdas_protegidas = 0

    # Agrupar por GRIDCODE y vehículo
    for _, row in df_excel.iterrows():
        grid = int(row['GRIDCODE'])
        vehiculo = row['Vehículo']
        valor_mg = row[col_excel]

        if pd.isna(valor_mg):
            valor_mg = 0.0

        # Obtener SCC de escape para este vehículo
        if vehiculo not in VEHICULOS:
            continue
        scc = VEHICULOS[vehiculo]["scc_escape"]

        # Convertir Mg/año → g/año y aplicar factor
        valor_g = float(valor_mg) * FACTOR_CONVERSION * factor

        # Verificar que el SCC existe en el CSV
        if scc not in sccs_list:
            # Agregar SCC si no existe
            sccs_list.append(scc)
            for g in datos:
                datos[g][scc] = 0.0

        # Verificar que el grid existe en el CSV
        if grid not in datos:
            datos[grid] = {s: 0.0 for s in sccs_list}

        # Protección Saltillo (estricta): Saltillo nunca se sobreescribe
        valor_actual = datos[grid].get(scc, 0.0)
        valor_final, fue_protegida = proteger_saltillo(grid, valor_g, valor_actual, grid_lon)

        if fue_protegida:
            # Saltillo → no escribir, conservar valor original del CSV
            celdas_protegidas += 1
        else:
            # ACUMULAR: múltiples vehículos pueden compartir el mismo SCC
            # (ej: Tractocamión y Veh<3.8 Ton ambos usan 2230074330)
            # Necesitamos sumar, no reemplazar
            datos[grid][scc] = datos[grid].get(scc, 0.0) + valor_final - valor_actual
            celdas_actualizadas += 1

    # Reordenar SCCs
    sccs_list = [s for s in SCCS_TODOS if s in sccs_list]
    # Agregar SCCs que no estaban en el orden predefinido
    for s in set(datos.get(list(datos.keys())[0], {}).keys()) - set(sccs_list):
        sccs_list.append(s)

    if not dry_run:
        escribir_csv_vehicular(output_path, encabezado_r1, sccs_list, datos, campo_final)

    return celdas_actualizadas, celdas_protegidas


def actualizar_voc(df_excel, csv_path, output_path, grid_lon, dry_run=False):
    """
    Actualiza M_VOC.csv con lógica especial:
    - SUM_TGHNoEva_CL → SCC escape (330) del vehículo
    - SUM_TGHEvap_CL  → SCC evaporativo (33V) del vehículo
    - Si vehículo no tiene SCC 33V → TGHEvap va al SCC escape
    """
    csv_name = os.path.basename(csv_path)
    col_no_evap = CONTAMINANTE_VOC["col_no_evap"]
    col_evap = CONTAMINANTE_VOC["col_evap"]

    if not os.path.exists(csv_path):
        print(f"  ⚠️  {csv_name} no encontrado — omitiendo")
        return 0, 0

    # Leer CSV original
    encabezado_r1, sccs_list, datos, campo_final = leer_csv_vehicular(csv_path)

    celdas_actualizadas = 0
    celdas_protegidas = 0

    # Primero: resetear a cero las celdas de Monterrey para los SCCs que vamos a escribir
    # (para evitar doble conteo al acumular)
    sccs_a_escribir = set()
    for v_info in VEHICULOS.values():
        sccs_a_escribir.add(v_info["scc_escape"])
        if v_info["scc_evap"]:
            sccs_a_escribir.add(v_info["scc_evap"])

    for grid in datos:
        if grid in grid_lon and grid_lon[grid] > -100.67:  # Monterrey
            for scc in sccs_a_escribir:
                if scc in datos[grid]:
                    datos[grid][scc] = 0.0

    for _, row in df_excel.iterrows():
        grid = int(row['GRIDCODE'])
        vehiculo = row['Vehículo']

        if vehiculo not in VEHICULOS:
            continue

        val_no_evap = float(row[col_no_evap]) if not pd.isna(row[col_no_evap]) else 0.0
        val_evap = float(row[col_evap]) if not pd.isna(row[col_evap]) else 0.0

        scc_esc = VEHICULOS[vehiculo]["scc_escape"]
        scc_ev = VEHICULOS[vehiculo]["scc_evap"]

        # Asegurar SCCs existen
        for scc in [scc_esc, scc_ev]:
            if scc and scc not in sccs_list:
                sccs_list.append(scc)
                for g in datos:
                    datos[g][scc] = 0.0

        if grid not in datos:
            datos[grid] = {s: 0.0 for s in sccs_list}

        # TGHNoEva → SCC escape
        valor_no_evap_g = val_no_evap * FACTOR_CONVERSION
        valor_actual_esc = datos[grid].get(scc_esc, 0.0)
        valor_final_esc, es_salt_esc = proteger_saltillo(
            grid, valor_no_evap_g, valor_actual_esc, grid_lon
        )

        if es_salt_esc:
            # Saltillo → no escribir, conservar valor original del CSV
            celdas_protegidas += 1
        else:
            datos[grid][scc_esc] = datos[grid].get(scc_esc, 0.0) + valor_final_esc
            celdas_actualizadas += 1

        # TGHEvap → SCC evaporativo (o escape si no tiene 33V)
        valor_evap_g = val_evap * FACTOR_CONVERSION
        scc_destino_evap = scc_ev if scc_ev else scc_esc
        valor_actual_ev = datos[grid].get(scc_destino_evap, 0.0)
        valor_final_ev, es_salt_ev = proteger_saltillo(
            grid, valor_evap_g, valor_actual_ev, grid_lon
        )

        if es_salt_ev:
            # Saltillo → no escribir, conservar valor original del CSV
            celdas_protegidas += 1
        else:
            datos[grid][scc_destino_evap] = datos[grid].get(scc_destino_evap, 0.0) + valor_final_ev
            celdas_actualizadas += 1

    # Reordenar SCCs
    sccs_list = [s for s in SCCS_TODOS if s in sccs_list]
    for s in set(datos.get(list(datos.keys())[0], {}).keys()) - set(sccs_list):
        sccs_list.append(s)

    if not dry_run:
        escribir_csv_vehicular(output_path, encabezado_r1, sccs_list, datos, campo_final)

    return celdas_actualizadas, celdas_protegidas


def actualizar_contaminante_escape_acumulado(df_excel, csv_path, output_path,
                                              col_excel, grid_lon, factor=1.0,
                                              dry_run=False):
    """
    Actualiza un M_*.csv reseteando primero Monterrey y luego acumulando.
    Necesario porque múltiples vehículos comparten SCCs.
    """
    csv_name = os.path.basename(csv_path)

    if not os.path.exists(csv_path):
        print(f"  ⚠️  {csv_name} no encontrado — omitiendo")
        return 0, 0

    if col_excel not in df_excel.columns:
        print(f"  ⚠️  Columna '{col_excel}' no encontrada — omitiendo")
        return 0, 0

    # Leer CSV original
    encabezado_r1, sccs_list, datos, campo_final = leer_csv_vehicular(csv_path)

    # Resetear celdas de Monterrey para SCCs que vamos a escribir
    sccs_escape_usados = set(v["scc_escape"] for v in VEHICULOS.values())
    for grid in datos:
        if grid in grid_lon and grid_lon[grid] > -100.67:
            for scc in sccs_escape_usados:
                if scc in datos[grid]:
                    datos[grid][scc] = 0.0

    celdas_actualizadas = 0
    celdas_protegidas = 0

    for _, row in df_excel.iterrows():
        grid = int(row['GRIDCODE'])
        vehiculo = row['Vehículo']

        if vehiculo not in VEHICULOS:
            continue

        valor_mg = row[col_excel]
        if pd.isna(valor_mg):
            valor_mg = 0.0

        scc = VEHICULOS[vehiculo]["scc_escape"]
        valor_g = float(valor_mg) * FACTOR_CONVERSION * factor

        # Asegurar SCC existe
        if scc not in sccs_list:
            sccs_list.append(scc)
            for g in datos:
                datos[g][scc] = 0.0

        if grid not in datos:
            datos[grid] = {s: 0.0 for s in sccs_list}

        # Protección Saltillo (estricta): Saltillo nunca se sobreescribe
        valor_actual = datos[grid].get(scc, 0.0)
        valor_final, fue_protegida = proteger_saltillo(grid, valor_g, valor_actual, grid_lon)

        if fue_protegida:
            # Saltillo → no escribir, conservar valor original del CSV
            celdas_protegidas += 1
        else:
            # Acumular (múltiples vehículos pueden compartir SCC)
            datos[grid][scc] = datos[grid].get(scc, 0.0) + valor_final
            celdas_actualizadas += 1

    # Reordenar SCCs
    sccs_list = [s for s in SCCS_TODOS if s in sccs_list]
    for s in set(datos.get(list(datos.keys())[0], {}).keys()) - set(sccs_list):
        sccs_list.append(s)

    if not dry_run:
        escribir_csv_vehicular(output_path, encabezado_r1, sccs_list, datos, campo_final)

    return celdas_actualizadas, celdas_protegidas


# ============================================================
# MAIN
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Actualizar emisiones vehiculares M_*.csv desde FM_emis_vehiculo_muni_AMv2.xlsx"
    )
    parser.add_argument("--dir-excel", default="datos/excel",
                        help="Directorio con el archivo Excel")
    parser.add_argument("--dir-csv", default="datos/csv",
                        help="Directorio con CSVs originales")
    parser.add_argument("--dir-output", default="csv_actualizados",
                        help="Directorio de salida")
    parser.add_argument("--dir-localiza", default="datos/csv",
                        help="Directorio con localiza.csv")
    parser.add_argument("--dry-run", action="store_true",
                        help="Simular sin escribir archivos")
    args = parser.parse_args()

    print("=" * 70)
    print("ACTUALIZACIÓN DE EMISIONES VEHICULARES — AMM 2023")
    print(f"Fuente: FM_emis_vehiculo_muni_AMv2.xlsx")
    print(f"Vehículos: {len(VEHICULOS)} tipos")
    print(f"SCCs: 8 escape + 5 evaporativos = 13 únicos")
    print(f"Contaminantes: 10 regulares + NOx→NO/NO2 + VOC(TGH) = 13 archivos")
    print("=" * 70)

    # Cargar localiza.csv
    ruta_localiza = os.path.join(args.dir_localiza, "localiza.csv")
    grid_lon = cargar_localiza(ruta_localiza)
    print(f"\n✓ localiza.csv: {len(grid_lon)} celdas")

    # Cargar Excel
    excel_path = os.path.join(args.dir_excel, "FM_emis_vehiculo_muni_AMv2.xlsx")
    if not os.path.exists(excel_path):
        print(f"✗ ERROR: {excel_path} no encontrado")
        sys.exit(1)

    df_excel = pd.read_excel(excel_path, sheet_name="FM_emis_vehiculo_muni_AM")
    print(f"✓ Excel: {len(df_excel)} filas, {df_excel['GRIDCODE'].nunique()} GRIDCODEs, "
          f"{df_excel['Vehículo'].nunique()} vehículos")

    # Crear directorio de salida
    os.makedirs(args.dir_output, exist_ok=True)

    total_actualizadas = 0
    total_protegidas = 0

    # ── CONTAMINANTES REGULARES (escape) ──
    print(f"\n{'─' * 70}")
    print("CONTAMINANTES REGULARES (8)")
    print(f"{'─' * 70}")

    for nombre, info in CONTAMINANTES_ESCAPE.items():
        csv_path = os.path.join(args.dir_csv, info["csv"])
        output_path = os.path.join(args.dir_output, info["csv"])

        # Copiar original si existe
        if os.path.exists(csv_path) and not os.path.exists(output_path):
            shutil.copy2(csv_path, output_path)

        act, prot = actualizar_contaminante_escape_acumulado(
            df_excel, csv_path if not os.path.exists(output_path) else output_path,
            output_path, info["col_excel"], grid_lon,
            dry_run=args.dry_run
        )
        total_actualizadas += act
        total_protegidas += prot
        print(f"  ✓ {nombre:6s} → {info['csv']:12s}  "
              f"({act:,} actualizadas, {prot:,} protegidas)")

    # ── NOx → NO + NO2 ──
    print(f"\n{'─' * 70}")
    print("DESCOMPOSICIÓN NOx → NO + NO₂")
    print(f"{'─' * 70}")

    for nombre, info in CONTAMINANTES_NOX.items():
        csv_path = os.path.join(args.dir_csv, info["csv"])
        output_path = os.path.join(args.dir_output, info["csv"])

        if os.path.exists(csv_path) and not os.path.exists(output_path):
            shutil.copy2(csv_path, output_path)

        act, prot = actualizar_contaminante_escape_acumulado(
            df_excel, csv_path if not os.path.exists(output_path) else output_path,
            output_path, info["col_excel"], grid_lon,
            factor=info["fraccion"], dry_run=args.dry_run
        )
        total_actualizadas += act
        total_protegidas += prot
        frac_str = f"×{info['fraccion']:.5f}"
        print(f"  ✓ {nombre:6s} → {info['csv']:12s}  "
              f"({act:,} actualizadas, {prot:,} protegidas) {frac_str}")

    # ── VOC (TGHEvap + TGHNoEva) ──
    print(f"\n{'─' * 70}")
    print("VOC: TGHNoEva → SCC escape, TGHEvap → SCC evaporativo")
    print(f"{'─' * 70}")

    csv_path_voc = os.path.join(args.dir_csv, CONTAMINANTE_VOC["csv"])
    output_path_voc = os.path.join(args.dir_output, CONTAMINANTE_VOC["csv"])

    if os.path.exists(csv_path_voc) and not os.path.exists(output_path_voc):
        shutil.copy2(csv_path_voc, output_path_voc)

    act, prot = actualizar_voc(
        df_excel,
        csv_path_voc if not os.path.exists(output_path_voc) else output_path_voc,
        output_path_voc, grid_lon, dry_run=args.dry_run
    )
    total_actualizadas += act
    total_protegidas += prot
    print(f"  ✓ VOC    → {CONTAMINANTE_VOC['csv']:12s}  "
          f"({act:,} actualizadas, {prot:,} protegidas)")

    # ── RESUMEN ──
    print(f"\n{'=' * 70}")
    print("RESUMEN")
    print(f"{'=' * 70}")
    print(f"  Celdas actualizadas: {total_actualizadas:,}")
    print(f"  Celdas protegidas (Saltillo): {total_protegidas:,}")
    print(f"  Archivos CSV procesados: 11")
    if args.dry_run:
        print("  ⚠️  MODO DRY-RUN: no se escribieron archivos")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()

