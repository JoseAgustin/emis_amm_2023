#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_emisiones_area.py
============================
Script para actualizar emisiones de AREA en el inventario AMM 2023.
Lee 9 archivos Excel y actualiza los archivos A*_2023.csv.

Fuentes: Maquinaria construccion, Agricultura (7 sub), Evaporativas (9 cat),
         Aguas residuales, Erosion, Aeropuerto, Ferrocarril, Maq. agricola,
         Combustion comercial.

Repositorio: emis_amm_2023
Autor: Jose Agustin Garcia Reynoso - ICAyCC, UNAM
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

# ============================================================
# DICCIONARIO DE FUENTES Y OPERACIONES
# ============================================================
FUENTES = {
    "FM_MaqCons_malla": {
        "excel": "FM_MaqCons_malla.xlsx",
        "hoja": "FM_MaqCons_malla",
        "operaciones": [
            {"col": "SUM_PM10_AW", "scc": "2311010000", "csv": "APM10_2023.csv"},
            {"col": "SUM_PM25_AW", "scc": "2311010000", "csv": "APM25_2023.csv"},
            {"col": "SUM_SO2_AW", "scc": "2311010000", "csv": "ASO2_2023.csv"},
            {"col": "SUM_CO_AW",  "scc": "2311010000", "csv": "ACO__2023.csv"},
            {"col": "SUM_NOX_AW", "scc": "2311010000", "csv": "ANOx_2023.csv"},
            {"col": "SUM_COV_AW", "scc": "2311010000", "csv": "AVOC_2023.csv"},
        ]
    },
    "FA_Agro_UDS_malla3x3": {
        "excel": "FA_Agro_UDS_malla3x3.xlsx",
        "hoja": "FA_Agro_UDS_malla3x3",
        "operaciones": [
            # F1.1 Combustion agricola -> SCC 2270005000
            {"col": "SUM_F1_1_PM10_AW", "scc": "2270005000", "csv": "APM10_2023.csv"},
            {"col": "SUM_F1_1_PM2_5_AW", "scc": "2270005000", "csv": "APM25_2023.csv"},
            {"col": "SUM_F1_1_SO2_AW",  "scc": "2270005000", "csv": "ASO2_2023.csv"},
            {"col": "SUM_F1_1_CO_AW",   "scc": "2270005000", "csv": "ACO__2023.csv"},
            {"col": "SUM_F1_1_NOX_AW",  "scc": "2270005000", "csv": "ANOx_2023.csv"},
            {"col": "SUM_F1_1_COV_AW",  "scc": "2270005000", "csv": "AVOC_2023.csv"},
            {"col": "SUM_F1_1_NH3_AW",  "scc": "2270005000", "csv": "ANH3_2023.csv"},
            # F3.1 Distribucion GLP -> SCC 3333333333
            {"col": "SUM_F3_1_COV_AW",  "scc": "3333333333", "csv": "AVOC_2023.csv"},
            # F5.1 Fertilizantes -> SCC 2801700000
            {"col": "SUM_F5_1_NH3_AW",  "scc": "2801700000", "csv": "ANH3_2023.csv"},
            # F5.2 Plaguicidas -> SCC 2461850000
            {"col": "SUM_F5_2_COV_AW",  "scc": "2461850000", "csv": "AVOC_2023.csv"},
            # F5.3 Ganaderia -> SCC 2805020000
            {"col": "SUM_F5_3_NH3_AW",  "scc": "2805020000", "csv": "ANH3_2023.csv"},
            {"col": "SUM_F5_3_CH4_AW",  "scc": "2805020000", "csv": "ACH4_2023.csv"},
            # F5.4 Corrales de engorda -> SCC 2805000000
            {"col": "SUM_F5_4_PM10_AW", "scc": "2805000000", "csv": "APM10_2023.csv"},
            {"col": "SUM_F5_4_PM25_AW", "scc": "2805000000", "csv": "APM25_2023.csv"},
            # F5.5 Labranza -> SCC 2801000002
            {"col": "SUM_F5_5_PM10_AW", "scc": "2801000002", "csv": "APM10_2023.csv"},
            {"col": "SUM_F5_5_PM25_AW", "scc": "2801000002", "csv": "APM25_2023.csv"},
            # F5.6 Quemas agricolas -> SCC 2801500100
            {"col": "SUM_F5_6_PM10_AW", "scc": "2801500100", "csv": "APM10_2023.csv"},
            {"col": "SUM_F5_6_PM25_AW", "scc": "2801500100", "csv": "APM25_2023.csv"},
            {"col": "SUM_F5_6_SO2_AW",  "scc": "2801500100", "csv": "ASO2_2023.csv"},
            {"col": "SUM_F5_6_CO_AW",   "scc": "2801500100", "csv": "ACO__2023.csv"},
            {"col": "SUM_F5_6_NOX_AW",  "scc": "2801500100", "csv": "ANOx_2023.csv"},
            {"col": "SUM_F5_6_COV_AW",  "scc": "2801500100", "csv": "AVOC_2023.csv"},
            {"col": "SUM_F5_6_NH3_AW",  "scc": "2801500100", "csv": "ANH3_2023.csv"},
        ]
    },
    "FA_evaporativas_9cat_malla": {
        "excel": "FA_evaporativas_9cat_malla.xlsx",
        "hoja": "FA_evaporativas_9cat_malla",
        "operaciones": [
            {"col": "F21ArtGraf",       "scc": "2425000000", "csv": "AVOC_2023.csv"},
            {"col": "F23LavSeco",       "scc": "2420000055", "csv": "AVOC_2023.csv"},
            {"col": "F24LimSupInd",     "scc": "2415000000", "csv": "AVOC_2023.csv"},
            {"col": "F25PintAuto",      "scc": "2401005000", "csv": "AVOC_2023.csv"},
            {"col": "F28RecSupInd",     "scc": "2401002000", "csv": "AVOC_2023.csv"},
            {"col": "F31GLPCom",        "scc": "3333333333", "csv": "AVOC_2023.csv"},
            {"col": "F31GLPCom",        "scc": "2390007000", "csv": "AVOC_2023.csv"},
            {"col": "F32ManDisGasolina","scc": "2501060000", "csv": "AVOC_2023.csv"},
            {"col": "F43Panifica",      "scc": "2302050000", "csv": "AVOC_2023.csv"},
        ]
    },
    "FA_AguasRes_malla": {
        "excel": "FA_AguasRes_RelleSant_malla.xlsx",
        "hoja": "FA_AguasRes_RelleSant_malla",
        "operaciones": [
            {"col": "COV", "scc": "2630030000", "csv": "AVOC_2023.csv"},
            {"col": "CH4", "scc": "2630030000", "csv": "ACH4_2023.csv"},
        ]
    },
    "FA_erosivas_malla": {
        "excel": "FA_Erosivas_PM_UDS_malla.xlsx",
        "hoja": "FA_Erosivas_PM_UDS_malla",
        "operaciones": [
            {"col": "SUM_PM10_AW", "scc": "2730100000", "csv": "APM10_2023.csv"},
            {"col": "SUM_PM25_AW", "scc": "2730100000", "csv": "APM25_2023.csv"},
        ]
    },
    "FM_Aero_CAutobus_malla": {
        "excel": "FM_Aero_CAutobus_malla.xlsx",
        "hoja": "FM_Aero_CAutobus_malla",
        "operaciones": [
            {"col": "SUM_PM10", "scc": "2270008000", "csv": "APM10_2023.csv"},
            {"col": "SUM_PM25", "scc": "2270008000", "csv": "APM25_2023.csv"},
            {"col": "SUM_SO2",  "scc": "2270008000", "csv": "ASO2_2023.csv"},
            {"col": "SUM_CO",   "scc": "2270008000", "csv": "ACO__2023.csv"},
            {"col": "SUM_NOX",  "scc": "2270008000", "csv": "ANOx_2023.csv"},
            {"col": "SUM_COV",  "scc": "2270008000", "csv": "AVOC_2023.csv"},
        ]
    },
    "FM_ferrocarril_emis_malla": {
        "excel": "FM_ferrocarril_emis_malla.xlsx",
        "hoja": "FM_ferrocarril_emis_malla",
        "operaciones": [
            {"col": "SUM_PM10_LW", "scc": "2285000000", "csv": "APM10_2023.csv"},
            {"col": "SUM_PM25_LW", "scc": "2285000000", "csv": "APM25_2023.csv"},
            {"col": "SUM_SO2_LW",  "scc": "2285000000", "csv": "ASO2_2023.csv"},
            {"col": "SUM_CO_LW",   "scc": "2285000000", "csv": "ACO__2023.csv"},
            {"col": "SUM_NOX_LW",  "scc": "2285000000", "csv": "ANOx_2023.csv"},
            {"col": "SUM_COV_LW",  "scc": "2285000000", "csv": "AVOC_2023.csv"},
        ]
    },
    "FM_maquinaria_agricola_malla": {
        "excel": "FM_maquinaria_agricola_malla.xlsx",
        "hoja": "FM_maquinaria_agricola_malla",
        "operaciones": [
            {"col": "SUM_PM10", "scc": "2265005000", "csv": "APM10_2023.csv"},
            {"col": "SUM_PM2_5","scc": "2265005000", "csv": "APM25_2023.csv"},
            {"col": "SUM_SO2",  "scc": "2265005000", "csv": "ASO2_2023.csv"},
            {"col": "SUM_CO",   "scc": "2265005000", "csv": "ACO__2023.csv"},
            {"col": "SUM_NOX",  "scc": "2265005000", "csv": "ANOx_2023.csv"},
            {"col": "SUM_COV",  "scc": "2265005000", "csv": "AVOC_2023.csv"},
            {"col": "SUM_NH3",  "scc": "2265005000", "csv": "ANH3_2023.csv"},
            {"col": "SUM_CN",   "scc": "2265005000", "csv": "ACN__2023.csv"},
        ]
    },
    "FA_comb_comer_regulada_malla": {
        "excel": "FA_comb_comer_regulada_malla.xlsx",
        "hoja": "FA_comb_comer_regulada_malla",
        "operaciones": [
            # Combustion -> SCC 2103007000
            {"col": "PM10",        "scc": "2103007000", "csv": "APM10_2023.csv"},
            {"col": "PM2_5",       "scc": "2103007000", "csv": "APM25_2023.csv"},
            {"col": "SO2",         "scc": "2103007000", "csv": "ASO2_2023.csv"},
            {"col": "CO",          "scc": "2103007000", "csv": "ACO__2023.csv"},
            {"col": "NOx",         "scc": "2103007000", "csv": "ANOx_2023.csv"},
            {"col": "COTComb",     "scc": "2103007000", "csv": "AVOC_2023.csv"},
            {"col": "NH3",         "scc": "2103007000", "csv": "ANH3_2023.csv"},
            {"col": "CN",          "scc": "2103007000", "csv": "ACN__2023.csv"},
            # Evaporacion -> SCC 2104011000
            {"col": "COVProc",   "scc": "2103004000", "csv": "AVOC_2023.csv"},
        ]
    },
}


# ============================================================
# FUNCIONES DE LECTURA/ESCRITURA CSV AREA
# ============================================================
def leer_csv_area(ruta):
    """Lee un archivo A*_2023.csv con formato Fortran de area."""
    with open(ruta, 'r') as f:
        lines = f.readlines()

    r2 = lines[1].strip().rstrip(',').split(',')
    r2_clean = [s.strip() for s in r2 if s.strip()]
    scc_count = int(r2_clean[0])
    scc_names = r2_clean[2:]  # saltar N y g_per_year

    data_rows = []
    for line in lines[2:]:
        stripped = line.strip().rstrip(',')
        if not stripped:
            continue
        parts = [p.strip() for p in stripped.split(',') if p.strip()]
        if len(parts) < 4:
            continue
        grid = int(parts[0])
        cid = int(parts[1])
        furb = float(parts[2])
        frural = float(parts[3])
        emisiones = [float(p) for p in parts[4:]]
        data_rows.append([grid, cid, furb, frural] + emisiones)

    col_names = ['grid', 'CID', 'Furb', 'Frural'] + scc_names
    df = pd.DataFrame(data_rows, columns=col_names[:len(data_rows[0])])

    for scc in scc_names:
        if scc not in df.columns:
            df[scc] = 0.0

    df['grid'] = df['grid'].astype(int)
    df['CID'] = df['CID'].astype(int)
    return df, scc_names, scc_count, lines


def escribir_csv_area(ruta, df, lines_orig):
    """Escribe CSV de area con formato Fortran compatible."""
    all_sccs = [c for c in df.columns if c not in ('grid', 'CID', 'Furb', 'Frural')]
    nuevo_count = len(all_sccs)

    with open(ruta, 'w') as f:
        f.write(lines_orig[0])
        scc_list = ",".join(all_sccs)
        f.write(f"{nuevo_count},g_per_year,{scc_list},\n")
        for _, row in df.iterrows():
            grid = int(row['grid'])
            cid = int(row['CID'])
            furb = row['Furb']
            frural = row['Frural']
            parts = [f"{grid}", f"{cid:>6d}", f" {furb:.4f}", f" {frural:.4f}"]
            for scc in all_sccs:
                val = row[scc]
                parts.append(f" {val:.5E}")
            f.write(",".join(parts) + ",\n")
    return nuevo_count


# ============================================================
# FUNCION PRINCIPAL
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description='Actualizar emisiones de area del inventario AMM 2023'
    )
    parser.add_argument('--dir-excel', default='datos/excel/',
                        help='Directorio con archivos Excel (default: datos)')
    parser.add_argument('--dir-csv', default='datos/csv/',
                        help='Directorio con CSVs originales (default: csv_originales)')
    parser.add_argument('--dir-output', default='csv_actualizados',
                        help='Directorio de salida (default: csv_actualizados)')
    parser.add_argument('--dir-localiza', default='datos/csv/',
                        help='Directorio con localiza.csv (default: datos)')
    parser.add_argument('--fuente', default=None,
                        help='Procesar solo esta fuente (default: todas)')
    parser.add_argument('--csv', default=None,
                        help='Procesar solo este CSV (default: todos)')
    parser.add_argument('--dry-run', action='store_true',
                        help='Simular sin escribir archivos')
    args = parser.parse_args()

    print("=" * 70)
    print("ACTUALIZACION DE EMISIONES DE AREA - INVENTARIO AMM 2023")
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Fuentes: {len(FUENTES)}")
    total_ops = sum(len(f['operaciones']) for f in FUENTES.values())
    print(f"Operaciones totales: {total_ops}")
    print("=" * 70)

    # Cargar localiza.csv
    ruta_localiza = os.path.join(args.dir_localiza, 'localiza.csv')
    grid_lon = cargar_localiza(ruta_localiza)
    print(f"\nlocaliza.csv: {len(grid_lon)} celdas")

    os.makedirs(args.dir_output, exist_ok=True)

    csv_cache = {}
    total_ops_ok = 0
    total_ops_skip = 0
    total_celdas_mty = 0
    total_celdas_saltillo = 0

    for fuente_name, fuente_info in FUENTES.items():
        if args.fuente and args.fuente.lower() not in fuente_name.lower():
            continue

        excel_path = os.path.join(args.dir_excel, fuente_info['excel'])
        if not os.path.exists(excel_path):
            print(f"\n  {fuente_info['excel']}: No encontrado - omitiendo")
            continue

        print(f"\n{'='*70}")
        print(f"FUENTE: {fuente_name}")
        print(f"  Excel: {fuente_info['excel']}")

        try:
            df_excel = pd.read_excel(excel_path, sheet_name=fuente_info['hoja'])
        except Exception as e:
            print(f"  Error leyendo Excel: {e}")
            continue

        col_grid = 'GRIDCODE' if 'GRIDCODE' in df_excel.columns else 'grid'
        df_excel[col_grid] = df_excel[col_grid].astype(int)
        print(f"  Celdas Excel: {len(df_excel)}")

        for op in fuente_info['operaciones']:
            col_excel = op['col']
            scc = op['scc']
            csv_name = op['csv']

            if args.csv and args.csv != csv_name:
                continue

            if col_excel not in df_excel.columns:
                print(f"    Columna '{col_excel}' no encontrada - omitiendo")
                total_ops_skip += 1
                continue

            csv_path = os.path.join(args.dir_output, csv_name)
            src_path = os.path.join(args.dir_csv, csv_name)

            if not os.path.exists(csv_path):
                if os.path.exists(src_path):
                    shutil.copy2(src_path, csv_path)
                else:
                    print(f"    {csv_name} no encontrado - omitiendo")
                    total_ops_skip += 1
                    continue

            if csv_path not in csv_cache:
                try:
                    df_csv, scc_names, scc_count, lines_orig = leer_csv_area(csv_path)
                    csv_cache[csv_path] = {
                        'df': df_csv, 'scc_names': scc_names,
                        'scc_count': scc_count, 'lines_orig': lines_orig
                    }
                except Exception as e:
                    print(f"    Error leyendo {csv_name}: {e}")
                    total_ops_skip += 1
                    continue

            cache = csv_cache[csv_path]
            df_csv = cache['df']

            if scc not in df_csv.columns:
                df_csv[scc] = 0.0

            csv_grid_idx = dict(zip(df_csv['grid'], df_csv.index))

            celdas_mty = 0
            celdas_salt = 0

            for _, row_excel in df_excel.iterrows():
                grid = int(row_excel[col_grid])
                valor = row_excel[col_excel]

                if pd.isna(valor):
                    valor = 0.0
                else:
                    valor = float(valor) * FACTOR_CONVERSION

                if grid in csv_grid_idx:
                    idx = csv_grid_idx[grid]
                    valor_actual = df_csv.at[idx, scc]

                    valor_final, fue_protegida = proteger_saltillo(
                        grid, valor, valor_actual, grid_lon
                    )
                    df_csv.at[idx, scc] = valor_final

                    if fue_protegida:
                        celdas_salt += 1
                    else:
                        celdas_mty += 1

            total_celdas_mty += celdas_mty
            total_celdas_saltillo += celdas_salt
            total_ops_ok += 1

            print(f"    {col_excel:25s} -> {scc} -> {csv_name:16s} "
                  f"({celdas_mty:,} MTY, {celdas_salt:,} Salt.)")

    # Escribir CSVs actualizados
    print(f"\n{'='*70}")
    print("ESCRITURA DE CSVs ACTUALIZADOS")
    print(f"{'='*70}")

    if not args.dry_run:
        for csv_path, cache in csv_cache.items():
            csv_name = os.path.basename(csv_path)
            nuevo_count = escribir_csv_area(csv_path, cache['df'], cache['lines_orig'])
            delta = nuevo_count - cache['scc_count']
            status = f"+{delta}" if delta > 0 else "="
            print(f"  {csv_name:16s}: {cache['scc_count']:2d} -> {nuevo_count:2d} SCCs ({status})")
    else:
        print("  [DRY-RUN] No se escribieron archivos")

    # Resumen final
    print(f"\n{'='*70}")
    print("RESUMEN FINAL")
    print(f"{'='*70}")
    print(f"  Operaciones exitosas:     {total_ops_ok}")
    print(f"  Operaciones omitidas:     {total_ops_skip}")
    print(f"  Celdas Monterrey:         {total_celdas_mty:,}")
    print(f"  Celdas Saltillo prot.:    {total_celdas_saltillo:,}")
    print(f"  CSVs actualizados:        {len(csv_cache)}")
    print(f"  Directorio salida:        {args.dir_output}/")


if __name__ == '__main__':
    main()
