#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_emisiones_carreteras.py
==================================
Script para actualizar emisiones VEHICULARES (carreteras) en el inventario AMM 2023.
Lee FM_carreteras_emis_malla.xlsx y actualiza los archivos M_*.csv.

Estructura de M_*.csv (formato Fortran):
  R1: ' GRIDCODE, emissions in gram per year,,,,,,,,,,,,,,,,,'
  R2: '      16, 2201001330, 2201020330, ..., 220108033V'
  R3+: '  GRIDCODE,  X.XXXXE+XX, ...,  X.XXXXE+XX, 6'

Caracteristicas:
- 16 SCCs: 11 numericos (sufijo 330) + 5 evaporativos (sufijo 33V)
- Proporciones UNICAS por contaminante (verificadas contra diccionario)
- SCCs evaporativos (33V) solo aplican a COV
- NOx se descompone en NO y NO2 con relacion NO/NO2 = 9.46
- Proteccion Saltillo integrada (LONG <= -100.67)
- 8 contaminantes: COV, CO, NOx, SO2, PM10, PM2.5, NH3, CN

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
RATIO_NO_NO2 = 9.46

FRAC_NO2 = 1.0 / (1.0 + RATIO_NO_NO2)
FRAC_NO  = RATIO_NO_NO2 / (1.0 + RATIO_NO_NO2)

CAMPO_FINAL_DEFAULT = 6

SCCS_NUMERICOS = [
    "2201001330", "2201020330", "2201040330", "2201070330", "2201080330",
    "2230001330", "2230060330", "2230071330", "2230073330", "2230074330",
    "2230075330",
]
SCCS_EVAPORATIVOS = [
    "220100133V", "220102033V", "220104033V", "220107033V", "220108033V",
]
SCCS_TODOS = SCCS_NUMERICOS + SCCS_EVAPORATIVOS

PROPORCIONES = {
    "COV": {
        "2201001330": 0.015023103, "2201020330": 0.156099105, "2201040330": 0.165576533,
        "2201070330": 0.000000000, "2201080330": 0.254777964, "2230001330": 0.000000000,
        "2230060330": 0.000000000, "2230071330": 0.229760167, "2230073330": 0.000000000,
        "2230074330": 0.096176200, "2230075330": 0.071459715,
        "220100133V": 0.005301532, "220102033V": 0.003520622,
        "220104033V": 0.000955741, "220107033V": 0.001006778, "220108033V": 0.000342539,
    },
    "CO": {
        "2201001330": 0.005720951, "2201020330": 0.081101311, "2201040330": 0.130747817,
        "2201070330": 0.002022259, "2201080330": 0.163118600, "2230001330": 0.000036974,
        "2230060330": 0.000273760, "2230071330": 0.168232521, "2230073330": 0.018486342,
        "2230074330": 0.160616289, "2230075330": 0.269643176,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
    "NOx": {
        "2201001330": 0.012433851, "2201020330": 0.029709813, "2201040330": 0.029000149,
        "2201070330": 0.000000000, "2201080330": 0.064568811, "2230001330": 0.000000000,
        "2230060330": 0.000000000, "2230071330": 0.214875435, "2230073330": 0.000000000,
        "2230074330": 0.103680028, "2230075330": 0.545731914,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
    "SO2": {
        "2201001330": 0.059148675, "2201020330": 0.172237103, "2201040330": 0.164833410,
        "2201070330": 0.000000000, "2201080330": 0.291503022, "2230001330": 0.000000000,
        "2230060330": 0.000000000, "2230071330": 0.211106260, "2230073330": 0.000000000,
        "2230074330": 0.049955222, "2230075330": 0.051216306,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
    "PM10": {
        "2201001330": 0.024198047, "2201020330": 0.133541264, "2201040330": 0.188139682,
        "2201070330": 0.000000000, "2201080330": 0.291458459, "2230001330": 0.000000000,
        "2230060330": 0.000000000, "2230071330": 0.214168032, "2230073330": 0.000000000,
        "2230074330": 0.092327891, "2230075330": 0.056166625,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
    "PM25": {
        "2201001330": 0.009247658, "2201020330": 0.107177403, "2201040330": 0.173341755,
        "2201070330": 0.000000000, "2201080330": 0.256409776, "2230001330": 0.000000000,
        "2230060330": 0.000000000, "2230071330": 0.200485996, "2230073330": 0.000000000,
        "2230074330": 0.082700699, "2230075330": 0.170636712,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
    "NH3": {
        "2201001330": 0.049647083, "2201020330": 0.164358934, "2201040330": 0.158561211,
        "2201070330": 0.000000000, "2201080330": 0.317135396, "2230001330": 0.000000000,
        "2230060330": 0.000000000, "2230071330": 0.056435585, "2230073330": 0.000000000,
        "2230074330": 0.019719183, "2230075330": 0.234142607,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
    "CN": {
        "2201001330": 0.005720951, "2201020330": 0.081101311, "2201040330": 0.130747817,
        "2201070330": 0.002022259, "2201080330": 0.163118600, "2230001330": 0.000036974,
        "2230060330": 0.000273760, "2230071330": 0.168232521, "2230073330": 0.018486342,
        "2230074330": 0.160616289, "2230075330": 0.269643176,
        "220100133V": 0.000000000, "220102033V": 0.000000000,
        "220104033V": 0.000000000, "220107033V": 0.000000000, "220108033V": 0.000000000,
    },
}

CONTAMINANTES = {
    "COV":  {"col_excel": "SUM_COV_LW",  "csv": "M_VOC.csv"},
    "CO":   {"col_excel": "SUM_CO_LW",   "csv": "M_CO.csv"},
    "SO2":  {"col_excel": "SUM_SO2_LW",  "csv": "M_SO2.csv"},
    "PM10": {"col_excel": "SUM_PM10_LW", "csv": "M_PM10.csv"},
    "PM25": {"col_excel": "SUM_PM25_LW", "csv": "M_PM25.csv"},
    "NH3":  {"col_excel": "SUM_NH3_LW",  "csv": "M_NH3.csv"},
    "CN":   {"col_excel": "SUM_CN_LW",   "csv": "M_CN.csv"},
}

NOX_CONFIG = {
    "col_excel": "SUM_NOx_LW",
    "NO":  {"csv": "M_NO.csv",  "fraccion": FRAC_NO},
    "NO2": {"csv": "M_NO2.csv", "fraccion": FRAC_NO2},
}


def leer_csv_vehicular(ruta):
    """Lee un archivo M_*.csv con formato Fortran vehicular."""
    with open(ruta, 'r') as f:
        lines = f.readlines()

    r2_raw = lines[1].strip()
    r2_parts = [s.strip() for s in r2_raw.split(',') if s.strip()]
    scc_count = int(r2_parts[0])
    scc_names = r2_parts[1:]

    data_rows = []
    campo_final = CAMPO_FINAL_DEFAULT

    for line in lines[2:]:
        stripped = line.strip()
        if not stripped:
            continue
        campos = [p.strip() for p in stripped.split(',') if p.strip()]
        if len(campos) < 2:
            continue

        ultimo = campos[-1]
        tiene_campo_final = False
        try:
            posible_final = int(ultimo)
            if 'E' not in ultimo.upper() and '.' not in ultimo and posible_final < 100:
                campo_final = posible_final
                tiene_campo_final = True
        except ValueError:
            pass

        gridcode = int(campos[0])
        if tiene_campo_final:
            emisiones = [float(p) for p in campos[1:-1]]
        else:
            emisiones = [float(p) for p in campos[1:]]

        data_rows.append([gridcode] + emisiones)

    if not data_rows:
        raise ValueError(f"No se encontraron datos en {ruta}")

    n_emis = len(data_rows[0]) - 1
    if n_emis > len(scc_names):
        for i in range(len(scc_names), n_emis):
            scc_names.append(f"SCC_EXTRA_{i+1}")
    elif n_emis < len(scc_names):
        scc_names = scc_names[:n_emis]

    col_names = ['GRIDCODE'] + scc_names
    df = pd.DataFrame(data_rows, columns=col_names)
    df['GRIDCODE'] = df['GRIDCODE'].astype(int)
    for scc in scc_names:
        df[scc] = pd.to_numeric(df[scc], errors='coerce').fillna(0.0)
    df = df[df['GRIDCODE'] > 0].reset_index(drop=True)
    return df, scc_names, scc_count, lines, campo_final


def escribir_csv_vehicular(ruta, df, lines_orig, campo_final=6):
    """Escribe CSV vehicular con formato Fortran compatible."""
    all_sccs = [c for c in df.columns if c != 'GRIDCODE']
    nuevo_count = len(all_sccs)

    with open(ruta, 'w') as f:
        f.write(lines_orig[0])
        scc_list = ",".join(all_sccs)
        f.write(f"      {nuevo_count},{scc_list}\n")
        for _, row in df.iterrows():
            gridcode = int(row['GRIDCODE'])
            parts = [f"{gridcode}"]
            for scc in all_sccs:
                val = row[scc]
                parts.append(f"{val:.4E}")
            parts.append(f"{campo_final}")
            f.write(",".join(parts) + "\n")
    return nuevo_count


def main():
    parser = argparse.ArgumentParser(
        description='Actualizar emisiones vehiculares (carreteras) del inventario AMM 2023'
    )
    parser.add_argument('--dir-excel', default='datos')
    parser.add_argument('--dir-csv', default='csv_originales')
    parser.add_argument('--dir-output', default='csv_actualizados')
    parser.add_argument('--dir-localiza', default='datos')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    print("=" * 70)
    print("ACTUALIZACION DE EMISIONES VEHICULARES - INVENTARIO AMM 2023")
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Relacion NO/NO2 = {RATIO_NO_NO2}")
    print(f"  Fraccion NO  = {FRAC_NO:.5f}")
    print(f"  Fraccion NO2 = {FRAC_NO2:.5f}")
    print(f"SCCs: {len(SCCS_TODOS)} ({len(SCCS_NUMERICOS)} escape + "
          f"{len(SCCS_EVAPORATIVOS)} evaporativos)")
    print("=" * 70)

    print("\nVerificacion de proporciones:")
    for cont, props in PROPORCIONES.items():
        total = sum(props.values())
        status = "OK" if abs(total - 1.0) < 1e-4 else "ERROR"
        print(f"  {cont:5s}: suma = {total:.6f} {status}")

    ruta_localiza = os.path.join(args.dir_localiza, 'localiza.csv')
    grid_lon = cargar_localiza(ruta_localiza)
    print(f"\nlocaliza.csv: {len(grid_lon)} celdas")

    excel_path = os.path.join(args.dir_excel, 'FM_carreteras_emis_malla.xlsx')
    if not os.path.exists(excel_path):
        print(f"ERROR: No se encontro {excel_path}")
        sys.exit(1)

    df_excel = pd.read_excel(excel_path, sheet_name='FM_carreteras_emis_malla')
    df_excel['GRIDCODE'] = df_excel['GRIDCODE'].astype(int)
    print(f"Excel carreteras: {len(df_excel)} celdas")

    os.makedirs(args.dir_output, exist_ok=True)

    csv_cache = {}
    total_ops = 0
    total_celdas_mty = 0
    total_celdas_saltillo = 0

    print(f"\n{'_'*70}")
    print("CONTAMINANTES REGULARES (7)")
    print(f"{'_'*70}")

    for cont_name, cont_info in CONTAMINANTES.items():
        col_excel = cont_info['col_excel']
        csv_name = cont_info['csv']
        proporciones = PROPORCIONES[cont_name]

        if col_excel not in df_excel.columns:
            print(f"  Columna '{col_excel}' no encontrada - omitiendo {cont_name}")
            continue

        csv_path = os.path.join(args.dir_output, csv_name)
        src_path = os.path.join(args.dir_csv, csv_name)

        if not os.path.exists(csv_path):
            if os.path.exists(src_path):
                shutil.copy2(src_path, csv_path)
            else:
                print(f"  {csv_name} no encontrado - omitiendo")
                continue

        if csv_path not in csv_cache:
            try:
                df_csv, scc_names, scc_count, lines_orig, campo_final_csv = \
                    leer_csv_vehicular(csv_path)
                csv_cache[csv_path] = {
                    'df': df_csv, 'scc_names': scc_names,
                    'scc_count': scc_count, 'lines_orig': lines_orig,
                    'campo_final': campo_final_csv
                }
            except Exception as e:
                print(f"  Error leyendo {csv_name}: {e}")
                continue

        cache = csv_cache[csv_path]
        df_csv = cache['df']
        csv_grid_idx = dict(zip(df_csv['GRIDCODE'], df_csv.index))

        for scc in proporciones:
            if scc not in df_csv.columns:
                df_csv[scc] = 0.0

        celdas_mty = 0
        celdas_salt = 0

        for _, row_excel in df_excel.iterrows():
            grid = int(row_excel['GRIDCODE'])
            valor_total = row_excel[col_excel]

            if pd.isna(valor_total):
                valor_total = 0.0
            else:
                valor_total = float(valor_total) * FACTOR_CONVERSION

            if grid in csv_grid_idx:
                idx = csv_grid_idx[grid]
                for scc, proporcion in proporciones.items():
                    if proporcion == 0.0:
                        continue
                    valor_scc = valor_total * proporcion
                    valor_actual = df_csv.at[idx, scc]
                    valor_final, fue_protegida = proteger_saltillo(
                        grid, valor_scc, valor_actual, grid_lon
                    )
                    df_csv.at[idx, scc] = valor_final
                    if fue_protegida:
                        celdas_salt += 1
                    else:
                        celdas_mty += 1

        total_celdas_mty += celdas_mty
        total_celdas_saltillo += celdas_salt
        total_ops += 1

        n_activos = sum(1 for p in proporciones.values() if p > 0)
        print(f"  {cont_name:5s} -> {csv_name:12s}  "
              f"({n_activos}/{len(proporciones)} SCCs activos, "
              f"{celdas_mty:,} MTY, {celdas_salt:,} Salt.)")

    print(f"\n{'_'*70}")
    print(f"DESCOMPOSICION NOx -> NO + NO2  (ratio NO/NO2 = {RATIO_NO_NO2})")
    print(f"{'_'*70}")

    col_nox = NOX_CONFIG['col_excel']
    if col_nox not in df_excel.columns:
        print(f"  Columna '{col_nox}' no encontrada - omitiendo NO/NO2")
    else:
        proporciones_nox = PROPORCIONES['NOx']

        for gas, gas_info in [('NO', NOX_CONFIG['NO']), ('NO2', NOX_CONFIG['NO2'])]:
            csv_name = gas_info['csv']
            fraccion = gas_info['fraccion']

            csv_path = os.path.join(args.dir_output, csv_name)
            src_path = os.path.join(args.dir_csv, csv_name)

            if not os.path.exists(csv_path):
                if os.path.exists(src_path):
                    shutil.copy2(src_path, csv_path)
                else:
                    print(f"  {csv_name} no encontrado - omitiendo")
                    continue

            if csv_path not in csv_cache:
                try:
                    df_csv, scc_names, scc_count, lines_orig, campo_final_csv = \
                        leer_csv_vehicular(csv_path)
                    csv_cache[csv_path] = {
                        'df': df_csv, 'scc_names': scc_names,
                        'scc_count': scc_count, 'lines_orig': lines_orig,
                        'campo_final': campo_final_csv
                    }
                except Exception as e:
                    print(f"  Error leyendo {csv_name}: {e}")
                    continue

            cache = csv_cache[csv_path]
            df_csv = cache['df']
            csv_grid_idx = dict(zip(df_csv['GRIDCODE'], df_csv.index))

            for scc in proporciones_nox:
                if scc not in df_csv.columns:
                    df_csv[scc] = 0.0

            celdas_mty = 0
            celdas_salt = 0

            for _, row_excel in df_excel.iterrows():
                grid = int(row_excel['GRIDCODE'])
                valor_nox = row_excel[col_nox]

                if pd.isna(valor_nox):
                    valor_nox = 0.0
                else:
                    valor_nox = float(valor_nox) * FACTOR_CONVERSION

                valor_gas = valor_nox * fraccion

                if grid in csv_grid_idx:
                    idx = csv_grid_idx[grid]
                    for scc, proporcion in proporciones_nox.items():
                        if proporcion == 0.0:
                            continue
                        valor_scc = valor_gas * proporcion
                        valor_actual = df_csv.at[idx, scc]
                        valor_final, fue_protegida = proteger_saltillo(
                            grid, valor_scc, valor_actual, grid_lon
                        )
                        df_csv.at[idx, scc] = valor_final
                        if fue_protegida:
                            celdas_salt += 1
                        else:
                            celdas_mty += 1

            total_celdas_mty += celdas_mty
            total_celdas_saltillo += celdas_salt
            total_ops += 1

            n_activos = sum(1 for p in proporciones_nox.values() if p > 0)
            print(f"  {gas:3s} (frac={fraccion:.5f}) -> {csv_name:12s}  "
                  f"({n_activos}/{len(proporciones_nox)} SCCs activos, "
                  f"{celdas_mty:,} MTY, {celdas_salt:,} Salt.)")

    print(f"\n{'_'*70}")
    print("ESCRITURA DE CSVs ACTUALIZADOS")
    print(f"{'_'*70}")

    if not args.dry_run:
        for csv_path, cache in csv_cache.items():
            csv_name = os.path.basename(csv_path)
            nuevo_count = escribir_csv_vehicular(
                csv_path, cache['df'], cache['lines_orig'], cache['campo_final']
            )
            delta = nuevo_count - cache['scc_count']
            status = f"+{delta}" if delta > 0 else "="
            print(f"  {csv_name:12s}: {cache['scc_count']:2d} -> {nuevo_count:2d} SCCs ({status})")
    else:
        print("  [DRY-RUN] No se escribieron archivos")

    print(f"\n{'='*70}")
    print("RESUMEN FINAL")
    print(f"{'='*70}")
    print(f"  Contaminantes procesados: {total_ops}")
    print(f"  Celdas Monterrey:         {total_celdas_mty:,}")
    print(f"  Celdas Saltillo prot.:    {total_celdas_saltillo:,}")
    print(f"  CSVs actualizados:        {len(csv_cache)}")
    print(f"  Directorio salida:        {args.dir_output}/")


if __name__ == '__main__':
    main()
