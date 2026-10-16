#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_emisiones_carreteras.py
Script para actualizar emisiones vehiculares de carreteras.

1 fuente | 11+ operaciones | 12 SCCs por contaminante → 11 archivos M_*.csv
Conversión: Mg/año → g/año (×1E+06)
Descomposición NOx: NO/NO2 ratio = 9.46
Protección Saltillo: LONGITUD <= -100.67

Uso:
  python actualizar_emisiones_carreteras.py [opciones]
"""

import os
import sys
import argparse
import pandas as pd
import numpy as np
from proteccion_saltillo import cargar_coordenadas, proteger_saltillo, LON_LIMITE

EXCEL_FILE = "FM_carreteras_emis_malla.xlsx"
HOJA = "FM_carreteras_emis_malla"
GRID_EXCEL = "GRIDCODE"
FACTOR_CONVERSION = 1e6
RATIO_NO_NO2 = 9.46

PROPORCIONES = {
    "COV": {
        "2201001000": 0.1847, "2201020000": 0.0923, "2201040000": 0.1847,
        "2201070000": 0.0923, "2202001000": 0.0462, "2202020000": 0.0231,
        "2202040000": 0.0462, "2202070000": 0.0231, "2230001000": 0.0462,
        "2230060000": 0.0462, "2230070000": 0.0462, "2230073000": 0.1688,
    },
    "CO": {
        "2201001000": 0.2105, "2201020000": 0.1053, "2201040000": 0.2105,
        "2201070000": 0.1053, "2202001000": 0.0263, "2202020000": 0.0132,
        "2202040000": 0.0263, "2202070000": 0.0132, "2230001000": 0.0263,
        "2230060000": 0.0263, "2230070000": 0.0263, "2230073000": 0.2105,
    },
    "NOx": {
        "2201001000": 0.0526, "2201020000": 0.0263, "2201040000": 0.0526,
        "2201070000": 0.0263, "2202001000": 0.1316, "2202020000": 0.0658,
        "2202040000": 0.1316, "2202070000": 0.0658, "2230001000": 0.0658,
        "2230060000": 0.0658, "2230070000": 0.0658, "2230073000": 0.2500,
    },
    "SO2": {
        "2201001000": 0.1250, "2201020000": 0.0625, "2201040000": 0.1250,
        "2201070000": 0.0625, "2202001000": 0.0625, "2202020000": 0.0313,
        "2202040000": 0.0625, "2202070000": 0.0313, "2230001000": 0.0625,
        "2230060000": 0.0625, "2230070000": 0.0625, "2230073000": 0.2499,
    },
    "PM10": {
        "2201001000": 0.0833, "2201020000": 0.0417, "2201040000": 0.0833,
        "2201070000": 0.0417, "2202001000": 0.1042, "2202020000": 0.0521,
        "2202040000": 0.1042, "2202070000": 0.0521, "2230001000": 0.0521,
        "2230060000": 0.0521, "2230070000": 0.0521, "2230073000": 0.2811,
    },
    "PM25": {
        "2201001000": 0.0833, "2201020000": 0.0417, "2201040000": 0.0833,
        "2201070000": 0.0417, "2202001000": 0.1042, "2202020000": 0.0521,
        "2202040000": 0.1042, "2202070000": 0.0521, "2230001000": 0.0521,
        "2230060000": 0.0521, "2230070000": 0.0521, "2230073000": 0.2811,
    },
    "NH3": {
        "2201001000": 0.1667, "2201020000": 0.0833, "2201040000": 0.1667,
        "2201070000": 0.0833, "2202001000": 0.0417, "2202020000": 0.0208,
        "2202040000": 0.0417, "2202070000": 0.0208, "2230001000": 0.0417,
        "2230060000": 0.0417, "2230070000": 0.0417, "2230073000": 0.2499,
    },
    "CO2": {
        "2201001000": 0.1250, "2201020000": 0.0625, "2201040000": 0.1250,
        "2201070000": 0.0625, "2202001000": 0.0625, "2202020000": 0.0313,
        "2202040000": 0.0625, "2202070000": 0.0313, "2230001000": 0.0625,
        "2230060000": 0.0625, "2230070000": 0.0625, "2230073000": 0.2499,
    },
    "CH4": {
        "2201001000": 0.1667, "2201020000": 0.0833, "2201040000": 0.1667,
        "2201070000": 0.0833, "2202001000": 0.0417, "2202020000": 0.0208,
        "2202040000": 0.0417, "2202070000": 0.0208, "2230001000": 0.0417,
        "2230060000": 0.0417, "2230070000": 0.0417, "2230073000": 0.2499,
    },
    "CN": {
        "2201001000": 0.0833, "2201020000": 0.0417, "2201040000": 0.0833,
        "2201070000": 0.0417, "2202001000": 0.1042, "2202020000": 0.0521,
        "2202040000": 0.1042, "2202070000": 0.0521, "2230001000": 0.0521,
        "2230060000": 0.0521, "2230070000": 0.0521, "2230073000": 0.2811,
    },
}

CONTAMINANTE_CSV = {
    "COV": "M_VOC.csv", "CO": "M_CO.csv", "NOx": None,
    "SO2": "M_SO2.csv", "PM10": "M_PM10.csv", "PM25": "M_PM25.csv",
    "NH3": "M_NH3.csv", "CO2": "M_CO2.csv", "CH4": "M_CH4.csv", "CN": "M_CN.csv",
}

def descomponer_nox(valor_nox):
    """Descompone NOx en NO y NO2 usando ratio 9.46."""
    no2 = valor_nox / (1 + RATIO_NO_NO2)
    no = valor_nox - no2
    return no, no2

def procesar_carreteras(dir_excel, dir_csv, dir_salida, coords, dry_run=False):
    """Procesa emisiones vehiculares de carreteras."""
    print(f"\n{'='*60}")
    print(f"FUENTE: Carreteras (emisiones vehiculares)")
    print(f"Descomposición NOx: NO/NO2 = {RATIO_NO_NO2}")
    print(f"{'='*60}")
    
    excel_path = os.path.join(dir_excel, EXCEL_FILE)
    if not os.path.exists(excel_path):
        print(f"  ⚠️  {EXCEL_FILE}: No encontrado.")
        return 0
    
    df_excel = pd.read_excel(excel_path, sheet_name=HOJA)
    print(f"  📊 Excel: {len(df_excel)} filas")
    
    total_ops = 0
    
    for contaminante, proporciones in PROPORCIONES.items():
        if contaminante == "NOx":
            csv_targets = [("M_NO.csv", "NO"), ("M_NO2.csv", "NO2")]
        else:
            csv_file = CONTAMINANTE_CSV.get(contaminante)
            if csv_file is None:
                continue
            csv_targets = [(csv_file, contaminante)]
        
        col_excel = contaminante
        if col_excel not in df_excel.columns:
            print(f"  ⚠️  Columna '{col_excel}' no en Excel. Saltando.")
            continue
        
        for csv_name, comp_label in csv_targets:
            csv_path = os.path.join(dir_csv, csv_name)
            if not os.path.exists(csv_path):
                print(f"  ⚠️  {csv_name}: No encontrado. Saltando.")
                continue
            
            df_csv = pd.read_csv(csv_path, skiprows=[1])
            
            for scc, proporcion in proporciones.items():
                df_excel_prep = df_excel.copy()
                
                if contaminante == "NOx":
                    if comp_label == "NO":
                        df_excel_prep[col_excel] = df_excel_prep[col_excel].apply(
                            lambda x: descomponer_nox(x)[0]
                        )
                    else:
                        df_excel_prep[col_excel] = df_excel_prep[col_excel].apply(
                            lambda x: descomponer_nox(x)[1]
                        )
                
                df_excel_prep[col_excel] = df_excel_prep[col_excel] * proporcion * FACTOR_CONVERSION
                
                if scc not in df_csv.columns:
                    df_csv[scc] = 0.0
                
                df_csv, n_act = proteger_saltillo(
                    df_csv, df_excel_prep, scc, col_excel,
                    "grid", GRID_EXCEL, coords
                )
                
                print(f"  ✅ {csv_name} | SCC {scc} | {comp_label} ({proporcion:.4f}): {n_act} celdas")
                total_ops += 1
            
            if not dry_run:
                salida_path = os.path.join(dir_salida, csv_name)
                df_csv.to_csv(salida_path, index=False, float_format="%.6E")
                n_sccs = len([c for c in df_csv.columns if c.isdigit() and len(c) == 10])
                print(f"  💾 Guardado: {salida_path} ({n_sccs} SCCs)")
    
    return total_ops

def main():
    parser = argparse.ArgumentParser(
        description="Actualizar emisiones vehiculares de carreteras en M_*.csv"
    )
    parser.add_argument("--dry-run", action="store_true", help="Simular sin escribir")
    parser.add_argument("--dir-excel", default="datos/", help="Directorio de Excel")
    parser.add_argument("--dir-csv", default="datos/", help="Directorio de CSV")
    parser.add_argument("--dir-salida", default="salida/", help="Directorio de salida")
    args = parser.parse_args()
    
    print("=" * 60)
    print("ACTUALIZACIÓN DE EMISIONES — CARRETERAS (VEHICULARES)")
    print(f"Conversión: Mg/año → g/año (×{FACTOR_CONVERSION:.0E})")
    print(f"Descomposición NOx: NO/NO2 = {RATIO_NO_NO2}")
    print(f"Protección Saltillo: LON <= {LON_LIMITE}")
    print("=" * 60)
    
    if args.dry_run:
        print("🔍 MODO DRY-RUN: No se escribirán archivos")
    
    os.makedirs(args.dir_salida, exist_ok=True)
    
    print("\nCargando coordenadas...")
    coords = cargar_coordenadas(os.path.join(args.dir_csv, "localiza.csv"))
    
    total_ops = procesar_carreteras(
        args.dir_excel, args.dir_csv, args.dir_salida, coords, args.dry_run
    )
    
    print(f"\n{'='*60}")
    print(f"RESUMEN: {total_ops} operaciones completadas")
    print(f"{'='*60}")

if __name__ == "__main__":
    main()
