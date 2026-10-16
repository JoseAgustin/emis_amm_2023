#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_emisiones_area.py
Script maestro para actualizar emisiones de fuentes de área.

9 fuentes | 78 operaciones | 21 SCCs → 10 archivos A*_2023.csv
Conversión: Mg/año → g/año (×1E+06)
Protección Saltillo: LONGITUD <= -100.67

Uso:
  python actualizar_emisiones_area.py [opciones]

Opciones:
  --dry-run              Simular sin escribir archivos
  --fuente NOMBRE        Procesar solo una fuente específica
  --csv NOMBRE           Procesar solo un CSV específico
  --dir-excel RUTA       Directorio de archivos Excel (default: datos/)
  --dir-csv RUTA         Directorio de archivos CSV (default: datos/)
  --dir-salida RUTA      Directorio de salida (default: salida/)
"""

import os
import sys
import argparse
import pandas as pd
import numpy as np
from proteccion_saltillo import cargar_coordenadas, proteger_saltillo, LON_LIMITE

# ============================================================
# CONFIGURACIÓN: 9 fuentes de área
# ============================================================
FUENTES = {
    "FM_MaqCons_malla": {
        "descripcion": "Maquinaria de construcción",
        "excel": "FM_MaqCons_malla.xlsx",
        "hoja": "FM_MaqCons_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV",  "scc": "2311010000"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO",   "scc": "2311010000"},
            {"csv": "ANOx_2023.csv", "col_excel": "NOx",  "scc": "2311010000"},
            {"csv": "ASO2_2023.csv", "col_excel": "SO2",  "scc": "2311010000"},
            {"csv": "APM10_2023.csv","col_excel": "PM10", "scc": "2311010000"},
            {"csv": "APM25_2023.csv","col_excel": "PM2_5","scc": "2311010000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3",  "scc": "2311010000"},
            {"csv": "ACO2_2023.csv", "col_excel": "CO2",  "scc": "2311010000"},
        ]
    },
    "FA_Agro_UDS_malla3x3": {
        "descripcion": "Agrícola (7 subcategorías)",
        "excel": "FA_Agro_UDS_malla3x3.xlsx",
        "hoja": "FA_Agro_UDS_malla3x3",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Quema",  "scc": "2801500000"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO_Quema",   "scc": "2801500000"},
            {"csv": "ANOx_2023.csv", "col_excel": "NOx_Quema",  "scc": "2801500000"},
            {"csv": "APM10_2023.csv","col_excel": "PM10_Quema", "scc": "2801500000"},
            {"csv": "APM25_2023.csv","col_excel": "PM25_Quema", "scc": "2801500000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Quema",  "scc": "2801500000"},
            {"csv": "ACO2_2023.csv", "col_excel": "CO2_Quema",  "scc": "2801500000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Fert",   "scc": "2801600000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Bov",    "scc": "2805001000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Bov",    "scc": "2805001000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Porc",   "scc": "2805002000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Porc",   "scc": "2805002000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Capr",   "scc": "2805003000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Capr",   "scc": "2805003000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Aves",   "scc": "2805007000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Aves",   "scc": "2805007000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Equi",   "scc": "2805008000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Equi",   "scc": "2805008000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Solv",   "scc": "2460100000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_GasLP",  "scc": "2501060000"},
            {"csv": "APM10_2023.csv","col_excel": "PM10_Ladr",  "scc": "2302002000"},
            {"csv": "APM25_2023.csv","col_excel": "PM25_Ladr",  "scc": "2302002000"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO_Ladr",    "scc": "2302002000"},
        ]
    },
    "FA_evaporativas_9cat_malla": {
        "descripcion": "Evaporativas (9 categorías)",
        "excel": "FA_evaporativas_9cat_malla.xlsx",
        "hoja": "FA_evaporativas_9cat_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Pint",    "scc": "2401001000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_LimpSup", "scc": "2420000000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_ArtLimp", "scc": "2630020000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_LavSeco", "scc": "2420008000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_ArtPers", "scc": "2660000000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_ImpArGr", "scc": "2425000000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_PavAsf",  "scc": "2461020000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_RecubSup","scc": "2401090000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Panad",   "scc": "2302050000"},
        ]
    },
    "FA_AguasRes_RelleSant_malla": {
        "descripcion": "Aguas residuales y rellenos sanitarios",
        "excel": "FA_AguasRes_RelleSant_malla.xlsx",
        "hoja": "FA_AguasRes_RelleSant_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV_AguRes", "scc": "2630030000"},
            {"csv": "ACH4_2023.csv", "col_excel": "CH4_RelSan", "scc": "2630030000"},
        ]
    },
    "FA_Erosivas_PM_UDS_malla": {
        "descripcion": "Erosivas (PM10 y PM2.5)",
        "excel": "FA_Erosivas_PM_UDS_malla.xlsx",
        "hoja": "FA_Erosivas_PM_UDS_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "APM10_2023.csv","col_excel": "PM10_Eros", "scc": "2730100000"},
            {"csv": "APM25_2023.csv","col_excel": "PM25_Eros", "scc": "2730100000"},
        ]
    },
    "FM_Aero_CAutobus_malla": {
        "descripcion": "Aeropuerto y centrales de autobuses",
        "excel": "FM_Aero_CAutobus_malla.xlsx",
        "hoja": "FM_Aero_CAutobus_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV",  "scc": "2275050011"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO",   "scc": "2275050011"},
            {"csv": "ANOx_2023.csv", "col_excel": "NOx",  "scc": "2275050011"},
            {"csv": "ASO2_2023.csv", "col_excel": "SO2",  "scc": "2275050011"},
            {"csv": "APM10_2023.csv","col_excel": "PM10", "scc": "2275050011"},
            {"csv": "APM25_2023.csv","col_excel": "PM25", "scc": "2275050011"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3",  "scc": "2275050011"},
            {"csv": "ACO2_2023.csv", "col_excel": "CO2",  "scc": "2275050011"},
        ]
    },
    "FM_ferrocarril_emis_malla": {
        "descripcion": "Ferrocarril",
        "excel": "FM_ferrocarril_emis_malla.xlsx",
        "hoja": "FM_ferrocarril_emis_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV",  "scc": "2285000000"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO",   "scc": "2285000000"},
            {"csv": "ANOx_2023.csv", "col_excel": "NOx",  "scc": "2285000000"},
            {"csv": "ASO2_2023.csv", "col_excel": "SO2",  "scc": "2285000000"},
            {"csv": "APM10_2023.csv","col_excel": "PM10", "scc": "2285000000"},
            {"csv": "APM25_2023.csv","col_excel": "PM25", "scc": "2285000000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3",  "scc": "2285000000"},
            {"csv": "ACO2_2023.csv", "col_excel": "CO2",  "scc": "2285000000"},
        ]
    },
    "FM_maquinaria_agricola_malla": {
        "descripcion": "Maquinaria agrícola",
        "excel": "FM_maquinaria_agricola_malla.xlsx",
        "hoja": "FM_maquinaria_agricola_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV",  "scc": "2265005000"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO",   "scc": "2265005000"},
            {"csv": "ANOx_2023.csv", "col_excel": "NOx",  "scc": "2265005000"},
            {"csv": "ASO2_2023.csv", "col_excel": "SO2",  "scc": "2265005000"},
            {"csv": "APM10_2023.csv","col_excel": "PM10", "scc": "2265005000"},
            {"csv": "APM25_2023.csv","col_excel": "PM25", "scc": "2265005000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3",  "scc": "2265005000"},
            {"csv": "ACO2_2023.csv", "col_excel": "CO2",  "scc": "2265005000"},
        ]
    },
    "FA_comb_comer_regulada_malla": {
        "descripcion": "Combustión comercial regulada",
        "excel": "FA_comb_comer_regulada_malla.xlsx",
        "hoja": "FA_comb_comer_regulada_malla",
        "grid_excel": "GRIDCODE",
        "operaciones": [
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Comb",  "scc": "2103007000"},
            {"csv": "ACO_2023.csv",  "col_excel": "CO_Comb",   "scc": "2103007000"},
            {"csv": "ANOx_2023.csv", "col_excel": "NOx_Comb",  "scc": "2103007000"},
            {"csv": "ASO2_2023.csv", "col_excel": "SO2_Comb",  "scc": "2103007000"},
            {"csv": "APM10_2023.csv","col_excel": "PM10_Comb", "scc": "2103007000"},
            {"csv": "APM25_2023.csv","col_excel": "PM25_Comb", "scc": "2103007000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Comb",  "scc": "2103007000"},
            {"csv": "ACO2_2023.csv", "col_excel": "CO2_Comb",  "scc": "2103007000"},
            {"csv": "AVOC_2023.csv", "col_excel": "COV_Evap",  "scc": "2104006000"},
            {"csv": "ANH3_2023.csv", "col_excel": "NH3_Evap",  "scc": "2104006000"},
        ]
    },
}

FACTOR_CONVERSION = 1e6

def leer_csv_fortran(filepath):
    """Lee un CSV con formato Fortran (coma final, notación científica)."""
    with open(filepath, 'r') as f:
        lines = f.readlines()
    header = lines[0].strip().rstrip(',').split(',')
    scc_count_line = lines[1].strip().rstrip(',')
    data_lines = lines[2:]
    return header, scc_count_line, data_lines, lines

def escribir_csv_fortran(filepath, header, scc_count, data_df, original_lines):
    """Escribe CSV manteniendo formato Fortran."""
    with open(filepath, 'w') as f:
        f.write(','.join(header) + ',\n')
        f.write(f"{scc_count},\n")
        for idx, row in data_df.iterrows():
            values = []
            for i, col in enumerate(data_df.columns):
                val = row[col]
                if i < 4:
                    values.append(str(int(val)) if isinstance(val, (int, float)) and col != header[0] else str(val))
                else:
                    values.append(f"{val:.6E}")
            f.write(','.join(values) + ',\n')

def actualizar_contador_scc(filepath, nuevo_conteo):
    """Actualiza el contador de SCCs en la línea 2 del CSV."""
    with open(filepath, 'r') as f:
        lines = f.readlines()
    lines[1] = f"{nuevo_conteo},\n"
    with open(filepath, 'w') as f:
        f.writelines(lines)

def procesar_fuente(nombre_fuente, config, dir_excel, dir_csv, dir_salida, coords, dry_run=False):
    """Procesa una fuente de área completa."""
    print(f"\n{'='*60}")
    print(f"Fuente: {nombre_fuente} — {config['descripcion']}")
    print(f"{'='*60}")
    
    excel_path = os.path.join(dir_excel, config["excel"])
    if not os.path.exists(excel_path):
        print(f"  ⚠️  {config['excel']}: No encontrado. Saltando.")
        return 0
    
    df_excel = pd.read_excel(excel_path, sheet_name=config["hoja"])
    print(f"  📊 Excel: {len(df_excel)} filas")
    
    total_ops = 0
    csv_ops = {}
    for op in config["operaciones"]:
        csv_name = op["csv"]
        if csv_name not in csv_ops:
            csv_ops[csv_name] = []
        csv_ops[csv_name].append(op)
    
    for csv_name, ops in csv_ops.items():
        csv_path = os.path.join(dir_csv, csv_name)
        if not os.path.exists(csv_path):
            print(f"  ⚠️  {csv_name}: No encontrado. Saltando.")
            continue
        
        df_csv = pd.read_csv(csv_path, skiprows=[1])
        
        for op in ops:
            col_excel = op["col_excel"]
            scc = op["scc"]
            
            if col_excel not in df_excel.columns:
                print(f"  ⚠️  Columna '{col_excel}' no encontrada en Excel. Saltando.")
                continue
            
            scc_col = str(scc)
            if scc_col not in df_csv.columns:
                print(f"  ➕ SCC {scc} no existe en {csv_name}. Agregando columna.")
                df_csv[scc_col] = 0.0
            
            df_excel_conv = df_excel.copy()
            df_excel_conv[col_excel] = df_excel_conv[col_excel] * FACTOR_CONVERSION
            
            df_csv, n_act = proteger_saltillo(
                df_csv, df_excel_conv, scc_col, col_excel,
                "grid", config["grid_excel"], coords
            )
            
            print(f"  ✅ {csv_name} | SCC {scc} | {col_excel}: {n_act} celdas actualizadas")
            total_ops += 1
        
        if not dry_run:
            salida_path = os.path.join(dir_salida, csv_name)
            df_csv.to_csv(salida_path, index=False, float_format="%.6E")
            n_sccs = len([c for c in df_csv.columns if c.isdigit() and len(c) == 10])
            actualizar_contador_scc(salida_path, n_sccs)
            print(f"  💾 Guardado: {salida_path} ({n_sccs} SCCs)")
    
    return total_ops

def main():
    parser = argparse.ArgumentParser(
        description="Actualizar emisiones de fuentes de área en archivos A*_2023.csv"
    )
    parser.add_argument("--dry-run", action="store_true", help="Simular sin escribir")
    parser.add_argument("--fuente", type=str, help="Procesar solo una fuente")
    parser.add_argument("--csv", type=str, help="Procesar solo un CSV")
    parser.add_argument("--dir-excel", default="datos/", help="Directorio de Excel")
    parser.add_argument("--dir-csv", default="datos/", help="Directorio de CSV")
    parser.add_argument("--dir-salida", default="salida/", help="Directorio de salida")
    args = parser.parse_args()
    
    print("=" * 60)
    print("ACTUALIZACIÓN DE EMISIONES — FUENTES DE ÁREA")
    print(f"9 fuentes | 78 operaciones | 21 SCCs")
    print(f"Conversión: Mg/año → g/año (×{FACTOR_CONVERSION:.0E})")
    print(f"Protección Saltillo: LON <= {LON_LIMITE}")
    print("=" * 60)
    
    if args.dry_run:
        print("🔍 MODO DRY-RUN: No se escribirán archivos")
    
    os.makedirs(args.dir_salida, exist_ok=True)
    
    print("\nCargando coordenadas...")
    coords = cargar_coordenadas(os.path.join(args.dir_csv, "localiza.csv"))
    
    total_ops = 0
    fuentes_a_procesar = FUENTES
    
    if args.fuente:
        if args.fuente in FUENTES:
            fuentes_a_procesar = {args.fuente: FUENTES[args.fuente]}
        else:
            print(f"❌ Fuente '{args.fuente}' no encontrada.")
            print(f"   Disponibles: {', '.join(FUENTES.keys())}")
            sys.exit(1)
    
    for nombre, config in fuentes_a_procesar.items():
        ops = procesar_fuente(
            nombre, config,
            args.dir_excel, args.dir_csv, args.dir_salida,
            coords, args.dry_run
        )
        total_ops += ops
    
    print(f"\n{'='*60}")
    print(f"RESUMEN: {total_ops} operaciones completadas")
    print(f"{'='*60}")

if __name__ == "__main__":
    main()
