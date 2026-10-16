#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
proteccion_saltillo.py
Módulo compartido para clasificar celdas Monterrey/Saltillo
y proteger emisiones existentes en la zona de Saltillo.

Criterio: LONGITUD <= -100.67 → Saltillo (protegida)
          LONGITUD >  -100.67 → Monterrey (actualizable)

Malla: 2,500 celdas (50×50 a 3×3 km)
  - Monterrey: ~1,750 celdas (actualizables)
  - Saltillo:  ~750 celdas (protegidas)
"""

import pandas as pd
import numpy as np

# Límite de longitud para separar Monterrey de Saltillo
LON_LIMITE = -100.67

def clasificar_zona(lon):
    """Clasifica una celda como Monterrey o Saltillo."""
    return "Saltillo" if lon <= LON_LIMITE else "Monterrey"

def cargar_coordenadas(localiza_csv="datos/localiza.csv"):
    """Carga el archivo de coordenadas y clasifica cada celda."""
    coords = pd.read_csv(localiza_csv)
    coords["zona"] = coords["LONGITUDE"].apply(clasificar_zona)
    n_mty = (coords["zona"] == "Monterrey").sum()
    n_sal = (coords["zona"] == "Saltillo").sum()
    print(f"  Celdas Monterrey: {n_mty} | Celdas Saltillo: {n_sal}")
    return coords

def proteger_saltillo(df_csv, df_excel, col_csv, col_excel, grid_csv, grid_excel, coords):
    """
    Actualiza valores en el CSV protegiendo Saltillo.
    
    - Monterrey (LONG > -100.67): siempre actualiza desde Excel
    - Saltillo (LONG <= -100.67): solo actualiza si Excel tiene dato > 0;
      si Excel no tiene dato o es 0, preserva el valor original del CSV
    
    Parámetros:
    -----------
    df_csv : DataFrame del archivo CSV
    df_excel : DataFrame del archivo Excel
    col_csv : nombre de la columna SCC en el CSV
    col_excel : nombre de la columna de emisiones en el Excel
    grid_csv : nombre de la columna grid en el CSV
    grid_excel : nombre de la columna GRIDCODE en el Excel
    coords : DataFrame con coordenadas y zona
    
    Retorna:
    --------
    df_csv actualizado, conteo de celdas actualizadas
    """
    actualizadas = 0
    
    for idx, row in df_excel.iterrows():
        gridcode = row[grid_excel]
        valor_excel = row[col_excel]
        
        # Buscar la celda en el CSV
        mask = df_csv[grid_csv] == gridcode
        if mask.sum() == 0:
            continue
        
        # Determinar zona
        coord_mask = coords["GRIDCODE"] == gridcode
        if coord_mask.sum() == 0:
            continue
        lon = coords.loc[coord_mask, "LONGITUDE"].values[0]
        zona = clasificar_zona(lon)
        
        if zona == "Monterrey":
            # Siempre actualizar
            df_csv.loc[mask, col_csv] = valor_excel
            actualizadas += 1
        else:
            # Saltillo: solo actualizar si Excel tiene dato válido
            if pd.notna(valor_excel) and valor_excel > 0:
                df_csv.loc[mask, col_csv] = valor_excel
                actualizadas += 1
            # Si no, preservar valor original (no hacer nada)
    
    return df_csv, actualizadas
