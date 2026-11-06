
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
proteccion_saltillo.py
======================
Módulo compartido para proteger las emisiones de Saltillo durante la
actualización del inventario de emisiones del AMM 2023.

Las celdas de Saltillo (LONG ≤ -100.67°) se protegen para no perder
sus emisiones originales cuando el Excel solo contiene datos de Monterrey.

Malla: ~750 celdas en Saltillo (protegidas), ~1,750 en Monterrey (actualizables)

Repositorio: emis_amm_2023
Autor: José Agustín García Reynoso — ICAyCC, UNAM
"""

import pandas as pd

# Umbral de longitud para Saltillo
UMBRAL_SALTILLO = -100.67


def cargar_localiza(ruta):
    """
    Carga localiza.csv y retorna diccionario {GRIDCODE: LONGITUD}.

    El archivo localiza.csv contiene las coordenadas de cada celda de la malla.
    La columna de longitud puede llamarse 'LONG' o 'LONGITUDE'.

    Args:
        ruta: Ruta al archivo localiza.csv

    Returns:
        dict: {gridcode (int): longitud (float)}
    """
    df = pd.read_csv(ruta)

    # Detectar nombre de columna de longitud
    if 'LONG' in df.columns:
        col_lon = 'LONG'
    elif 'LONGITUDE' in df.columns:
        col_lon = 'LONGITUDE'
    else:
        raise ValueError(f"No se encontró columna LONG o LONGITUDE en {ruta}")

    # Detectar nombre de columna de grid
    if 'GRIDCODE' in df.columns:
        col_grid = 'GRIDCODE'
    elif 'grid' in df.columns:
        col_grid = 'grid'
    else:
        raise ValueError(f"No se encontró columna GRIDCODE o grid en {ruta}")

    return dict(zip(df[col_grid].astype(int), df[col_lon].astype(float)))


def es_saltillo(grid, grid_lon):
    """
    Determina si una celda pertenece a Saltillo.

    Args:
        grid: GRIDCODE de la celda
        grid_lon: diccionario {GRIDCODE: longitud}

    Returns:
        bool: True si la celda está en Saltillo (LONG ≤ -100.67)
    """
    if grid in grid_lon:
        return grid_lon[grid] <= UMBRAL_SALTILLO
    return False


def proteger_saltillo(grid, valor_nuevo, valor_actual, grid_lon):
    """
    Aplica protección ESTRICTA de Saltillo a una celda (Opción A).

    Lógica:
    - Si la celda ES de Saltillo (LONG ≤ -100.67) → SIEMPRE preservar el
      valor actual del CSV, sin importar valor_nuevo. La celda nunca se
      sobreescribe, ni siquiera cuando el Excel trae datos.
    - Si la celda NO es de Saltillo (Monterrey) → usar valor nuevo (actualizar).

    Args:
        grid: GRIDCODE de la celda
        valor_nuevo: valor calculado desde el Excel (en g/año)
        valor_actual: valor actual en el CSV (en g/año)
        grid_lon: diccionario {GRIDCODE: longitud}

    Returns:
        tuple: (valor_final, fue_protegida)
            - valor_final: el valor a escribir en el CSV
            - fue_protegida: True si se aplicó protección de Saltillo
    """
    if es_saltillo(grid, grid_lon):
        # Saltillo → NUNCA sobreescribir: preservar el valor original del CSV.
        return valor_actual, True
    else:
        # Monterrey → siempre actualizar
        return valor_nuevo, False

