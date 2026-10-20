#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
proteccion_saltillo.py
======================
Módulo compartido para proteger las emisiones de Saltillo durante la
actualización del inventario de emisiones del AMM 2023.

Las celdas de Saltillo (LONG <= -100.67) se protegen para no perder
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
    La columna de longitud puede llamarse 'LONG' o 'LONGITUDE'.
    """
    df = pd.read_csv(ruta)

    if 'LONG' in df.columns:
        col_lon = 'LONG'
    elif 'LONGITUDE' in df.columns:
        col_lon = 'LONGITUDE'
    else:
        raise ValueError(f"No se encontro columna LONG o LONGITUDE en {ruta}")

    if 'GRIDCODE' in df.columns:
        col_grid = 'GRIDCODE'
    elif 'grid' in df.columns:
        col_grid = 'grid'
    else:
        raise ValueError(f"No se encontro columna GRIDCODE o grid en {ruta}")

    return dict(zip(df[col_grid].astype(int), df[col_lon].astype(float)))


def es_saltillo(grid, grid_lon):
    """Determina si una celda pertenece a Saltillo (LONG <= -100.67)."""
    if grid in grid_lon:
        return grid_lon[grid] <= UMBRAL_SALTILLO
    return False


def proteger_saltillo(grid, valor_nuevo, valor_actual, grid_lon):
    """
    Aplica proteccion de Saltillo a una celda.
    - Monterrey: siempre actualizar
    - Saltillo con dato en Excel: actualizar
    - Saltillo sin dato en Excel: preservar valor original
    """
    if es_saltillo(grid, grid_lon):
        if valor_nuevo == 0.0:
            return valor_actual, True
        else:
            return valor_nuevo, True
    else:
        return valor_nuevo, False
