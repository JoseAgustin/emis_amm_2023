# emis_amm_2023

**Scripts Python para actualizar el inventario de emisiones del Área Metropolitana de Monterrey (AMM) 2023**

Distribución espacial en malla 3×3 km para el modelo de calidad del aire WRF-Chem.

## Descripción

Este repositorio contiene scripts Python que leen emisiones desde archivos Excel georreferenciados y actualizan los archivos CSV del inventario de emisiones (A*_2023.csv para fuentes de área y M_*.csv para fuentes vehiculares).

### Características principales

- **10 fuentes de emisión**: 9 de área + 1 vehicular (carreteras)
- **89+ operaciones** de actualización
- **25+ códigos SCC** (Source Classification Code)
- **Conversión automática**: Mg/año → g/año (×1E+06)
- **Protección Saltillo**: Preserva emisiones existentes en celdas con LONGITUD ≤ -100.67 cuando el Excel no tiene datos
- **Descomposición NOx**: Relación NO/NO₂ = 9.46 para emisiones vehiculares
- **Formato Fortran**: Mantiene coma final, notación científica y contador de SCCs

## Estructura del repositorio

```
emis_amm_2023/
├── scripts/
│   ├── proteccion_saltillo.py          # Módulo compartido de protección Saltillo
│   ├── actualizar_emisiones_area.py    # Script maestro: 9 fuentes de área
│   └── actualizar_emisiones_carreteras.py  # Script vehicular: carreteras
├── datos/                              # Archivos de entrada (no versionados)
│   ├── *.xlsx                          # Archivos Excel con emisiones actualizadas
│   ├── A*_2023.csv                     # CSVs de fuentes de área
│   ├── M_*.csv                         # CSVs de fuentes vehiculares
│   └── localiza.csv                    # Coordenadas de la malla
├── salida/                             # Archivos actualizados (no versionados)
├── README.md
├── requirements.txt
├── .gitignore
└── LICENSE
```

## Fuentes de emisión

### Fuentes de área (A*_2023.csv)

| # | Fuente | Excel | SCCs |
|---|--------|-------|------|
| 1 | Maquinaria de construcción | FM_MaqCons_malla.xlsx | 2311010000 |
| 2 | Agrícola (7 subcategorías) | FA_Agro_UDS_malla3x3.xlsx | 2801500000, 2801600000, 2805001000-2805008000, 2460100000, 2501060000, 2302002000 |
| 3 | Evaporativas (9 categorías) | FA_evaporativas_9cat_malla.xlsx | 2401001000, 2420000000, 2630020000, 2420008000, 2660000000, 2425000000, 2461020000, 2401090000, 2302050000 |
| 4 | Aguas residuales | FA_AguasRes_RelleSant_malla.xlsx | 2630030000 |
| 5 | Erosivas | FA_Erosivas_PM_UDS_malla.xlsx | 2730100000 |
| 6 | Aeropuerto y autobuses | FM_Aero_CAutobus_malla.xlsx | 2275050011 |
| 7 | Ferrocarril | FM_ferrocarril_emis_malla.xlsx | 2285000000 |
| 8 | Maquinaria agrícola | FM_maquinaria_agricola_malla.xlsx | 2265005000 |
| 9 | Combustión comercial | FA_comb_comer_regulada_malla.xlsx | 2103007000, 2104006000 |

### Fuente vehicular (M_*.csv)

| # | Fuente | Excel | SCCs |
|---|--------|-------|------|
| 10 | Carreteras (16 tipos de vehículo) | FM_carreteras_emis_malla.xlsx | 12 SCCs (2201-2230) |

## Protección Saltillo

La malla de 2,500 celdas (50×50) cubre tanto Monterrey como Saltillo:

- **Monterrey** (~1,750 celdas, LONG > -100.67): Siempre se actualizan desde Excel
- **Saltillo** (~750 celdas, LONG ≤ -100.67): Solo se actualizan si el Excel tiene datos válidos (>0); de lo contrario, se preservan los valores originales del CSV

Esta protección aplica tanto para fuentes de área como vehiculares.

## Uso

### Requisitos

```bash
pip install -r requirements.txt
```

Python 3.8+ con pandas ≥ 1.5.0 y openpyxl ≥ 3.0.0

### Actualizar fuentes de área

```bash
# Todas las fuentes
python scripts/actualizar_emisiones_area.py

# Solo una fuente específica
python scripts/actualizar_emisiones_area.py --fuente FM_MaqCons_malla

# Modo simulación (sin escribir archivos)
python scripts/actualizar_emisiones_area.py --dry-run

# Directorios personalizados
python scripts/actualizar_emisiones_area.py --dir-excel datos/ --dir-csv datos/ --dir-salida salida/
```

### Actualizar fuentes vehiculares

```bash
# Carreteras
python scripts/actualizar_emisiones_carreteras.py

# Modo simulación
python scripts/actualizar_emisiones_carreteras.py --dry-run
```

## Formato de archivos CSV

Los archivos CSV mantienen compatibilidad con el código Fortran del repositorio `emis_2023`:

- **Línea 1**: Encabezado con nombres de columnas + coma final
- **Línea 2**: Contador de SCCs (se actualiza automáticamente) + coma final
- **Líneas 3+**: Datos en notación científica (X.XXXXXXE+XX) + coma final

## Relación con otros repositorios

| Repositorio | Contenido | Lenguaje |
|---|---|---|
| **emis_2023** | Distribución espacial original (genera los CSV) | Fortran |
| **emis_amm_2023** | Actualización de la distribución (modifica los CSV) | Python |

## Autores

- José Agustín García Reynoso — ICAyCC, UNAM

## Licencia

MIT License
