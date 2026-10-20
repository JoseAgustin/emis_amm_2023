# emis_amm_2023

## Sistema de Actualización del Inventario de Emisiones — Área Metropolitana de Monterrey 2023

Scripts en Python para actualizar la distribución espacial de emisiones atmosféricas
del inventario 2023 del Área Metropolitana de Monterrey (AMM), empleando datos
georreferenciados en malla de 3×3 km para el modelo WRF-Chem.

**Institución:** Instituto de Ciencias de la Atmósfera y Cambio Climático (ICAyCC), UNAM  
**Autor:** José Agustín García Reynoso  
**Versión:** 1.2.0

---

## Descripción General

El sistema lee archivos Excel con emisiones actualizadas por fuente y actualiza los
archivos CSV del inventario que alimentan al modelo de calidad del aire WRF-Chem.
Se procesan **10 fuentes de emisión** (9 de área + 1 vehicular) con **80 operaciones**
de actualización y **24 códigos SCC** de área + **16 SCCs** vehiculares.

### Tipos de archivos CSV

| Tipo | Prefijo | Formato | Ejemplo |
|------|---------|---------|---------|
| Área | `A*_2023.csv` | `grid,CID,Furb,Frural,SCCs` | `AVOC_2023.csv` |
| Vehicular | `M_*.csv` | `GRIDCODE, emissions in gram per year` | `M_VOC.csv` |

### Convención de nombres CSV (compatibilidad Fortran)

Los nombres de archivos CSV mantienen longitud fija para compatibilidad con el código Fortran:

| Contaminante | Archivo Área | Archivo Vehicular |
|---|---|---|
| COV | `AVOC_2023.csv` | `M_VOC.csv` |
| CO | `ACO__2023.csv` | `M_CO.csv` |
| NOx | `ANOx_2023.csv` | — |
| NO | — | `M_NO.csv` |
| NO₂ | — | `M_NO2.csv` |
| SO₂ | `ASO2_2023.csv` | `M_SO2.csv` |
| PM10 | `APM10_2023.csv` | `M_PM10.csv` |
| PM2.5 | `APM25_2023.csv` | `M_PM25.csv` |
| NH₃ | `ANH3_2023.csv` | `M_NH3.csv` |
| CN | `ACN__2023.csv`  | `M_CN.csv` |
| CO₂ | `ACO2_2023.csv` | — |
| CH₄ | `ACH4_2023.csv` | — |

> **Nota:** `ACO__2023.csv` y `ACN__2023.csv` usan doble guion bajo para mantener la misma longitud que los demás archivos.

---

## Estructura del Repositorio

```
emis_amm_2023/
├── README.md
├── .gitignore
├── requirements.txt
├── LICENSE
├── scripts/
│   ├── proteccion_saltillo.py
│   ├── actualizar_emisiones_area.py
│   └── actualizar_emisiones_carreteras.py
├── datos/
│   ├── localiza.csv
│   ├── FM_MaqCons_malla.xlsx
│   ├── FA_Agro_UDS_malla3x3.xlsx
│   ├── FA_evaporativas_9cat_malla.xlsx
│   ├── FA_AguasRes_RelleSant_malla.xlsx
│   ├── FA_Erosivas_PM_UDS_malla.xlsx
│   ├── FM_Aero_CAutobus_malla.xlsx
│   ├── FM_ferrocarril_emis_malla.xlsx
│   ├── FM_maquinaria_agricola_malla.xlsx
│   ├── FA_comb_comer_regulada_malla.xlsx
│   └── FM_carreteras_emis_malla.xlsx
├── csv_originales/
│   ├── AVOC_2023.csv ... ACH4_2023.csv
│   └── M_VOC.csv ... M_CN.csv
└── csv_actualizados/
```

---

## Fuentes de Emisión

### Fuentes de Área (9 fuentes → A*_2023.csv)

| # | Fuente | Excel | Ops | SCCs |
|---|--------|-------|-----|------|
| 1 | Maquinaria de construcción | FM_MaqCons_malla.xlsx | 6 | 2311010000 |
| 2 | Agricultura y ganadería | FA_Agro_UDS_malla3x3.xlsx | 23 | 8 SCCs |
| 3 | Evaporativas (9 cat.) | FA_evaporativas_9cat_malla.xlsx | 9 | 9 SCCs |
| 4 | Aguas residuales y relleno sanitario | FA_AguasRes_RelleSant_malla.xlsx | 2 | 2630030000 |
| 5 | Erosión de suelos | FA_Erosivas_PM_UDS_malla.xlsx | 2 | 2730100000 |
| 6 | Aeropuerto y C. Autobuses | FM_Aero_CAutobus_malla.xlsx | 6 | 2270008000 |
| 7 | Ferrocarril | FM_ferrocarril_emis_malla.xlsx | 6 | 2285000000 |
| 8 | Maquinaria agrícola | FM_maquinaria_agricola_malla.xlsx | 8 | 2265005000 |
| 9 | Combustión comercial | FA_comb_comer_regulada_malla.xlsx | 9 | 2103007000, 2103004000 |

#### Detalle Fuente 2: Agricultura y ganadería (7 subcategorías)

| Subcategoría | Columnas Excel | SCC | Contaminantes |
|---|---|---|---|
| F1.1 Combustión agrícola | SUM_F1_1_*_AW | 2270005000 | PM10, PM2.5, SO2, CO, NOx, COV, NH3 |
| F3.1 Distribución GLP | SUM_F3_1_COV_AW | 3333333333 | COV |
| F5.1 Fertilizantes | SUM_F5_1_NH3_AW | 2801700000 | NH3 |
| F5.2 Plaguicidas | SUM_F5_2_COV_AW | 2461850000 | COV |
| F5.3 Ganadería | SUM_F5_3_*_AW | 2805020000 | NH3, CH4 |
| F5.4 Corrales de engorda | SUM_F5_4_*_AW | 2805000000 | PM10, PM2.5 |
| F5.5 Labranza | SUM_F5_5_*_AW | 2801000002 | PM10, PM2.5 |
| F5.6 Quemas agrícolas | SUM_F5_6_*_AW | 2801500100 | PM10, PM2.5, SO2, CO, NOx, COV, NH3 |

#### Detalle Fuente 3: Evaporativas (9 categorías)

| Categoría | Columna Excel | SCC |
|---|---|---|
| Artes gráficas | F21ArtGraf | 2425000000 |
| Lavado en seco | F23LavSeco | 2420000055 |
| Limpieza superficies industriales | F24LimSupInd | 2415000000 |
| Pintura automotriz | F25PintAuto | 2401005000 |
| Recubrimiento superficies industriales | F28RecSupInd | 2401002000 |
| Distribución GLP comercial | F31GLPCom | 3333333333 |
| Distribución GLP residencial | F31GLPCom | 2390007000 |
| Manejo y distribución gasolina | F32ManDisGasolina | 2501060000 |
| Panificación | F43Panifica | 2302050000 |

#### Detalle Fuente 8: Maquinaria agrícola (8 contaminantes)

| Columna Excel | SCC | CSV |
|---|---|---|
| SUM_PM10 | 2265005000 | APM10_2023.csv |
| SUM_PM2_5 | 2265005000 | APM25_2023.csv |
| SUM_SO2 | 2265005000 | ASO2_2023.csv |
| SUM_CO | 2265005000 | ACO__2023.csv |
| SUM_NOX | 2265005000 | ANOx_2023.csv |
| SUM_COV | 2265005000 | AVOC_2023.csv |
| SUM_NH3 | 2265005000 | ANH3_2023.csv |
| SUM_CN | 2265005000 | ACN__2023.csv |

#### Detalle Fuente 9: Combustión comercial (2 SCCs)

| Tipo | Columna Excel | SCC | CSV |
|---|---|---|---|
| Combustión PM10 | PM10 | 2103007000 | APM10_2023.csv |
| Combustión PM2.5 | PM2_5 | 2103007000 | APM25_2023.csv |
| Combustión SO2 | SO2 | 2103007000 | ASO2_2023.csv |
| Combustión CO | CO | 2103007000 | ACO__2023.csv |
| Combustión NOx | NOx | 2103007000 | ANOx_2023.csv |
| Combustión COV | COTComb | 2103007000 | AVOC_2023.csv |
| Combustión NH3 | NH3 | 2103007000 | ANH3_2023.csv |
| Combustión CN | CN | 2103007000 | ACN__2023.csv |
| Proceso COV | COVProc | 2103004000 | AVOC_2023.csv |

### Fuente Vehicular (1 fuente → M_*.csv)

| # | Fuente | Excel | Contaminantes | SCCs |
|---|--------|-------|---------------|------|
| 10 | Carreteras | FM_carreteras_emis_malla.xlsx | 7 (+NO/NO2) | 16 |

#### Contaminantes vehiculares

| Contaminante | Columna Excel | CSV |
|---|---|---|
| COV | SUM_COV_LW | M_VOC.csv |
| CO | SUM_CO_LW | M_CO.csv |
| SO₂ | SUM_SO2_LW | M_SO2.csv |
| PM10 | SUM_PM10_LW | M_PM10.csv |
| PM2.5 | SUM_PM25_LW | M_PM25.csv |
| NH₃ | SUM_NH3_LW | M_NH3.csv |
| CN | SUM_CN_LW | M_CN.csv |
| NOx → NO | SUM_NOx_LW × 0.90440 | M_NO.csv |
| NOx → NO₂ | SUM_NOx_LW × 0.09560 | M_NO2.csv |

#### 16 SCCs vehiculares

| Tipo | SCCs |
|---|---|
| Escape (330) | 2201001330, 2201020330, 2201040330, 2201070330, 2201080330, 2230001330, 2230060330, 2230071330, 2230073330, 2230074330, 2230075330 |
| Evaporativos (33V) | 220100133V, 220102033V, 220104033V, 220107033V, 220108033V |

> **Nota:** Los SCCs evaporativos (33V) solo tienen proporción > 0 para COV.

#### Descomposición NOx → NO + NO₂

```
Relación NO/NO₂ = 9.46
Fracción NO  = 9.46 / (1 + 9.46) = 0.90440
Fracción NO₂ = 1.00 / (1 + 9.46) = 0.09560
```

---

## Protección de Saltillo

Los archivos Excel contienen datos actualizados **solo para Monterrey**. Las celdas de
Saltillo (LONGITUD ≤ -100.67°) se protegen para preservar sus emisiones originales.

Esta protección aplica tanto a **fuentes de área** (A*_2023.csv) como a **fuentes vehiculares** (M_*.csv).

| Zona | Celdas | Criterio | Comportamiento |
|------|--------|----------|----------------|
| Monterrey | ~1,750 | LONG > -100.67 | Siempre actualizar |
| Saltillo | ~750 | LONG ≤ -100.67 | Preservar si Excel no tiene dato |

---

## Conversión de Unidades

```
Excel: Mg/año (megagramos)  →  CSV: g/año (gramos)
Factor: 1 Mg = 1,000,000 g = 1×10⁶ g
```

---

## Formato de Archivos CSV

### Archivos de Área (A*_2023.csv)

```
R1: grid,CID,Furb,Frural,SCCs,,,,
R2: N,g_per_year,SCC_1,SCC_2,...,SCC_N,
R3+: grid, CID, Furb, Frural, emis_1, ..., emis_N,
```

- 4 columnas fijas + N emisiones
- Coma final en cada línea (Fortran)
- Notación científica: X.XXXXXE+XX
- R2: contador de SCCs se actualiza automáticamente

### Archivos Vehiculares (M_*.csv)

```
R1: GRIDCODE, emissions in gram per year
R2:      16, SCC_1, SCC_2, ..., SCC_16
R3+: GRIDCODE, emis_1, ..., emis_16, 6
```

- 1 columna fija (GRIDCODE) + N emisiones + campo final (6 = huso UTM)
- Notación científica: X.XXXXE+XX

---

## Instalación y Uso

### Requisitos

```bash
pip install -r requirements.txt
```

### Ejecución

```bash
# Emisiones de área (9 fuentes)
python scripts/actualizar_emisiones_area.py \
    --dir-excel datos --dir-csv csv_originales \
    --dir-output csv_actualizados --dir-localiza datos

# Emisiones vehiculares (carreteras)
python scripts/actualizar_emisiones_carreteras.py \
    --dir-excel datos --dir-csv csv_originales \
    --dir-output csv_actualizados --dir-localiza datos
```

### Opciones

| Opción | Descripción | Default |
|--------|-------------|---------|
| `--dir-excel` | Directorio con archivos Excel | `datos` |
| `--dir-csv` | Directorio con CSVs originales | `csv_originales` |
| `--dir-output` | Directorio de salida | `csv_actualizados` |
| `--dir-localiza` | Directorio con localiza.csv | `datos` |
| `--fuente` | Solo esta fuente (área) | Todas |
| `--csv` | Solo este CSV (área) | Todos |
| `--dry-run` | Simular sin escribir | No |

---

## Malla de Cálculo

- **Resolución**: 3 × 3 km
- **Celdas totales**: 2,500 (50 × 50)
- **Cobertura**: Monterrey (~1,750) + Saltillo (~750)
- **Coordenadas**: localiza.csv (GRIDCODE, LONG, LAT)

---

## Licencia

MIT License — ICAyCC, UNAM
