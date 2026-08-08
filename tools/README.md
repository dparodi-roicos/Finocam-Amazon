# Dashboard Semanal Amazon — Builder

Script Python que genera el HTML del dashboard semanal de sell-out (MerchantSpring).

## Cómo replicar para otro cliente

**Solo hay que editar el bloque `── CONFIG ──` al inicio de `build_weekly.py`:**

| Parámetro | Qué es | Ejemplo Finocam |
|---|---|---|
| `CLIENT_NAME` | Nombre en el header | `"Finocam"` |
| `CLIENT_LETTER` | Letra del logo | `"F"` |
| `MARKETPLACE_LABEL` | Código mercado | `"ES"` |
| `MARKETPLACE_FLAG` | Emoji bandera | `"🇪🇸"` |
| `CSV_DIR` | Carpeta con los 4 CSVs semanales | ver abajo |
| `EXCEL_PATH` | Excel catálogo ASINs | `FINOCAM_Familias_Subfamilias_ASIN.xlsx` |
| `OUT_PATH` | HTML de salida | `FINOCAM_Weekly.html` |
| `WEEKS` | Etiquetas de las 4 semanas | `[{'label':'7–13 Jul','wk':'W1'}, ...]` |
| `COLS` | Hojas "colección" cross-familia | `{'Moniquilla','Talkual'}` (o `set()` si no hay) |
| `SKIP` | Hojas a ignorar | `{'Resumen'}` |
| `CAT_ORDER` | Orden de categorías en tabla | lista de nombres de familia |
| `UPDATE_DATE` | Fecha visible en toolbar | `"2026-08-07"` |

## Preparar los CSVs

1. En MerchantSpring → Reports → `generateOrderedRevenueReport` para cada semana
2. Descargar el CSV de cada semana
3. Renombrar como `weekly_0.csv` (W1) … `weekly_3.csv` (W4)
4. Guardar en la carpeta que apunta `CSV_DIR`

Columnas que usa el script: `asin`, `orderedRevenue`, `orderedUnits`, `priorOrderedRevenue`, `priorOrderedUnits`, `title`

## Estructura del Excel catálogo

El script espera estas columnas según el tipo de hoja:

**Hojas normales** (una hoja = una familia):
- A = Subfamilia | B = ASIN | C = Anualidad | D = Descripción

**Hojas colección** (listadas en `COLS`):
- A = Familia | B = Subfamilia | C = ASIN | D = Anualidad | E = Descripción

Si el cliente no tiene Excel catálogo, hay que adaptar la sección `── 1. Catálogo ──` del script para cargar los ASINs de otra fuente.

## Ejecutar

```
python tools/build_weekly.py
```

Requiere: `openpyxl` (pip install openpyxl)
