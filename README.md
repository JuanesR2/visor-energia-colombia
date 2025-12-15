# visor-energia-colombia

Visor interactivo (Streamlit) para analizar y visualizar variables clave del sistema eléctrico colombiano: hidrología, generación, ZNI, demanda, mercado y predicciones. Desarrollado como demo para la Dirección de Energía Eléctrica (DEE).

---

## 🚀 Qué incluye

La app está organizada por secciones (menú lateral):

- **Panel principal** (tarjetas de navegación)
- **Hidrología – Embalses y aportes** (KPIs, series y análisis)
- **Fuentes de generación del SIN** (agregación por categorías y renovables)
- **Zonas No Interconectadas (ZNI)** (filtrado, agregación y descargas)
- **Mercado y variables macroeconómicas** (demanda + TRM/IPC/PIB + precio bolsa/generación)
- **Demanda (DemaSIN)** (análisis dedicado + descarga CSV)
- **Predicciones** (modelo base Ridge con tendencia y calendario; fallback a promedio móvil)

Incluye botones de descarga CSV en varias secciones (por ejemplo ZNI).

---

## 🧱 Requisitos

- Python 3.10+ recomendado
- Paquetes principales:
  - `streamlit`
  - `pandas`, `numpy`
  - `plotly`
  - `scikit-learn` (para Ridge; opcional pero recomendado)
  - `statsmodels` (opcional; si habilitas prueba ADF)
  - `pydataxm` (consumo de datos XM)

> Nota: si no instalas `scikit-learn`, el módulo de predicción puede usar un método alterno simple (promedio móvil).

---

## 📦 Archivos esperados en la carpeta del proyecto

Además de `app.py`, la app busca estos archivos **en la misma carpeta**:

### 1) Catálogos/plantillas
- `Consulta_API_SIMEM.xlsm` (catálogo Si-MEM)
- `Consulta_API_XM.xlsm` (catálogo XM)

Si no están, la app muestra una advertencia indicando que los pongas junto a `app.py`.

### 2) Logo (opcional)
- `logo.png` (si existe, se renderiza en la barra superior)

### 3) Variables macro (opcionales, en CSV)
- `TRM.csv` (diaria)
- `IPC.csv` (mensual)
- `PIB.csv` (trimestral)

Si no existen, simplemente no se cargan (quedarán vacías).

---

## ▶️ Instalación

Crea un entorno y instala dependencias:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

pip install -U pip
pip install streamlit pandas numpy plotly scikit-learn statsmodels pydataxm
