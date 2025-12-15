# visor-energia-colombia

Visor interactivo (Streamlit) para analizar y visualizar variables clave del sistema eléctrico colombiano: hidrología, generación, ZNI, demanda, mercado y predicciones. Desarrollado como **demo** para la **Dirección de Energía Eléctrica (DEE)**.  

> Título en la app: “Visor del sistema eléctrico — Hidrología, generación, mercado y predicciones del SIN” :contentReference[oaicite:1]{index=1}

---

## 🚀 Qué incluye

La app está organizada por secciones (menú lateral):  
- **Panel principal** (tarjetas de navegación) :contentReference[oaicite:2]{index=2}  
- **Hidrología – Embalses y aportes** (KPIs, series y análisis) :contentReference[oaicite:3]{index=3}  
- **Fuentes de generación del SIN** (agregación por categorías y renovables) :contentReference[oaicite:4]{index=4}  
- **Zonas No Interconectadas (ZNI)** (filtrado, agregación y descargas) :contentReference[oaicite:5]{index=5}  
- **Mercado y variables macroeconómicas** (demanda + TRM/IPC/PIB + precio bolsa/generación) :contentReference[oaicite:6]{index=6}  
- **Demanda (DemaSIN)** (análisis dedicado + descarga CSV) :contentReference[oaicite:7]{index=7}  
- **Predicciones** (modelo base Ridge con tendencia y calendario; fallback a promedio móvil) :contentReference[oaicite:8]{index=8} :contentReference[oaicite:9]{index=9}  

Incluye botones de descarga CSV en varias secciones (por ejemplo ZNI). :contentReference[oaicite:10]{index=10}

---

## 🧱 Requisitos

- Python 3.10+ recomendado
- Paquetes principales:
  - `streamlit`
  - `pandas`, `numpy`
  - `plotly`
  - `scikit-learn` (para Ridge; opcional pero recomendado) :contentReference[oaicite:11]{index=11}
  - `statsmodels` (opcional; habilita prueba ADF si la usas)
  - `pydataxm` (consumo de datos XM)
  
> Nota: si no instalas `scikit-learn`, el módulo de predicción usa un método alterno simple (promedio móvil). :contentReference[oaicite:12]{index=12}

---

## 📦 Archivos esperados en la carpeta del proyecto

Además de `app.py`, la app busca estos archivos **en la misma carpeta**:

### 1) Catálogos/plantillas
- `Consulta_API_SIMEM.xlsm` (catálogo Si-MEM) :contentReference[oaicite:13]{index=13}  
- `Consulta_API_XM.xlsm` (catálogo XM) :contentReference[oaicite:14]{index=14}  

Si no están, la app muestra una advertencia indicando que los pongas junto a `app.py`. :contentReference[oaicite:15]{index=15}

### 2) Logo
- `logo.png` (si existe, se renderiza en la barra superior) :contentReference[oaicite:16]{index=16}  

### 3) Variables macro (opcionales, en CSV)
- `TRM.csv` (diaria) :contentReference[oaicite:17]{index=17}  
- `IPC.csv` (mensual) :contentReference[oaicite:18]{index=18}  
- `PIB.csv` (trimestral) :contentReference[oaicite:19]{index=19}  

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
