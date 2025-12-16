from __future__ import annotations

import base64
import datetime as dt
from pathlib import Path
from typing import List, Optional, Dict, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import unicodedata
from plotly.subplots import make_subplots

from pydataxm.pydataxm import ReadDB
from pydataxm.pydatasimem import ReadSIMEM

import numpy as np
import numpy as np
from sklearn.linear_model import Ridge

try:
    # Para test ADF de estacionariedad (opcional)
    from statsmodels.tsa.stattools import adfuller
except ImportError:
    adfuller = None


try:
    from sklearn.ensemble import RandomForestRegressor
    from sklearn.metrics import mean_absolute_error
except ImportError:
    RandomForestRegressor = None
    mean_absolute_error = None
import numpy as np

try:
    from sklearn.linear_model import Ridge
    from sklearn.preprocessing import StandardScaler
    from sklearn.pipeline import Pipeline
    SKLEARN_AVAILABLE = True
except Exception:
    SKLEARN_AVAILABLE = False

# -------------------------------------------------------------------
# RUTAS Y CONFIGURACIÓN
# -------------------------------------------------------------------
SIMEM_EXCEL = Path("Consulta_API_SIMEM.xlsm")
XM_EXCEL = Path("Consulta_API_XM.xlsm")
LOGO_PATH = Path("logo.png")

CACHE_DIR = Path("xm_cache")
CACHE_DIR.mkdir(exist_ok=True)

# Fecha mínima para buscar "históricos" en embalses
# (se reduce a 2010 para evitar errores 400 de XM en años viejos)
GLOBAL_EARLIEST_DATE = dt.date(2010, 1, 1)

# Paleta pastel para gráficos (bien en tema oscuro)
PRIMARY_PURPLE = "#8F339F"
PASTEL_BLUE = "#2B618D"
PASTEL_GREEN = "#2D9230"
PASTEL_RED = "#ED1B1B"
PASTEL_YELLOW = "#EEBA20"
# Número de armónicos de Fourier para la estacionalidad anual.
# Si quieres pronósticos más suaves, baja a 3; más “picos”, sube a 6–7.
FOURIER_ORDER_YEARLY = 5


# -------------------------------------------------------------------
# ESTILO GLOBAL
# -------------------------------------------------------------------
# Código API para Precio de Bolsa de Energía (diario)
# ⚠️ AJUSTA ESTE STRING según aparezca en tu Consulta_API_XM.xlsm
# Ejemplos típicos: "PrecBolsaEner", "PrecBolsaEnerDia", etc.
CODIGO_API_PRECIO_BOLSA = "PrecBolsNaci"
def set_global_style() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 0.15rem;
            padding-bottom: 0.4rem;
            padding-left: 1rem;
            padding-right: 1rem;
            max-width: 1300px;
        }

        header[data-testid="stHeader"] {
            background: transparent;
        }

        footer {
            visibility: hidden;
        }

        /* ---------- TOP BAR ---------- */

        .top-bar {
            background: #111111;
            border-bottom: 2px solid #EEBA20;
            padding: 0.3rem 0.75rem;
            margin: 0 -0.75rem 0.4rem -0.75rem;
        }

        .top-bar-inner {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.5rem;
        }

        .top-bar-left {
            display: flex;
            flex-direction: column;
            gap: 0.1rem;
        }

        .top-bar-title {
            font-size: 1.05rem;
            font-weight: 700;
            color: #f7f7f7;
        }

        .top-bar-subtitle {
            font-size: 0.85rem;
            color: #f0e7c0;
        }

        .top-bar-author {
            font-size: 0.75rem;
            color: #c9c4a2;
        }

        .top-bar-logo {
            max-height: 32px;
        }

        .top-bar-right {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        /* “Píldora” con el nombre de la sección actual  */
        .top-bar-section {
            font-size: 0.9rem;
            padding: 0.15rem 0.7rem;
            border-radius: 999px;
            border: 1px solid #EEBA20;
            color: #f7f7f7;
            white-space: nowrap;
        }

        /* ---------- KPIs SIN BARRA NEGRA ---------- */

        .kpi-card {
            background: transparent;   /* antes: #181818 */
            border: none;              /* sin bordes ni sombra */
            padding: 0;
            margin-top: 0;
            box-shadow: none;
        }

        .kpi-card .stMetric {
            padding: 0;
        }

        .kpi-card div[data-testid="stMetricLabel"] {
            font-size: 0.8rem;
        }

        .kpi-card div[data-testid="stMetricValue"] {
            font-size: 1.05rem;
        }

        .kpi-card div[data-testid="stMetricDelta"] {
            font-size: 0.7rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_top_header() -> None:
    """Barra superior con título, subtítulo y logo a la derecha."""
    logo_html = ""
    if LOGO_PATH.exists():
        with open(LOGO_PATH, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{b64}" class="top-bar-logo"/>'

    st.markdown(
        f"""
        <div class="top-bar">
          <div class="top-bar-inner">
            <div class="top-bar-left">
              <div class="top-bar-title">Visor del sistema eléctrico</div>
              <div class="top-bar-subtitle">
                Hidrología, generación, mercado y predicciones del SIN
              </div>
              <div class="top-bar-author">
                Ing. Juan E.R. Villada – Para la Dirección de Energía Eléctrica (DEE)
              </div>
            </div>
            <div class="top-bar-right">
              {logo_html}
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -------------------------------------------------------------------
# LOGO CENTRADO
# -------------------------------------------------------------------
def add_logo_center(logo_path: Path = LOGO_PATH, height: int = 64) -> None:
    if not logo_path.exists():
        return
    with open(logo_path, "rb") as f:
        img_bytes = f.read()
    b64 = base64.b64encode(img_bytes).decode()
    st.markdown(
        f"""
        <div style="text-align:center; margin-top:5px; margin-bottom:5px;">
            <img src="data:image/png;base64,{b64}" height="{height}">
        </div>
        """,
        unsafe_allow_html=True,
    )


# -------------------------------------------------------------------
# UTILIDADES CATÁLOGOS SI-MEM / XM
# -------------------------------------------------------------------
@st.cache_data
def load_simem_catalog() -> pd.DataFrame:
    if not SIMEM_EXCEL.exists():
        st.warning(
            f"No se encontró {SIMEM_EXCEL}.\n"
            "Pon Consulta_API_SIMEM.xlsm en la misma carpeta que app.py."
        )
        return pd.DataFrame()
    df = pd.read_excel(SIMEM_EXCEL, sheet_name="ListadoVariables")
    df.columns = [c.strip() for c in df.columns]
    return df


@st.cache_data
def load_xm_catalog() -> pd.DataFrame:
    if not XM_EXCEL.exists():
        st.warning(
            f"No se encontró {XM_EXCEL}.\n"
            "Pon Consulta_API_XM.xlsm en la misma carpeta que app.py."
        )
        return pd.DataFrame()
    df = pd.read_excel(XM_EXCEL, sheet_name="Parametros")
    df.columns = [c.strip() for c in df.columns]
    return df


@st.cache_data
def get_dataset_unit(dataset_id: str) -> Optional[str]:
    catalog = load_simem_catalog()
    if catalog.empty or "IdDataset" not in catalog.columns:
        return None
    mask = catalog["IdDataset"].astype(str) == str(dataset_id)
    if not mask.any():
        return None
    row = catalog[mask].iloc[0]
    for col in ["Unidad", "UnidadMedida", "Unidad Variable", "UnidadVariable", "Unidad Medida"]:
        if col in row.index and pd.notna(row[col]):
            return str(row[col])
    return None


def parse_unit_info(unit_str: Optional[str]) -> Dict[str, Optional[str]]:
    if not unit_str or not isinstance(unit_str, str):
        return {"kind": "unknown", "base_unit": None}
    s = unit_str.lower().replace(" ", "")
    # Energía
    if "gwh" in s:
        return {"kind": "energy", "base_unit": "GWh"}
    if "mwh" in s:
        return {"kind": "energy", "base_unit": "MWh"}
    if "kwh" in s:
        return {"kind": "energy", "base_unit": "kWh"}
    # Potencia
    if "gw" in s and "gwh" not in s:
        return {"kind": "power", "base_unit": "GW"}
    if "mw" in s and "mwh" not in s:
        return {"kind": "power", "base_unit": "MW"}
    if "kw" in s and "kwh" not in s:
        return {"kind": "power", "base_unit": "kW"}
    return {"kind": "unknown", "base_unit": None}


def convert_series_numeric(
    series: pd.Series,
    base_unit: Optional[str],
    target_unit: Optional[str],
    kind: str,
) -> pd.Series:
    num = pd.to_numeric(series, errors="coerce")
    if kind not in ("power", "energy") or not base_unit or not target_unit:
        return num
    if kind == "power":
        scale = {"kW": 1e3, "MW": 1e6, "GW": 1e9}
    else:
        scale = {"kWh": 1e3, "MWh": 1e6, "GWh": 1e9}
    if base_unit not in scale or target_unit not in scale:
        return num
    factor = scale[base_unit] / scale[target_unit]
    return num * factor

def detect_zni_mask(df: pd.DataFrame) -> pd.Series:
    """
    Intenta identificar qué filas corresponden a Zonas No Interconectadas (ZNI)
    usando heurísticas sobre las columnas del dataset enriquecido de generación.

    Estrategia:
      1) Si existe una columna tipo 'Sistema' / 'SistemaElectrico', filtra
         filas donde el texto contenga 'ZNI'.
      2) Si no, busca columnas cuyo nombre contenga 'zni' y las interpreta como
         bandera (bool) o indicador (>0).

    Devuelve:
      pd.Series(bool) con el mismo índice que df.
    """
    if df is None or df.empty:
        return pd.Series(False, index=df.index)

    cols = list(df.columns)
    mask = pd.Series(False, index=df.index)

    # 1) Columnas tipo 'Sistema'
    col_sistema = None
    for c in cols:
        cl = c.lower()
        if "sistema" in cl:
            col_sistema = c
            break

    if col_sistema is not None:
        txt = df[col_sistema].astype(str).str.upper()
        m = txt.str.contains("ZNI")
        if m.any():
            return m.fillna(False)

    # 2) Columnas cuyo nombre contenga 'zni'
    for c in cols:
        cl = c.lower()
        if "zni" in cl:
            serie = df[c]
            # Si es booleana: True => ZNI
            if pd.api.types.is_bool_dtype(serie):
                m2 = serie.fillna(False)
            else:
                # numérico o texto, interpretamos >0 como ZNI
                s_num = pd.to_numeric(serie, errors="coerce")
                m2 = (s_num > 0).fillna(False)

            if m2.any():
                return m2

    # Si nada funcionó, devolvemos todo False
    return mask

# -------------------------------------------------------------------
# CATÁLOGO DE RECURSOS XM (plantas)
# -------------------------------------------------------------------
@st.cache_data
def load_recursos_xm() -> pd.DataFrame:
    api = ReadDB()
    dummy_date = "2020-01-01"
    df = api.request_data(
        coleccion="ListadoRecursos",
        metrica="Sistema",
        start_date=dummy_date,
        end_date=dummy_date,
        filtros=None,
    )
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = [c.strip() for c in df.columns]
    return df


# -------------------------------------------------------------------
# UTILIDADES GENERALES
# -------------------------------------------------------------------

import numpy as np  # asegúrate de tener este import al inicio del archivo


import numpy as np  # asegúrate de tenerlo en los imports de arriba


def build_time_features(index: pd.DatetimeIndex, t_start: int = 0) -> pd.DataFrame:
    """
    Crea features temporales para una serie diaria:
      - t, t^2 (tendencia)
      - sin / cos anual (estacionalidad)
      - dummies de día de la semana y mes

    t_start permite continuar la numeración de t entre histórico y pronóstico.
    """
    idx = pd.to_datetime(index)
    n = len(idx)

    # t continua desde t_start: [t_start, t_start+1, ..., t_start+n-1]
    t = np.arange(t_start, t_start + n, dtype=float)

    df_feat = pd.DataFrame(
        {
            "t": t,
            "t2": t ** 2,
            "sin_year": np.sin(2 * np.pi * t / 365.25),
            "cos_year": np.cos(2 * np.pi * t / 365.25),
            "dow": idx.dayofweek,  # 0–6
            "month": idx.month,    # 1–12
        },
        index=idx,
    )

    # One-hot de día de semana y mes
    df_feat = pd.get_dummies(df_feat, columns=["dow", "month"], drop_first=True)

    return df_feat

def numeric_columns(df: pd.DataFrame) -> List[str]:
    return df.select_dtypes(include=["number"]).columns.tolist()


def _sanitize_max_dias(raw_val) -> int:
    try:
        val = int(raw_val)
    except Exception:
        val = 365
    if val <= 0:
        val = 365
    return val


def _detectar_col_fecha(df: pd.DataFrame) -> str:
    for c in df.columns:
        cl = c.lower()
        if "fecha" in cl or "date" in cl:
            return c
    return df.columns[0]


def _detectar_col_valor(df: pd.DataFrame) -> Optional[str]:
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if not num_cols:
        return None
    for c in num_cols:
        cl = c.lower()
        if not any(k in cl for k in ["anio", "año", "mes", "dia", "día"]):
            return c
    return num_cols[0]


def _detectar_col_embalse(df: pd.DataFrame) -> str:
    for c in df.columns:
        cl = c.lower()
        if "embals" in cl or "embalse" in cl:
            return c
    for c in df.columns:
        if "nombre" in c.lower():
            return c
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]


def _detectar_col_rio(df: pd.DataFrame) -> str:
    for c in df.columns:
        cl = c.lower()
        if "rio" in cl or "río" in cl or "cuenca" in cl:
            return c
    for c in df.columns:
        if "nombre" in c.lower():
            return c
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]


# -------------------------------------------------------------------
# XM – METADATOS (código API → metrica, max_dias)
# -------------------------------------------------------------------
@st.cache_data
def get_xm_meta(codigo_api: str) -> Dict[str, object]:
    catalog = load_xm_catalog()
    if catalog.empty:
        return {"metrica": "Sistema", "max_dias": 365}

    col_cod = "Código API" if "Código API" in catalog.columns else "Codigo API"
    if col_cod not in catalog.columns:
        return {"metrica": "Sistema", "max_dias": 365}

    mask = catalog[col_cod].astype(str) == str(codigo_api)
    if not mask.any():
        return {"metrica": "Sistema", "max_dias": 365}

    fila = catalog[mask].iloc[0]
    metrica = str(fila.get("Granularidad", fila.get("Metrica", "Sistema")))
    max_raw = fila.get("Máximo Días", fila.get("MaxDias", 365))

    return {"metrica": metrica, "max_dias": _sanitize_max_dias(max_raw)}


# -------------------------------------------------------------------
# CACHE LOCAL DE XM
# -------------------------------------------------------------------
def _cache_file_path(coleccion: str, metrica: str) -> Path:
    safe_col = "".join(ch if ch.isalnum() else "_" for ch in str(coleccion))
    safe_met = "".join(ch if ch.isalnum() else "_" for ch in str(metrica))
    return CACHE_DIR / f"xm_{safe_col}_{safe_met}.csv"


@st.cache_data
def _load_xm_cache(coleccion: str, metrica: str) -> Optional[pd.DataFrame]:
    path = _cache_file_path(coleccion, metrica)
    if path.exists():
        df = pd.read_csv(path)
        df.columns = [c.strip() for c in df.columns]
        return df
    return None


def _save_xm_cache(coleccion: str, metrica: str, df: pd.DataFrame) -> None:
    try:
        path = _cache_file_path(coleccion, metrica)
        CACHE_DIR.mkdir(exist_ok=True, parents=True)
        df.to_csv(path, index=False)
    except Exception:
        pass


def _fetch_xm_from_api(
    coleccion: str,
    metrica: str,
    start_date: dt.date,
    end_date: dt.date,
    max_dias: int,
) -> pd.DataFrame:
    api = ReadDB()
    dfs: List[pd.DataFrame] = []
    max_dias = _sanitize_max_dias(max_dias)
    current_start = start_date

    while current_start <= end_date:
        current_end = min(current_start + dt.timedelta(days=max_dias - 1), end_date)
        try:
            df_tmp = api.request_data(
                coleccion=coleccion,
                metrica=metrica,
                start_date=str(current_start),
                end_date=str(current_end),
                filtros=None,
            )
        except Exception as e:
            # Aviso corto, sin romper la app
            st.info(
                f"XM no entregó datos válidos para {coleccion}-{metrica} "
                f"entre {current_start} y {current_end} ({e}). Se continúa con el resto del período."
            )
            break

        if df_tmp is not None and not df_tmp.empty:
            dfs.append(df_tmp)

        current_start = current_end + dt.timedelta(days=1)

    if not dfs:
        return pd.DataFrame()

    df = pd.concat(dfs, ignore_index=True).drop_duplicates()
    df.columns = [c.strip() for c in df.columns]
    return df


def fetch_xm_data_chunked(
    coleccion: str,
    metrica: str,
    start_date: dt.date,
    end_date: dt.date,
    max_dias: int,
) -> pd.DataFrame:
    # 1) Intentar leer cache local
    df_cache = _load_xm_cache(coleccion, metrica)

    if df_cache is None or df_cache.empty:
        df_new = _fetch_xm_from_api(
            coleccion=coleccion,
            metrica=metrica,
            start_date=start_date,
            end_date=end_date,
            max_dias=max_dias,
        )
        if not df_new.empty:
            _save_xm_cache(coleccion, metrica, df_new)
        return df_new

    # Con cache: detectar si faltan rangos
    try:
        fecha_col = _detectar_col_fecha(df_cache)
        dfc = df_cache.copy()
        dfc[fecha_col] = pd.to_datetime(dfc[fecha_col], errors="coerce").dt.date
        cache_min = dfc[fecha_col].min()
        cache_max = dfc[fecha_col].max()
    except Exception:
        df_new = _fetch_xm_from_api(
            coleccion=coleccion,
            metrica=metrica,
            start_date=start_date,
            end_date=end_date,
            max_dias=max_dias,
        )
        if not df_new.empty:
            _save_xm_cache(coleccion, metrica, df_new)
        return df_new

    missing_ranges: List[Tuple[dt.date, dt.date]] = []

    if cache_min is None or cache_max is None or pd.isna(cache_min) or pd.isna(cache_max):
        missing_ranges.append((start_date, end_date))
    else:
        if start_date < cache_min:
            ms = start_date
            me = min(cache_min - dt.timedelta(days=1), end_date)
            if ms <= me:
                missing_ranges.append((ms, me))
        if end_date > cache_max:
            ms = max(start_date, cache_max + dt.timedelta(days=1))
            me = end_date
            if ms <= me:
                missing_ranges.append((ms, me))

    # Descargar solo tramos faltantes
    if missing_ranges:
        dfs_new = []
        for s, e in missing_ranges:
            df_seg = _fetch_xm_from_api(
                coleccion=coleccion,
                metrica=metrica,
                start_date=s,
                end_date=e,
                max_dias=max_dias,
            )
            if df_seg is not None and not df_seg.empty:
                dfs_new.append(df_seg)
        if dfs_new:
            df_all = pd.concat([dfc] + dfs_new, ignore_index=True).drop_duplicates()
            _save_xm_cache(coleccion, metrica, df_all)
            dfc = df_all

    # Devolver solo el rango solicitado
    try:
        fecha_col = _detectar_col_fecha(dfc)
        dfc[fecha_col] = pd.to_datetime(dfc[fecha_col], errors="coerce").dt.date
        mask = (dfc[fecha_col] >= start_date) & (dfc[fecha_col] <= end_date)
        return dfc.loc[mask].reset_index(drop=True)
    except Exception:
        return dfc


# -------------------------------------------------------------------
# HIDROLOGÍA – CONSTRUCCIÓN DE SERIES
# -------------------------------------------------------------------
@st.cache_data
def build_hidro_diaria(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    """
    Tabla diaria a nivel SIN con variables hidrológicas:

        Fecha
        nivel_embalse_pct
        energia_embalsada_gwh
        aportes_gwh
        aportes_hist_gwh
        demanda_gwh
        aportes_pct_hist
        dias_respaldo
        delta_embalse_gwh
    """
    variables = {
        "nivel_embalse_pct": ("PorcVoluUtilDiar", "mean"),
        "energia_embalsada_gwh": ("VoluUtilDiarEner", "sum"),
        "aportes_gwh": ("AporEner", "sum"),
        "aportes_hist_gwh": ("AporEnerMediHist", "sum"),
        "demanda_gwh": ("DemaSIN", "sum"),
    }

    series = []

    for nombre, (codigo_api, agg) in variables.items():
        meta = get_xm_meta(codigo_api)
        df_raw = fetch_xm_data_chunked(
            coleccion=codigo_api,
            metrica=meta["metrica"],
            start_date=start_date,
            end_date=end_date,
            max_dias=meta["max_dias"],
        )
        if df_raw.empty:
            continue

        fecha_col = _detectar_col_fecha(df_raw)
        valor_col = _detectar_col_valor(df_raw)
        if valor_col is None:
            continue

        df_tmp = df_raw[[fecha_col, valor_col]].copy()
        df_tmp[fecha_col] = pd.to_datetime(df_tmp[fecha_col], errors="coerce").dt.date
        df_tmp = df_tmp.dropna(subset=[fecha_col])

        if agg == "mean":
            s = df_tmp.groupby(fecha_col)[valor_col].mean()
        else:
            s = df_tmp.groupby(fecha_col)[valor_col].sum()

        s.name = nombre
        series.append(s)

    if not series:
        return pd.DataFrame()

    df_all = pd.concat(series, axis=1).sort_index()
    df_all.index.name = "Fecha"
    df_all = df_all.reset_index()

    # Ajuste unidades: energía → GWh si vienen gigantes
    energy_cols = [
        "energia_embalsada_gwh",
        "aportes_gwh",
        "aportes_hist_gwh",
        "demanda_gwh",
    ]
    for col in energy_cols:
        if col in df_all.columns:
            mean_val = abs(df_all[col]).mean()
            if pd.notna(mean_val) and mean_val > 1_000_000:
                df_all[col] = df_all[col] / 1_000_000.0

    # Nivel 0–1 → %
    if "nivel_embalse_pct" in df_all.columns:
        max_val = df_all["nivel_embalse_pct"].max()
        if pd.notna(max_val) and max_val <= 2:
            df_all["nivel_embalse_pct"] = df_all["nivel_embalse_pct"] * 100.0

    # Derivadas
    if {"aportes_gwh", "aportes_hist_gwh"}.issubset(df_all.columns):
        df_all["aportes_pct_hist"] = (
            100.0 * df_all["aportes_gwh"] / df_all["aportes_hist_gwh"]
        )

    if {"energia_embalsada_gwh", "demanda_gwh"}.issubset(df_all.columns):
        df_all["dias_respaldo"] = (
            df_all["energia_embalsada_gwh"] / df_all["demanda_gwh"]
        )

    if "energia_embalsada_gwh" in df_all.columns:
        df_all["delta_embalse_gwh"] = df_all["energia_embalsada_gwh"].diff()

    return df_all

@st.cache_data
def build_precio_bolsa_diario(
    start_date: dt.date,
    end_date: dt.date,
) -> pd.DataFrame:
    """
    Precio de Bolsa de Energía diario (promedio diario).

    Devuelve:
        Fecha
        precio_bolsa   -> en la unidad que reporte XM (p.ej. $/MWh)
    """
    meta = get_xm_meta(CODIGO_API_PRECIO_BOLSA)
    df_raw = fetch_xm_data_chunked(
        coleccion=CODIGO_API_PRECIO_BOLSA,
        metrica=meta["metrica"],
        start_date=start_date,
        end_date=end_date,
        max_dias=meta["max_dias"],
    )

    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    fecha_col = _detectar_col_fecha(df_raw)
    valor_col = _detectar_col_valor(df_raw)
    if fecha_col is None or valor_col is None:
        return pd.DataFrame()

    df = df_raw[[fecha_col, valor_col]].copy()
    df[fecha_col] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date
    df = df.dropna(subset=[fecha_col])

    # Promedio diario
    df_d = df.groupby(fecha_col)[valor_col].mean().reset_index()
    df_d.rename(
        columns={
            fecha_col: "Fecha",
            valor_col: "precio_bolsa",
        },
        inplace=True,
    )

    return df_d.sort_values("Fecha")

@st.cache_data
def build_detalle_embalses(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    """
    Resumen por embalse con nivel promedio (%) en el rango:

        Embalse
        nivel_promedio_pct
    """
    try:
        df_raw = fetch_xm_data_chunked(
            coleccion="PorcVoluUtilDiar",
            metrica="Embalse",
            start_date=start_date,
            end_date=end_date,
            max_dias=365,
        )
    except Exception:
        return pd.DataFrame()

    if df_raw.empty:
        return pd.DataFrame()

    fecha_col = _detectar_col_fecha(df_raw)
    valor_col = _detectar_col_valor(df_raw)
    emb_col = _detectar_col_embalse(df_raw)

    df = df_raw[[fecha_col, emb_col, valor_col]].copy()
    df[fecha_col] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date
    df = df.dropna(subset=[fecha_col])

    df = df[(df[fecha_col] >= start_date) & (df[fecha_col] <= end_date)]

    df_emb = df.groupby(emb_col)[valor_col].mean().reset_index()

    df_emb.rename(
        columns={
            emb_col: "Embalse",
            valor_col: "nivel_promedio_pct",
        },
        inplace=True,
    )

    max_val = df_emb["nivel_promedio_pct"].max()
    if pd.notna(max_val) and max_val <= 2:
        df_emb["nivel_promedio_pct"] = df_emb["nivel_promedio_pct"] * 100.0

    return df_emb.sort_values("nivel_promedio_pct", ascending=False)


@st.cache_data
def build_capacidad_embalses(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    """
    Capacidad energética por embalse (en el rango):

        Embalse
        energia_promedio_gwh
        energia_max_gwh
        energia_min_gwh
    """
    try:
        df_raw = fetch_xm_data_chunked(
            coleccion="VoluUtilDiarEner",
            metrica="Embalse",
            start_date=start_date,
            end_date=end_date,
            max_dias=365,
        )
    except Exception:
        return pd.DataFrame()

    if df_raw.empty:
        return pd.DataFrame()

    fecha_col = _detectar_col_fecha(df_raw)
    valor_col = _detectar_col_valor(df_raw)
    emb_col = _detectar_col_embalse(df_raw)

    df = df_raw[[fecha_col, emb_col, valor_col]].copy()
    df[fecha_col] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date
    df = df.dropna(subset=[fecha_col])

    df = df[(df[fecha_col] >= start_date) & (df[fecha_col] <= end_date)]

    df_grp = df.groupby(emb_col)[valor_col].agg(["mean", "max", "min"]).reset_index()
    df_grp.rename(
        columns={
            emb_col: "Embalse",
            "mean": "energia_promedio_raw",
            "max": "energia_max_raw",
            "min": "energia_min_raw",
        },
        inplace=True,
    )

    mean_val = abs(df_grp["energia_promedio_raw"]).mean()
    factor = 1.0
    if pd.notna(mean_val) and mean_val > 1_000_000:
        factor = 1_000_000.0

    df_grp["energia_promedio_gwh"] = df_grp["energia_promedio_raw"] / factor
    df_grp["energia_max_gwh"] = df_grp["energia_max_raw"] / factor
    df_grp["energia_min_gwh"] = df_grp["energia_min_raw"] / factor

    return df_grp[["Embalse", "energia_promedio_gwh", "energia_max_gwh", "energia_min_gwh"]]


@st.cache_data
def get_embalse_hist_extremos(embalse: str) -> Tuple[float, float]:
    """
    Mínimo y máximo histórico del nivel útil [%] de un embalse,
    usando PorcVoluUtilDiar desde GLOBAL_EARLIEST_DATE.
    """
    df_raw = fetch_xm_data_chunked(
        coleccion="PorcVoluUtilDiar",
        metrica="Embalse",
        start_date=GLOBAL_EARLIEST_DATE,
        end_date=dt.date.today(),
        max_dias=365,
    )
    if df_raw.empty:
        return float("nan"), float("nan")

    valor_col = _detectar_col_valor(df_raw)
    emb_col = _detectar_col_embalse(df_raw)

    df = df_raw[[emb_col, valor_col]].copy()
    df = df[df[emb_col].astype(str) == str(embalse)]

    if df.empty:
        return float("nan"), float("nan")

    serie = pd.to_numeric(df[valor_col], errors="coerce").dropna()
    if serie.empty:
        return float("nan"), float("nan")

    # Escalar a % si viene 0–1
    max_val = serie.max()
    if pd.notna(max_val) and max_val <= 2:
        serie = serie * 100.0

    return float(serie.min()), float(serie.max())


@st.cache_data
def get_embalse_hist_energia_extremos(embalse: str) -> Tuple[float, float]:
    """
    Mínimo y máximo histórico de energía embalsada [GWh] para un embalse,
    usando VoluUtilDiarEner desde GLOBAL_EARLIEST_DATE.
    """
    df_raw = fetch_xm_data_chunked(
        coleccion="VoluUtilDiarEner",
        metrica="Embalse",
        start_date=GLOBAL_EARLIEST_DATE,
        end_date=dt.date.today(),
        max_dias=365,
    )
    if df_raw.empty:
        return float("nan"), float("nan")

    valor_col = _detectar_col_valor(df_raw)
    emb_col = _detectar_col_embalse(df_raw)

    df = df_raw[[emb_col, valor_col]].copy()
    df = df[df[emb_col].astype(str) == str(embalse)]
    if df.empty:
        return float("nan"), float("nan")

    serie = pd.to_numeric(df[valor_col], errors="coerce").dropna()
    if serie.empty:
        return float("nan"), float("nan")

    mean_val = abs(serie).mean()
    factor = 1.0
    if pd.notna(mean_val) and mean_val > 1_000_000:
        factor = 1_000_000.0

    serie_gwh = serie / factor
    return float(serie_gwh.min()), float(serie_gwh.max())


@st.cache_data
def build_serie_embalse(
    embalse: str, start_date: dt.date, end_date: dt.date
) -> pd.DataFrame:
    """
    Serie diaria para un embalse:

        Fecha
        nivel_embalse_pct
        energia_embalsada_gwh
        delta_embalse_gwh
    """
    # Nivel
    df_nivel = fetch_xm_data_chunked(
        coleccion="PorcVoluUtilDiar",
        metrica="Embalse",
        start_date=start_date,
        end_date=end_date,
        max_dias=365,
    )
    # Energía
    df_ener = fetch_xm_data_chunked(
        coleccion="VoluUtilDiarEner",
        metrica="Embalse",
        start_date=start_date,
        end_date=end_date,
        max_dias=365,
    )

    if df_nivel.empty and df_ener.empty:
        return pd.DataFrame()

    # Nivel
    if not df_nivel.empty:
        f_n = _detectar_col_fecha(df_nivel)
        v_n = _detectar_col_valor(df_nivel)
        e_n = _detectar_col_embalse(df_nivel)
        df_n = df_nivel[[f_n, e_n, v_n]].copy()
        df_n[f_n] = pd.to_datetime(df_n[f_n], errors="coerce").dt.date
        df_n = df_n.dropna(subset=[f_n])
        df_n = df_n[df_n[e_n].astype(str) == str(embalse)]
        df_n = df_n.groupby(f_n)[v_n].mean().reset_index()
        df_n.rename(columns={f_n: "Fecha", v_n: "nivel_embalse_pct"}, inplace=True)
        max_val = df_n["nivel_embalse_pct"].max()
        if pd.notna(max_val) and max_val <= 2:
            df_n["nivel_embalse_pct"] *= 100.0
    else:
        df_n = pd.DataFrame(columns=["Fecha", "nivel_embalse_pct"])

    # Energía
    if not df_ener.empty:
        f_e = _detectar_col_fecha(df_ener)
        v_e = _detectar_col_valor(df_ener)
        e_e = _detectar_col_embalse(df_ener)
        df_e = df_ener[[f_e, e_e, v_e]].copy()
        df_e[f_e] = pd.to_datetime(df_e[f_e], errors="coerce").dt.date
        df_e = df_e.dropna(subset=[f_e])
        df_e = df_e[df_e[e_e].astype(str) == str(embalse)]
        df_e = df_e.groupby(f_e)[v_e].sum().reset_index()
        df_e.rename(columns={f_e: "Fecha", v_e: "energia_embalsada_raw"}, inplace=True)

        mean_val = abs(df_e["energia_embalsada_raw"]).mean()
        factor = 1.0
        if pd.notna(mean_val) and mean_val > 1_000_000:
            factor = 1_000_000.0
        df_e["energia_embalsada_gwh"] = df_e["energia_embalsada_raw"] / factor
        df_e = df_e[["Fecha", "energia_embalsada_gwh"]]
    else:
        df_e = pd.DataFrame(columns=["Fecha", "energia_embalsada_gwh"])

    df = pd.merge(df_n, df_e, on="Fecha", how="outer").sort_values("Fecha")
    if "energia_embalsada_gwh" in df.columns:
        df["delta_embalse_gwh"] = df["energia_embalsada_gwh"].diff()
    return df


@st.cache_data
def build_aportes_rio_diario(
    start_date: dt.date, end_date: dt.date
) -> pd.DataFrame:
    """
    Aportes por río (ajusta la métrica 'Rio' según tu catálogo XM si es distinta).

        Fecha
        Rio
        aportes_gwh
    """
    try:
        df_raw = fetch_xm_data_chunked(
            coleccion="AporEner",
            metrica="Rio",   # ⇐ AJUSTAR SI EN EL CATÁLOGO ES OTRA COSA
            start_date=start_date,
            end_date=end_date,
            max_dias=365,
        )
    except Exception:
        return pd.DataFrame()

    if df_raw.empty:
        return pd.DataFrame()

    fecha_col = _detectar_col_fecha(df_raw)
    valor_col = _detectar_col_valor(df_raw)
    rio_col = _detectar_col_rio(df_raw)

    df = df_raw[[fecha_col, rio_col, valor_col]].copy()
    df[fecha_col] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date
    df = df.dropna(subset=[fecha_col])

    df_grp = (
        df.groupby([fecha_col, rio_col])[valor_col].sum().reset_index()
    )
    df_grp.rename(
        columns={
            fecha_col: "Fecha",
            rio_col: "Rio",
            valor_col: "aportes_raw",
        },
        inplace=True,
    )

    mean_val = abs(df_grp["aportes_raw"]).mean()
    factor = 1.0
    if pd.notna(mean_val) and mean_val > 1_000_000:
        factor = 1_000_000.0

    df_grp["aportes_gwh"] = df_grp["aportes_raw"] / factor
    return df_grp[["Fecha", "Rio", "aportes_gwh"]]


# -------------------------------------------------------------------
# SIMEM – GENERACIÓN DETALLADA (E17D25)
# -------------------------------------------------------------------
@st.cache_data
def consultar_generacion_detallada_simem(
    start_date: dt.date,
    end_date: dt.date,
    chunk_days: int = 90,
) -> pd.DataFrame:
    dataset_id = "E17D25"
    dfs = []
    current_start = start_date
    while current_start <= end_date:
        current_end = min(
            current_start + dt.timedelta(days=chunk_days - 1),
            end_date,
        )
        simem = ReadSIMEM(
            dataset_id=dataset_id,
            start_date=str(current_start),
            end_date=str(current_end),
            filter_column=None,
            filter_values=None,
        )
        df_tmp = simem.main(filter=False)
        if df_tmp is not None and not df_tmp.empty:
            df_tmp.columns = [c.strip() for c in df_tmp.columns]
            dfs.append(df_tmp)
        current_start = current_end + dt.timedelta(days=1)
    if not dfs:
        return pd.DataFrame()
    df = pd.concat(dfs, ignore_index=True).drop_duplicates()
    df.columns = [c.strip() for c in df.columns]
    return df


@st.cache_data
def generar_tabla_generacion_enriquecida(
    start_date: dt.date,
    end_date: dt.date,
) -> pd.DataFrame:
    df_gen = consultar_generacion_detallada_simem(start_date, end_date)
    if df_gen.empty:
        return df_gen

    df_rec = load_recursos_xm()
    if df_rec.empty:
        return df_gen

    df_gen.columns = [c.strip() for c in df_gen.columns]
    df_rec.columns = [c.strip() for c in df_rec.columns]

    # Código recurso en generación
    col_codigo_gen = None
    for c in df_gen.columns:
        cl = c.lower()
        if "codigoplanta" in cl or "codrecurso" in cl or "codigorecurso" in cl:
            col_codigo_gen = c
            break
    if col_codigo_gen is None:
        if "CodigoPlanta" in df_gen.columns:
            col_codigo_gen = "CodigoPlanta"
        else:
            return df_gen

    col_codigo_rec = None
    for c in df_rec.columns:
        cl = c.lower()
        if "codigoplanta" in cl or "codigorecurso" in cl or "idrecurso" in cl:
            col_codigo_rec = c
            break
    if col_codigo_rec is None:
        return df_gen

    cols_info = [col_codigo_rec]
    for c in df_rec.columns:
        cl = c.lower()
        if (
            ("nombre" in cl and ("recurso" in cl or "planta" in cl or "central" in cl))
            or c in ["Recurso", "Planta", "Central"]
        ):
            if c not in cols_info:
                cols_info.append(c)
        if any(k in cl for k in ["municipio", "departamento", "agente"]):
            if c not in cols_info:
                cols_info.append(c)

    df_rec_sub = df_rec[cols_info].drop_duplicates(subset=[col_codigo_rec])

    df_merged = df_gen.merge(
        df_rec_sub,
        left_on=col_codigo_gen,
        right_on=col_codigo_rec,
        how="left",
        suffixes=("", "_cat"),
    )

    front_cols: List[str] = []
    for c in ["Fecha", col_codigo_gen]:
        if c in df_merged.columns and c not in front_cols:
            front_cols.append(c)

    for c in df_merged.columns:
        cl = c.lower()
        if (
            ("nombre" in cl and ("planta" in cl or "recurso" in cl or "central" in cl))
            or c in ["Recurso", "Planta", "Central"]
        ):
            if c not in front_cols:
                front_cols.append(c)

    for c in df_merged.columns:
        cl = c.lower()
        if any(k in cl for k in ["municipio", "departamento", "agente"]):
            if c not in front_cols:
                front_cols.append(c)

    other_cols = [c for c in df_merged.columns if c not in front_cols]
    df_merged = df_merged[front_cols + other_cols]
    return df_merged


def detect_gen_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Detecta nombres de columnas relevantes en la tabla de generación."""
    cols = [c for c in df.columns]

    def find(*keywords) -> Optional[str]:
        for c in cols:
            cl = c.lower()
            if all(k in cl for k in keywords):
                return c
        return None

    # Fecha (en muchos datasets se llama exactamente 'Fecha')
    col_fecha = "Fecha" if "Fecha" in df.columns else find("fecha")

    # Nombre planta / recurso (por si lo necesitas en otros módulos)
    col_planta = None
    for k in ["planta", "nombreplanta", "recurso", "central"]:
        c = find(k)
        if c:
            col_planta = c
            break

    # Tipo de fuente / tipo de generación (hidráulica, térmica, etc.)
    col_tipo_fuente = find("tipogener") or find("tipofuente") or find("tiporecur") or find("tipo")

    # Otros (despacho, GD) – se mantienen para otras partes del visor
    col_tipo_despacho = find("tipodespacho")
    col_gd = find("gd")

    # Columna numérica de generación/energía
    # Evitamos columnas que tengan 'tipo' en el nombre, para no confundir con 'TipoGeneracion'
    gen_cols = [
        c for c in cols
        if (("gener" in c.lower() or "energ" in c.lower()) and "tipo" not in c.lower())
    ]
    col_gen = gen_cols[0] if gen_cols else None

    return {
        "fecha": col_fecha,
        "planta": col_planta,
        "tipo_fuente": col_tipo_fuente,
        "tipo_despacho": col_tipo_despacho,
        "gd": col_gd,
        "gen": col_gen,
    }


@st.cache_data
def renewable_mask(tipo_fuente_serie: pd.Series) -> pd.Series:
    if tipo_fuente_serie is None:
        return pd.Series([], dtype=bool)
    txt = tipo_fuente_serie.astype(str).str.lower()
    renovables = [
        "solar",
        "eolica",
        "eólico",
        "hidraul",
        "biomasa",
        "geoterm",
        "ocea",
        "mareo",
        "viento",
    ]
    mask = pd.Series(False, index=tipo_fuente_serie.index)
    for k in renovables:
        mask = mask | txt.str.contains(k)
    return mask


# -------------------------------------------------------------------
# UI – UTILIDAD PARA RANGOS CON PERIODOS PREDEFINIDOS
# -------------------------------------------------------------------
# -------------------------------------------------------------------
# UI – UTILIDAD PARA RANGOS CON PERIODOS PREDEFINIDOS
# -------------------------------------------------------------------
def select_rango_fechas_presets(
    key_prefix: str,
    default_years: int = 1,
) -> Tuple[dt.date, dt.date]:
    """
    Selector de fechas con periodos:
    - Último mes, 6 meses, 1, 2, 5 y 10 años
    - Personalizado (permite mover manualmente los calendarios)
    """
    hoy = dt.date.today()

    opciones = [
        ("Último mes", 30),
        ("Últimos 6 meses", 182),
        ("Último año", 365),
        ("Últimos 2 años", 365 * 2),
        ("Últimos 5 años", 365 * 5),
        ("Últimos 10 años", 365 * 10),
        ("Personalizado", None),
    ]
    labels = [o[0] for o in opciones]

    opcion = st.selectbox(
        "Periodo",
        labels,
        index=2,  # por defecto: último año
        key=f"{key_prefix}_periodo",
    )

    dias = dict(opciones)[opcion]

    # Caso PERSONALIZADO: mostramos los date_input normales
    if dias is None:
        default_start = dt.date.today() - dt.timedelta(days=365 * default_years)
        default_end = dt.date.today()

        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input(
                "Fecha inicio",
                value=default_start,
                format="YYYY-MM-DD",
                key=f"{key_prefix}_start",
            )
        with col2:
            end_date = st.date_input(
                "Fecha fin",
                value=default_end,
                format="YYYY-MM-DD",
                key=f"{key_prefix}_end",
            )

    # Caso con preset fijo (1 mes, 6 meses, 1 año, etc.)
    else:
        start_date = hoy - dt.timedelta(days=dias)
        end_date = hoy
        # Solo mostramos una leyenda para que el usuario sepa qué rango se está usando
        st.caption(
            f"Período seleccionado: desde **{start_date}** hasta **{end_date}** "
            f"({dias} días). Si quieres cambiarlo usa la opción *Personalizado*."
        )

    return start_date, end_date


def ui_hidrologia() -> None:
    st.subheader("Hidrología – Embalses y aportes")
    st.caption("Seguimiento hidrológico del SIN y seguridad energética hídrica")

    # Selección de periodo con presets
    start_date, end_date = select_rango_fechas_presets(
        key_prefix="hidro",
        default_years=3,
    )

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    with st.spinner("Consultando XM y construyendo serie hidrológica..."):
        df = build_hidro_diaria(start_date, end_date)

    if df.empty:
        st.warning("No se obtuvieron datos para el rango seleccionado.")
        return

    df = df.sort_values("Fecha")
    last_row = df.iloc[-1]
    first_row = df.iloc[0]

    # ------------------ KPIs compactos ------------------
    c1, c2, c3, c4 = st.columns(4)

    if "nivel_embalse_pct" in df.columns:
        delta_nivel = last_row["nivel_embalse_pct"] - first_row["nivel_embalse_pct"]
        with c1:
            st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
            st.metric(
                "Nivel de embalses",
                f"{last_row['nivel_embalse_pct']:.2f} %",
                f"{delta_nivel:+.2f} pp",
            )
            st.markdown("</div>", unsafe_allow_html=True)

    if "energia_embalsada_gwh" in df.columns:
        with c2:
            st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
            st.metric(
                "Energía embalsada",
                f"{last_row['energia_embalsada_gwh']:,.0f} GWh",
            )
            st.markdown("</div>", unsafe_allow_html=True)

    if "aportes_pct_hist" in df.columns:
        delta_ap = last_row["aportes_pct_hist"] - df["aportes_pct_hist"].mean()
        with c3:
            st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
            st.metric(
                "% Aportes vs histórico",
                f"{last_row['aportes_pct_hist']:.1f} %",
                f"{delta_ap:+.1f} pts",
            )
            st.markdown("</div>", unsafe_allow_html=True)

    if "dias_respaldo" in df.columns:
        with c4:
            st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
            st.metric(
                "Días de respaldo hidro",
                f"{last_row['dias_respaldo']:.1f} días",
            )
            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    # Precalcular detalles de embalses y ríos (se usan en tabs)
    df_emb = build_detalle_embalses(start_date, end_date)
    df_cap = build_capacidad_embalses(start_date, end_date)
    df_rio = build_aportes_rio_diario(start_date, end_date)

    tabs = st.tabs(
        ["SIN general", "Embalses y cuencas", "Ríos / cuencas", "Embalse individual"]
    )

    # ------------------------------------------------------------------
    # TAB 1 – SIN GENERAL
    # ------------------------------------------------------------------
    with tabs[0]:
        col_a, col_b = st.columns(2)

        if "nivel_embalse_pct" in df.columns:
            with col_a:
                fig1 = px.area(
                    df,
                    x="Fecha",
                    y="nivel_embalse_pct",
                    title="Nivel de embalses (%)",
                )
                fig1.update_traces(
                    line_color=PASTEL_BLUE,
                    fillcolor="rgba(144,202,249,0.35)",
                )
                fig1.update_layout(
                    template="plotly_dark",
                    yaxis_title="% volumen útil",
                    height=320,
                    margin=dict(l=20, r=10, t=40, b=30),
                )
                st.plotly_chart(fig1, use_container_width=True)

        if {"aportes_gwh", "aportes_hist_gwh"}.issubset(df.columns):
            with col_b:
                fig2 = go.Figure()
                fig2.add_trace(
                    go.Scatter(
                        x=df["Fecha"],
                        y=df["aportes_gwh"],
                        mode="lines",
                        name="Aportes [GWh/día]",
                        line=dict(color="#1B8C3B", width=2),
                        fill="tozeroy",
                        fillcolor="rgba(27,140,59,0.30)",
                    )
                )
                fig2.add_trace(
                    go.Scatter(
                        x=df["Fecha"],
                        y=df["aportes_hist_gwh"],
                        mode="lines",
                        name="Media histórica",
                        line=dict(color="#0B9999", width=3),
                    )
                )
                fig2.update_layout(
                    template="plotly_dark",
                    title="Aportes hídricos vs media histórica",
                    yaxis_title="GWh/día",
                    height=320,
                    margin=dict(l=20, r=10, t=40, b=30),
                )
                st.plotly_chart(fig2, use_container_width=True)

        col_c, col_d = st.columns(2)

        if "delta_embalse_gwh" in df.columns:
            with col_c:
                df_bal = df.dropna(subset=["delta_embalse_gwh"]).copy()
                colores = [
                    "#1B8C3B" if v >= 0 else "#C7253E"
                    for v in df_bal["delta_embalse_gwh"]
                ]
                fig_bal = go.Figure()
                fig_bal.add_trace(
                    go.Bar(
                        x=df_bal["Fecha"],
                        y=df_bal["delta_embalse_gwh"],
                        marker_color=colores,
                        name="Δ energía [GWh/día]",
                    )
                )
                fig_bal.update_layout(
                    template="plotly_dark",
                    title="Variación diaria de energía embalsada",
                    yaxis_title="GWh/día",
                    xaxis_title="Fecha",
                    bargap=0,
                    height=280,
                    margin=dict(l=20, r=10, t=40, b=30),
                )
                fig_bal.update_yaxes(
                    zeroline=True, zerolinewidth=1, zerolinecolor="#FFFFFF"
                )
                st.plotly_chart(fig_bal, use_container_width=True)

        if "dias_respaldo" in df.columns:
            with col_d:
                fig3 = px.area(
                    df,
                    x="Fecha",
                    y="dias_respaldo",
                    title="Días de respaldo con energía embalsada",
                )
                fig3.update_traces(
                    line_color=PRIMARY_PURPLE,
                    fillcolor="rgba(206,147,216,0.25)",
                )
                fig3.update_layout(
                    template="plotly_dark",
                    yaxis_title="días",
                    height=280,
                    margin=dict(l=20, r=10, t=40, b=30),
                )
                st.plotly_chart(fig3, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 2 – EMBALSES Y CUENCAS
    # ------------------------------------------------------------------
    with tabs[1]:
        if df_emb.empty:
            st.info(
                "No se pudo obtener el detalle por embalse para este rango de fechas."
            )
        else:
            df_seg = df_emb.merge(df_cap, on="Embalse", how="left")

            c1, c2 = st.columns(2)
            c1.metric("Embalses", len(df_emb))
            c2.metric(
                "Nivel promedio de embalses",
                f"{df_emb['nivel_promedio_pct'].mean():.1f} %",
            )

            # Índice de seguridad energética: tamaño * nivel
            if "energia_max_gwh" in df_seg.columns:
                df_seg["indice_seguridad"] = (
                    df_seg["energia_max_gwh"].fillna(0.0)
                    * df_seg["nivel_promedio_pct"].fillna(0.0)
                    / 100.0
                )
            else:
                df_seg["indice_seguridad"] = df_seg["nivel_promedio_pct"]

            df_seg = df_seg.sort_values("indice_seguridad", ascending=False)

            # ----- Mapa -----
            df_seg["Embalse_norm"] = df_seg["Embalse"].apply(_normalize_embalse_name)
            df_meta = EMBALSES_SIN_META.copy()
            df_map = df_seg.merge(
                df_meta,
                on="Embalse_norm",
                how="left",
                suffixes=("", "_meta"),
            )

            faltantes = df_map[df_map["Lat"].isna()]["Embalse"].unique()
            if len(faltantes) > 0:
                st.warning(
                    "Sin coordenadas en EMBALSES_SIN_META para: "
                    + ", ".join(str(x) for x in faltantes)
                )

            df_map_plot = df_map.dropna(subset=["Lat", "Lon"]).copy()
            if not df_map_plot.empty:
                def _nivel_riesgo(nivel_pct: float) -> str:
                    if pd.isna(nivel_pct):
                        return "Sin dato"
                    if nivel_pct < 40:
                        return "Riesgo MEDIO"
                    else:
                        return "Riesgo BAJO"

                df_map_plot["NivelRiesgo"] = df_map_plot["nivel_promedio_pct"].apply(
                    _nivel_riesgo
                )

                if (
                    "energia_max_gwh" in df_map_plot.columns
                    and df_map_plot["energia_max_gwh"].notna().any()
                ):
                    size_col = "energia_max_gwh"
                else:
                    size_col = "indice_seguridad"

                color_map = {
                    "Riesgo MEDIO": "#FFC107",
                    "Riesgo BAJO": "#4CAF50",
                    "Sin dato": "#9E9E9E",
                }

                df_map_plot["Hidroeléctrica_asociada"] = df_map_plot["Hidro_asociada"]

                fig_map = px.scatter_mapbox(
                    df_map_plot,
                    lat="Lat",
                    lon="Lon",
                    size=size_col,
                    color="NivelRiesgo",
                    hover_name="Embalse",
                    hover_data={
                        "Hidroeléctrica_asociada": True,
                        "Rio": True,
                        "Cuenca": True,
                        "nivel_promedio_pct": ":.1f",
                        "energia_max_gwh": ":.0f"
                        if "energia_max_gwh" in df_map_plot.columns
                        else False,
                    },
                    zoom=5.6,
                    height=420,
                    title="Embalses conectados al SIN – río y cuenca",
                )

                max_size_val = df_map_plot[size_col].max()
                if pd.notna(max_size_val) and max_size_val > 0:
                    fig_map.update_traces(
                        marker=dict(
                            sizemode="area",
                            sizeref=2.0 * max_size_val / (40.0**2),
                            sizemin=8,
                        )
                    )

                fig_map.update_layout(
                    mapbox_style="open-street-map",
                    template="plotly_dark",
                    margin=dict(l=0, r=0, t=40, b=0),
                    legend_title_text="Nivel de riesgo",
                )
                for tr in fig_map.data:
                    if tr.name in color_map:
                        tr.marker.color = color_map[tr.name]

                st.plotly_chart(fig_map, use_container_width=True)

            col_t1, col_t2 = st.columns(2)

            if "energia_max_gwh" in df_seg.columns:
                with col_t1:
                    top_energia = (
                        df_seg.sort_values("energia_max_gwh", ascending=False)
                        .head(10)
                        .sort_values("energia_max_gwh")
                    )
                    fig_ene = px.bar(
                        top_energia,
                        x="energia_max_gwh",
                        y="Embalse",
                        orientation="h",
                        title="Top 10 por energía almacenada [GWh]",
                    )
                    fig_ene.update_layout(
                        template="plotly_dark",
                        height=320,
                        margin=dict(l=60, r=10, t=40, b=30),
                    )
                    st.plotly_chart(fig_ene, use_container_width=True)

            with col_t2:
                top_nivel = (
                    df_seg.sort_values("nivel_promedio_pct", ascending=False)
                    .head(10)
                    .sort_values("nivel_promedio_pct")
                )
                fig_b = px.bar(
                    top_nivel,
                    x="nivel_promedio_pct",
                    y="Embalse",
                    orientation="h",
                    title="Top 10 por nivel promedio (%)",
                )
                fig_b.update_layout(
                    template="plotly_dark",
                    height=320,
                    margin=dict(l=60, r=10, t=40, b=30),
                )
                st.plotly_chart(fig_b, use_container_width=True)

            # Análisis textual corto
            if not df_seg.empty and "energia_max_gwh" in df_seg.columns:
                top3 = df_seg.head(3)[
                    ["Embalse", "nivel_promedio_pct", "energia_max_gwh"]
                ]
                st.markdown("**Embalses que más seguridad energética aportan:**")
                lines = []
                for _, row in top3.iterrows():
                    lines.append(
                        f"- **{row['Embalse']}**: nivel "
                        f"{row['nivel_promedio_pct']:.1f} %, "
                        f"energía máx. ~ {row['energia_max_gwh'] or 0:.0f} GWh."
                    )
                st.markdown("\n".join(lines))

            with st.expander("Ver tabla por embalse y descargar CSV"):
                st.dataframe(df_seg, use_container_width=True, height=280)
                csv = df_seg.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Descargar CSV embalses",
                    data=csv,
                    file_name=f"embalses_{start_date}_{end_date}.csv",
                    mime="text/csv",
                )

    # ------------------------------------------------------------------
    # TAB 3 – APORTES POR RÍO
    # ------------------------------------------------------------------
    with tabs[2]:
        if df_rio.empty:
            st.info(
                "No se encontraron datos de aportes por río para este rango.\n"
                "Revisa la granularidad de 'AporEner' en el catálogo XM."
            )
        else:
            df_rio_tot = (
                df_rio.groupby("Rio")["aportes_gwh"].sum().reset_index()
            )
            df_rio_tot = df_rio_tot.sort_values("aportes_gwh", ascending=False)

            c1, c2 = st.columns(2)
            c1.metric("Ríos / cuencas", len(df_rio_tot))
            c2.metric(
                "Aportes totales en el rango",
                f"{df_rio_tot['aportes_gwh'].sum():,.0f} GWh",
            )

            top_rio = df_rio_tot.head(12).sort_values("aportes_gwh")
            fig_rio = px.bar(
                top_rio,
                x="aportes_gwh",
                y="Rio",
                orientation="h",
                title="Top 12 ríos / cuencas por aportes [GWh]",
                labels={"aportes_gwh": "GWh", "Rio": "Río / cuenca"},
            )
            fig_rio.update_layout(
                template="plotly_dark",
                height=340,
                margin=dict(l=80, r=10, t=40, b=30),
            )
            st.plotly_chart(fig_rio, use_container_width=True)

            with st.expander("Ver tabla de aportes diarios por río y descargar CSV"):
                st.dataframe(df_rio, use_container_width=True, height=280)
                csv_rio = df_rio.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Descargar CSV aportes por río",
                    data=csv_rio,
                    file_name=f"aportes_rio_{start_date}_{end_date}.csv",
                    mime="text/csv",
                )

    # ------------------------------------------------------------------
    # TAB 4 – EMBALSE INDIVIDUAL
    # ------------------------------------------------------------------
    with tabs[3]:
        if df_emb.empty:
            st.info("No hay listado de embalses para seleccionar.")
            return

        embalse_sel = st.selectbox(
            "Selecciona un embalse:",
            df_emb["Embalse"].unique().tolist(),
        )

        if embalse_sel:
            df_e = build_serie_embalse(embalse_sel, start_date, end_date)
            if df_e.empty:
                st.info(
                    "No se encontraron datos de serie temporal para este embalse "
                    "en el rango seleccionado."
                )
            else:
                hist_min, hist_max = get_embalse_hist_extremos(embalse_sel)
                ene_min, ene_max = get_embalse_hist_energia_extremos(embalse_sel)

                c1, c2, c3 = st.columns(3)
                last = df_e.dropna(subset=["nivel_embalse_pct"]).iloc[-1]
                c1.metric(
                    "Nivel actual",
                    f"{last['nivel_embalse_pct']:.1f} %",
                )
                if not pd.isna(hist_min) and not pd.isna(hist_max):
                    c2.metric("Mín. histórico", f"{hist_min:.1f} %")
                    c3.metric("Máx. histórico", f"{hist_max:.1f} %")

                fig_e1 = px.area(
                    df_e,
                    x="Fecha",
                    y="nivel_embalse_pct",
                    title=f"Nivel útil del embalse {embalse_sel}",
                )
                fig_e1.update_traces(
                    line_color=PASTEL_BLUE,
                    fillcolor="rgba(144,202,249,0.35)",
                )
                if not pd.isna(hist_min):
                    fig_e1.add_hline(
                        y=hist_min,
                        line=dict(color="red", width=1, dash="dash"),
                        annotation_text="Mín. histórico",
                        annotation_position="bottom left",
                    )
                if not pd.isna(hist_max):
                    fig_e1.add_hline(
                        y=hist_max,
                        line=dict(color="red", width=1, dash="dash"),
                        annotation_text="Máx. histórico",
                        annotation_position="top left",
                    )
                fig_e1.update_layout(
                    template="plotly_dark",
                    yaxis_title="% volumen útil",
                    height=320,
                    margin=dict(l=20, r=10, t=40, b=30),
                )
                st.plotly_chart(fig_e1, use_container_width=True)

                if "energia_embalsada_gwh" in df_e.columns:
                    fig_e2 = px.line(
                        df_e,
                        x="Fecha",
                        y="energia_embalsada_gwh",
                        title=f"Energía embalsada en {embalse_sel} [GWh]",
                    )
                    fig_e2.update_traces(line_color=PASTEL_GREEN)
                    if not pd.isna(ene_min):
                        fig_e2.add_hline(
                            y=ene_min,
                            line=dict(color="red", width=1, dash="dash"),
                            annotation_text="Energía mín. histórica",
                            annotation_position="bottom left",
                        )
                    if not pd.isna(ene_max):
                        fig_e2.add_hline(
                            y=ene_max,
                            line=dict(color="red", width=1, dash="dash"),
                            annotation_text="Energía máx. histórica",
                            annotation_position="top left",
                        )
                    fig_e2.update_layout(
                        template="plotly_dark",
                        yaxis_title="GWh",
                        height=300,
                        margin=dict(l=20, r=10, t=40, b=30),
                    )
                    st.plotly_chart(fig_e2, use_container_width=True)

                if "delta_embalse_gwh" in df_e.columns:
                    df_bal_e = df_e.dropna(subset=["delta_embalse_gwh"]).copy()
                    colores_e = [
                        "#1B8C3B" if v >= 0 else "#C7253E"
                        for v in df_bal_e["delta_embalse_gwh"]
                    ]
                    fig_e3 = go.Figure()
                    fig_e3.add_trace(
                        go.Bar(
                            x=df_bal_e["Fecha"],
                            y=df_bal_e["delta_embalse_gwh"],
                            marker_color=colores_e,
                            name="Δ energía [GWh/día]",
                        )
                    )
                    fig_e3.update_layout(
                        template="plotly_dark",
                        yaxis_title="GWh/día",
                        xaxis_title="Fecha",
                        height=260,
                        margin=dict(l=20, r=10, t=40, b=30),
                    )
                    fig_e3.update_yaxes(
                        zeroline=True,
                        zerolinewidth=1,
                        zerolinecolor="#FFFFFF",
                    )
                    st.plotly_chart(fig_e3, use_container_width=True)
               
def ui_zni() -> None:
    """
    Vista para Zonas No Interconectadas (ZNI):
      - Usa la tabla de generación enriquecida (E17D25 + catálogo XM)
      - Filtra filas que pertenezcan a ZNI
      - Muestra generación diaria total y por tipo de fuente / tecnología renovable
    """
    st.subheader("Zonas No Interconectadas (ZNI)")
    st.caption("Generación y participación de las ZNI según SIMEM / XM.")

    # ----- RANGO DE FECHAS CON PRESETS -----
    start_date, end_date = select_rango_fechas_presets(
        key_prefix="zni",
        default_years=2,
    )

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    # ----- CARGA DE GENERACIÓN ENRIQUECIDA -----
    with st.spinner("Consultando SIMEM (E17D25) y filtrando ZNI..."):
        df_res = generar_tabla_generacion_enriquecida(start_date, end_date)

    if df_res is None or df_res.empty:
        st.warning("No se obtuvieron datos de generación para el rango seleccionado.")
        return

    # ----- DETECTAR FILAS ZNI -----
    zni_mask = detect_zni_mask(df_res)
    if not zni_mask.any():
        st.warning(
            "No se pudieron identificar filas correspondientes a ZNI.\n"
            "Revisa la función detect_zni_mask() para ajustar la columna o el criterio."
        )
        st.write("Columnas disponibles en el dataset:", list(df_res.columns))
        return

    df_zni = df_res.loc[zni_mask].copy()
    if df_zni.empty:
        st.warning("El subconjunto identificado como ZNI está vacío.")
        return

    # ----- DETECCIÓN DE COLUMNAS BÁSICAS -----
    info_cols = detect_gen_columns(df_zni)
    col_fecha = info_cols["fecha"]
    col_tipo = info_cols["tipo_fuente"]
    col_gen = info_cols["gen"]

    if not col_fecha or not col_tipo or not col_gen:
        st.error(
            "No se pudieron detectar las columnas de fecha, tipo de fuente o "
            "energía dentro de las filas ZNI."
        )
        st.write("Columnas disponibles:", list(df_zni.columns))
        return

    # ----- NORMALIZACIÓN -----
    df_base = df_zni[[col_fecha, col_tipo, col_gen]].copy()
    df_base[col_fecha] = pd.to_datetime(df_base[col_fecha], errors="coerce").dt.date
    df_base[col_gen] = pd.to_numeric(df_base[col_gen], errors="coerce")
    df_base = df_base.dropna(subset=[col_fecha, col_gen])

    if df_base.empty:
        st.warning("Tras limpiar fechas y valores numéricos no quedaron datos ZNI.")
        return

    df_base.rename(
        columns={
            col_fecha: "Fecha",
            col_tipo: "TipoFuente",
            col_gen: "Gen_raw",
        },
        inplace=True,
    )

    # ----- UNIDADES (E17D25) -----
    base_unit_raw = get_dataset_unit("E17D25")
    unit_info = parse_unit_info(base_unit_raw)
    kind = unit_info["kind"]
    base_unit_norm = unit_info["base_unit"]

    if kind == "energy":
        target_unit = "GWh"
    elif kind == "power":
        target_unit = "MW"
    else:
        target_unit = None

    df_base["Gen"] = convert_series_numeric(
        df_base["Gen_raw"],
        base_unit=base_unit_norm,
        target_unit=target_unit,
        kind=kind,
    )
    label_unit = target_unit or base_unit_norm or "unid."

    # ----- CLASIFICADORES LOCALES PARA TIPO Y SUBTECNOLOGÍA -----
    def _zni_clasificar_fuente(txt: str) -> str:
        if not isinstance(txt, str):
            return "Otras"
        t = txt.lower()
        if "hidra" in t:
            return "Hidráulica"
        if any(w in t for w in ["term", "carb", "gas", "diesel", "diésel", "fuel", "mazut", "petcoke"]):
            return "Térmica"
        if any(
            w in t
            for w in [
                "solar",
                "fotovol",
                "eolic",
                "eólico",
                "viento",
                "bioma",
                "biog",
                "fernc",
                "no convencional",
                "menor",
                "pch",
            ]
        ):
            return "Menores y FERNC"
        return "Otras"

    def _zni_clasificar_subtec(txt: str) -> str:
        if not isinstance(txt, str):
            return "Otras"
        t = txt.lower()
        if any(w in t for w in ["solar", "fotovol"]):
            return "Solar"
        if any(w in t for w in ["eolic", "eólico", "viento"]):
            return "Eólica"
        if any(w in t for w in ["bioma", "biog", "residu"]):
            return "Biomasa"
        return "Otras"

    df_base["Categoria"] = df_base["TipoFuente"].astype(str).apply(_zni_clasificar_fuente)
    df_base["SubTec"] = df_base["TipoFuente"].astype(str).apply(_zni_clasificar_subtec)

    # ----- AGREGACIÓN DIARIA POR CATEGORÍA -----
    df_daily = (
        df_base.groupby(["Fecha", "Categoria"], as_index=False)["Gen"]
        .sum()
        .sort_values("Fecha")
    )

    if df_daily.empty:
        st.warning("No hay datos agregados por categoría para ZNI en este rango.")
        return

    # ----- KPIs PRINCIPALES -----
    resumen = df_daily.groupby("Categoria")["Gen"].sum().sort_values(ascending=False)
    total = resumen.sum()

    st.markdown("### Indicadores ZNI en el período seleccionado")

    col_k1, col_k2, col_k3 = st.columns(3)
    col_k1.metric("Energía total ZNI", f"{total:,.1f} {label_unit}")

    energia_renov = float(
        resumen.get("Menores y FERNC", 0.0) + resumen.get("Hidráulica", 0.0)
    )
    share_renov = 100.0 * energia_renov / total if total > 0 else 0.0
    col_k2.metric("Energía renovable (aprox.)", f"{energia_renov:,.1f} {label_unit}")
    col_k3.metric("% renovable sobre ZNI", f"{share_renov:,.1f} %")

    # ----- GRÁFICA: ÁREA APILADA POR CATEGORÍA -----
    st.markdown("### Generación ZNI por tipo de fuente")

    color_map = {
        "Hidráulica": "#64B5F6",
        "Térmica": "#FF8A65",
        "Menores y FERNC": "#81C784",
        "Otras": "#B39DDB",
    }

    fig_zni = px.area(
        df_daily,
        x="Fecha",
        y="Gen",
        color="Categoria",
        title=f"Generación diaria ZNI por tipo de fuente [{label_unit}/día]",
    )
    fig_zni.update_layout(
        template="plotly_dark",
        legend_title_text="Categoría",
        hovermode="x unified",
    )

    for tr in fig_zni.data:
        name = tr.name
        if name in color_map:
            tr.line.color = color_map[name]

    st.plotly_chart(fig_zni, use_container_width=True)

    # ----- DETALLE RENOVABLE: SOLAR / EÓLICA / BIOMASA -----
    st.markdown("### Renovables ZNI por tecnología")

    df_sub = df_base[df_base["SubTec"].isin(["Solar", "Eólica", "Biomasa"])].copy()
    if df_sub.empty:
        st.info("No se identificaron tecnologías Solar / Eólica / Biomasa en ZNI para este rango.")
    else:
        df_sub_daily = (
            df_sub.groupby(["Fecha", "SubTec"], as_index=False)["Gen"]
            .sum()
            .sort_values("Fecha")
        )

        color_sub = {
            "Solar": "#FFD54F",
            "Eólica": "#81D4FA",
            "Biomasa": "#A5D6A7",
        }

        fig_sub = px.area(
            df_sub_daily,
            x="Fecha",
            y="Gen",
            color="SubTec",
            title=f"Generación renovable ZNI por tecnología [{label_unit}/día]",
        )
        fig_sub.update_layout(
            template="plotly_dark",
            legend_title_text="Tecnología",
            hovermode="x unified",
        )
        for tr in fig_sub.data:
            name = tr.name
            if name in color_sub:
                tr.line.color = color_sub[name]

        st.plotly_chart(fig_sub, use_container_width=True)

        resumen_sub = (
            df_sub_daily.groupby("SubTec")["Gen"].sum().sort_values(ascending=False)
        )
        df_sub_share = resumen_sub.reset_index()
        df_sub_share.columns = ["SubTec", "Gen"]

        fig_sub_pie = px.pie(
            df_sub_share,
            names="SubTec",
            values="Gen",
            title=f"Participación renovable ZNI [{label_unit} en el período]",
            hole=0.4,
        )
        fig_sub_pie.update_layout(template="plotly_dark", legend_title_text="Tecnología")
        st.plotly_chart(fig_sub_pie, use_container_width=True)

    # ----- TOP PLANTAS / RECURSOS ZNI -----
    st.markdown("### Top recursos / plantas ZNI por energía generada")

    # Intentamos detectar alguna columna de nombre/planta
    col_planta = info_cols["planta"]
    if col_planta and col_planta in df_zni.columns:
        df_planta = df_zni[[col_planta, col_gen]].copy()
        df_planta[col_gen] = pd.to_numeric(df_planta[col_gen], errors="coerce")
        df_planta = df_planta.dropna(subset=[col_planta, col_gen])

        if not df_planta.empty:
            df_planta_agg = (
                df_planta.groupby(col_planta)[col_gen].sum().reset_index()
            )
            df_planta_agg.rename(
                columns={col_planta: "Recurso", col_gen: "Gen_raw"},
                inplace=True,
            )
            df_planta_agg["Gen"] = convert_series_numeric(
                df_planta_agg["Gen_raw"],
                base_unit=base_unit_norm,
                target_unit=target_unit,
                kind=kind,
            )
            top_rec = df_planta_agg.sort_values("Gen", ascending=False).head(15)

            fig_top = px.bar(
                top_rec.sort_values("Gen"),
                x="Gen",
                y="Recurso",
                orientation="h",
                title=f"Top recursos ZNI por energía generada [{label_unit}]",
                labels={"Gen": label_unit, "Recurso": "Recurso / planta"},
            )
            fig_top.update_layout(template="plotly_dark", height=400, margin=dict(l=80, r=10, t=40, b=30))
            st.plotly_chart(fig_top, use_container_width=True)

            with st.expander("Ver tabla ZNI detallada y descargar CSV"):
                st.dataframe(df_base, use_container_width=True, height=320)
                csv_zni = df_base.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Descargar CSV ZNI (diario por recurso/fuente)",
                    data=csv_zni,
                    file_name=f"zni_generacion_{start_date}_{end_date}.csv",
                    mime="text/csv",
                    key="zni_csv",
                )
        else:
            st.info("No se pudieron construir totales por planta para ZNI.")
    else:
        st.info("No se identificó una columna de nombre de planta/recurso en las filas ZNI.")
                    


# -------------------------------------------------------------------
# UI – FUENTES DE GENERACIÓN
# -------------------------------------------------------------------
def ui_fuentes_generacion() -> None:
    """
    Tablero de Fuentes de generación:
    - Usa la generación detallada SIMEM (E17D25) + catálogo XM.
    - Agrega por día y por categoría (Hidráulica, Térmica, Menores y FERNC, Otras).
    - Además, muestra un detalle renovable por tecnología: Solar, Eólica y Biomasa.
    """

    st.subheader("Fuentes de generación del SIN")

    key_prefix = "fuentes_gen"

    # ----- RANGO DE FECHAS CON PRESETS -----
    start_date, end_date = select_rango_fechas_presets(
        key_prefix=key_prefix,
        default_years=1,
    )

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    # ----- CONSULTA / CARGA DE DATOS -----
    with st.spinner("Consultando SIMEM (E17D25) y catálogo XM..."):
        df_res = generar_tabla_generacion_enriquecida(start_date, end_date)

    if df_res is None or df_res.empty:
        st.warning("No se obtuvieron datos de generación para el rango seleccionado.")
        return

    # ----- DETECCIÓN DE COLUMNAS -----
    info_cols = detect_gen_columns(df_res)
    col_fecha = info_cols["fecha"]
    col_tipo = info_cols["tipo_fuente"]
    col_gen = info_cols["gen"]

    if not col_fecha or not col_tipo or not col_gen:
        st.error(
            "No se pudieron detectar las columnas de fecha, tipo de generación o "
            "energía en el dataset E17D25. Revisa el catálogo SiMEM / XM."
        )
        st.write("Columnas disponibles:", list(df_res.columns))
        return

    # ----- NORMALIZACIÓN DE COLUMNAS -----
    df_base = df_res[[col_fecha, col_tipo, col_gen]].copy()
    df_base[col_fecha] = pd.to_datetime(df_base[col_fecha], errors="coerce").dt.date
    df_base[col_gen] = pd.to_numeric(df_base[col_gen], errors="coerce")
    df_base = df_base.dropna(subset=[col_fecha, col_gen])

    if df_base.empty:
        st.warning("Tras limpiar fechas y valores numéricos no quedaron datos.")
        return

    df_base.rename(
        columns={
            col_fecha: "Fecha",
            col_tipo: "TipoFuente",
            col_gen: "Gen_raw",
        },
        inplace=True,
    )

    # ----- UNIDADES (E17D25) -> normalmente energía -----
    base_unit_raw = get_dataset_unit("E17D25")
    unit_info = parse_unit_info(base_unit_raw)
    kind = unit_info["kind"]
    base_unit_norm = unit_info["base_unit"]

    # Para este tablero queremos, si es energía, trabajar en GWh/día
    if kind == "energy":
        target_unit = "GWh"
    elif kind == "power":
        target_unit = "MW"
    else:
        target_unit = None

    df_base["Gen"] = convert_series_numeric(
        df_base["Gen_raw"],
        base_unit=base_unit_norm,
        target_unit=target_unit,
        kind=kind,
    )

    label_unit = target_unit or base_unit_norm or "unid."

    # ----- CLASIFICACIÓN EN CATEGORÍAS (igual que antes) -----
    def clasificar_fuente(txt: str) -> str:
        if not isinstance(txt, str):
            return "Otras"
        t = txt.lower()

        # Hidráulica
        if "hidra" in t:
            return "Hidráulica"

        # Térmica: gas, carbón, diésel, fuel, etc.
        if any(w in t for w in ["term", "carb", "gas", "diesel", "diésel", "fuel", "mazut", "petcoke"]):
            return "Térmica"

        # Menores y FERNC: solares, eólicas, biomasa, menores, etc.
        if any(
            w in t
            for w in [
                "solar",
                "fotovol",
                "eolic",
                "eólico",
                "viento",
                "bioma",
                "biog",
                "fernc",
                "no convencional",
                "menor",
                "pch",
            ]
        ):
            return "Menores y FERNC"

        return "Otras"

    # ----- CLASIFICACIÓN EN SUBTECNOLOGÍAS (Solar / Eólica / Biomasa) -----
    def clasificar_subtec(txt: str) -> str:
        if not isinstance(txt, str):
            return "Otras"
        t = txt.lower()

        if any(w in t for w in ["solar", "fotovol"]):
            return "Solar"
        if any(w in t for w in ["eolic", "eólico", "viento"]):
            return "Eólica"
        if any(w in t for w in ["bioma", "biog", "residu"]):
            return "Biomasa"

        # Para referencia si quieres ver otras:
        if "hidra" in t:
            return "Hidráulica"
        if any(w in t for w in ["term", "carb", "gas", "diesel", "diésel", "fuel"]):
            return "Térmica"

        return "Otras"

    df_base["Categoria"] = df_base["TipoFuente"].astype(str).apply(clasificar_fuente)
    df_base["SubTec"] = df_base["TipoFuente"].astype(str).apply(clasificar_subtec)

    # ----- AGREGACIÓN DIARIA POR CATEGORÍA -----
    df_daily = (
        df_base.groupby(["Fecha", "Categoria"], as_index=False)["Gen"]
        .sum()
        .sort_values("Fecha")
    )

    if df_daily.empty:
        st.warning("No hay datos agregados por categoría para este rango.")
        return

    # ----- KPIs PRINCIPALES -----
    resumen = df_daily.groupby("Categoria")["Gen"].sum().sort_values(ascending=False)
    total = resumen.sum()

    orden_categorias = ["Hidráulica", "Térmica", "Menores y FERNC", "Otras"]
    categorias_presentes = [c for c in orden_categorias if c in resumen.index]

    if total <= 0:
        st.warning("La energía total del período es cero (o muy pequeña).")
    else:
        cols_kpi = st.columns(len(categorias_presentes))
        for col, cat in zip(cols_kpi, categorias_presentes):
            valor = resumen[cat]
            pct = 100.0 * valor / total if total else 0.0
            col.metric(
                f"{cat} [{label_unit}·día]",
                f"{valor:,.1f}",
                f"{pct:,.1f} % del total",
            )

    # ----- GRÁFICA GLOBAL: ÁREA APILADA POR CATEGORÍA -----
    st.markdown("### Generación diaria por tipo de fuente")

    color_map = {
        "Hidráulica": "#64B5F6",
        "Térmica": "#FF8A65",
        "Menores y FERNC": "#81C784",
        "Otras": "#B39DDB",
    }

    fig_tot = px.area(
        df_daily,
        x="Fecha",
        y="Gen",
        color="Categoria",
        category_orders={"Categoria": orden_categorias},
        title=f"Generación diaria por tipo de fuente [{label_unit}/día]",
    )
    fig_tot.update_layout(
        template="plotly_dark",  # fondo oscuro del visor
        legend_title_text="Categoría",
        hovermode="x unified",
    )

    for tr in fig_tot.data:
        name = tr.name
        if name in color_map:
            tr.line.color = color_map[name]

    st.plotly_chart(fig_tot, use_container_width=True)

    # ----- PEQUEÑAS GRÁFICAS INDIVIDUALES (Hidráulica / Térmica / Menores&FERNC) -----
    st.markdown("### Detalle por categoría")

    col_h, col_t, col_m = st.columns(3)
    pares = [
        ("Hidráulica", col_h),
        ("Térmica", col_t),
        ("Menores y FERNC", col_m),
    ]

    for cat, col in pares:
        with col:
            df_cat = df_daily[df_daily["Categoria"] == cat]
            if df_cat.empty:
                st.info(f"Sin datos para {cat} en el rango seleccionado.")
                continue

            fig_cat = px.area(
                df_cat,
                x="Fecha",
                y="Gen",
                title=cat,
            )
            fig_cat.update_layout(
                template="plotly_dark",
                showlegend=False,
                yaxis_title=f"{label_unit}/día",
                xaxis_title=None,
                margin=dict(l=10, r=10, t=40, b=10),
            )

            if cat in color_map:
                for tr in fig_cat.data:
                    tr.line.color = color_map[cat]

            st.plotly_chart(fig_cat, use_container_width=True)


    # ==================================================================
    # NUEVO GRÁFICO: LAS TRES CATEGORÍAS JUNTAS
    # ==================================================================
    st.markdown("### Comparativo Hidráulica vs Térmica vs Menores y FERNC")

    categorias_principales = ["Hidráulica", "Térmica", "Menores y FERNC"]
    df_3cat = df_daily[df_daily["Categoria"].isin(categorias_principales)].copy()

    if df_3cat.empty:
        st.info("No hay datos para las tres categorías principales en el rango seleccionado.")
    else:
        fig_3cat = px.line(
            df_3cat,
            x="Fecha",
            y="Gen",
            color="Categoria",
            category_orders={"Categoria": categorias_principales},
            title=f"Generación diaria por categoría principal [{label_unit}/día]",
            labels={
                "Gen": f"{label_unit}/día",
                "Categoria": "Categoría",
            },
        )

        # aplicar mismo tema oscuro
        fig_3cat.update_layout(
            template="plotly_dark",
            hovermode="x unified",
        )

        # usar mismo mapa de colores que antes
        for tr in fig_3cat.data:
            name = tr.name
            if name in color_map:
                tr.line.color = color_map[name]

        st.plotly_chart(fig_3cat, use_container_width=True)


    # ----- PARTICIPACIÓN ACUMULADA (PASTEL) -----
    st.markdown("### Participación acumulada en el período")

    df_share = resumen.reset_index()
    df_share.columns = ["Categoria", "Gen"]

    fig_share = px.pie(
        df_share,
        names="Categoria",
        values="Gen",
        title=f"Participación por fuente [{label_unit} en el período]",
        hole=0.4,
    )
    fig_share.update_layout(
        template="plotly_dark",
        legend_title_text="Categoría",
    )
    st.plotly_chart(fig_share, use_container_width=True)

    # ==================================================================
    # 🔥 NUEVO BLOQUE: DETALLE RENOVABLE POR TECNOLOGÍA
    # ==================================================================
    st.markdown("### Renovables por tecnología (Solar, Eólica y Biomasa)")

    df_sub = df_base[df_base["SubTec"].isin(["Solar", "Eólica", "Biomasa"])].copy()

    if df_sub.empty:
        st.info(
            "No se encontraron tecnologías renovables detalladas "
            "(Solar / Eólica / Biomasa) en este rango o dataset."
        )
    else:
        df_sub_daily = (
            df_sub.groupby(["Fecha", "SubTec"], as_index=False)["Gen"]
            .sum()
            .sort_values("Fecha")
        )

        color_sub = {
            "Solar": "#FFD54F",
            "Eólica": "#81D4FA",
            "Biomasa": "#A5D6A7",
        }

        # Área apilada renovable por tecnología
        fig_sub = px.area(
            df_sub_daily,
            x="Fecha",
            y="Gen",
            color="SubTec",
            title=f"Generación renovable diaria por tecnología [{label_unit}/día]",
        )
        fig_sub.update_layout(
            template="plotly_dark",
            legend_title_text="Tecnología",
            hovermode="x unified",
        )
        for tr in fig_sub.data:
            name = tr.name
            if name in color_sub:
                tr.line.color = color_sub[name]

        st.plotly_chart(fig_sub, use_container_width=True)

        # Participación total de cada tecnología
        resumen_sub = (
            df_sub_daily.groupby("SubTec")["Gen"].sum().sort_values(ascending=False)
        )
        df_sub_share = resumen_sub.reset_index()
        df_sub_share.columns = ["SubTec", "Gen"]

        fig_sub_pie = px.pie(
            df_sub_share,
            names="SubTec",
            values="Gen",
            title=f"Participación Solar / Eólica / Biomasa [{label_unit} en el período]",
            hole=0.4,
        )
        fig_sub_pie.update_layout(
            template="plotly_dark",
            legend_title_text="Tecnología",
        )
        st.plotly_chart(fig_sub_pie, use_container_width=True)

    # ----- TABLA Y DESCARGA -----
    with st.expander("Ver tabla diaria por categoría y descargar CSV"):
        st.dataframe(df_daily, use_container_width=True, height=320)
        csv = df_daily.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar CSV (diario por categoría)",
            data=csv,
            file_name=f"generacion_fuentes_{start_date}_{end_date}.csv",
            mime="text/csv",
            key=f"{key_prefix}_csv",
        )


# -------------------------------------------------------------------
# SERIES MACROECONÓMICAS DESDE CSV (TRM, IPC, PIB)
# -------------------------------------------------------------------
@st.cache_data
def load_trm_series(path: Path = Path("TRM.csv")) -> pd.DataFrame:
    """TRM diaria: Fecha, TRM."""
    if not path.exists():
        return pd.DataFrame()

    df = pd.read_csv(path)
    if "DateTime" not in df.columns or len(df.columns) < 2:
        return pd.DataFrame()

    val_col = df.columns[1]  # segunda columna = valor
    df["Fecha"] = pd.to_datetime(df["DateTime"], errors="coerce").dt.date
    df.rename(columns={val_col: "TRM"}, inplace=True)
    df = df[["Fecha", "TRM"]].dropna()
    df["TRM"] = pd.to_numeric(df["TRM"], errors="coerce")
    df = df.dropna(subset=["TRM"]).sort_values("Fecha")
    return df


@st.cache_data
def load_ipc_series(path: Path = Path("IPC.csv")) -> pd.DataFrame:
    """IPC mensual: Periodo (M), IPC."""
    if not path.exists():
        return pd.DataFrame()

    df = pd.read_csv(path)
    if "DateTime" not in df.columns or len(df.columns) < 2:
        return pd.DataFrame()

    val_col = df.columns[1]
    df["Periodo"] = pd.to_datetime(df["DateTime"], errors="coerce").dt.to_period("M")
    df.rename(columns={val_col: "IPC"}, inplace=True)
    df = df[["Periodo", "IPC"]].dropna()
    df["IPC"] = pd.to_numeric(df["IPC"], errors="coerce")
    df = df.dropna(subset=["IPC"]).sort_values("Periodo")
    return df


@st.cache_data
def load_pib_series(path: Path = Path("PIB.csv")) -> pd.DataFrame:
    """PIB real trimestral: Periodo (Q), PIB."""
    if not path.exists():
        return pd.DataFrame()

    df = pd.read_csv(path)
    if "DateTime" not in df.columns or len(df.columns) < 2:
        return pd.DataFrame()

    val_col = df.columns[1]
    df["Periodo"] = pd.to_datetime(df["DateTime"], errors="coerce").dt.to_period("Q")
    df.rename(columns={val_col: "PIB"}, inplace=True)
    df = df[["Periodo", "PIB"]].dropna()
    df["PIB"] = pd.to_numeric(df["PIB"], errors="coerce")
    df = df.dropna(subset=["PIB"]).sort_values("Periodo")
    return df


# -------------------------------------------------------------------
# UI – MERCADO (demanda + TRM, IPC, PIB)
# -------------------------------------------------------------------
def ui_mercado() -> None:
    st.subheader("Mercado y variables macroeconómicas")
    st.caption(
        "Demanda diaria del SIN, días de respaldo hidro y relación con "
        "TRM, IPC, PIB real y Precio de Bolsa frente a la generación por fuente."
    )

    hoy = dt.date.today()
    default_start = hoy - dt.timedelta(days=365 * 5)

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Fecha inicio",
            value=default_start,
            format="YYYY-MM-DD",
            key="mercado_start",
        )
    with col2:
        end_date = st.date_input(
            "Fecha fin",
            value=hoy,
            format="YYYY-MM-DD",
            key="mercado_end",
        )

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    # ----------------------- series XM (demanda / respaldo) -----------------------
    with st.spinner("Consultando series diarias de demanda y respaldo..."):
        df = build_hidro_diaria(start_date, end_date)

    if df.empty or "demanda_gwh" not in df.columns:
        st.info("No se pudo obtener la demanda diaria DemaSIN en este rango.")
        return

    df = df.sort_values("Fecha").copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"])

    # ----------------------- series macroeconómicas desde CSV -----------------------
    trm = load_trm_series()
    ipc = load_ipc_series()
    pib = load_pib_series()

    tab_pbe, tab_trm, tab_ipc, tab_pib = st.tabs(
        [
            "Precio Bolsa y generación",
            "TRM y demanda",
            "IPC mensual",
            "PIB trimestral",
        ]
    )

    # ======================================================================
    # 1) PRECIO DE BOLSA (XM) vs GENERACIÓN POR FUENTE (XM/SIMEM)
    # ======================================================================
    with tab_pbe:
        st.markdown("### Precio de Bolsa y generación por tipo de fuente")

        # -------- Precio de Bolsa diario --------
        df_pbe = build_precio_bolsa_diario(start_date, end_date)
        if df_pbe.empty:
            st.info("No se pudo obtener el Precio de Bolsa para este rango.")
            return

        df_pbe["Fecha"] = pd.to_datetime(df_pbe["Fecha"])

        # -------- Generación por fuente desde SIMEM E17D25 --------
        with st.spinner("Consultando y agregando generación por fuente desde SIMEM..."):
            df_gen = generar_tabla_generacion_enriquecida(start_date, end_date)

        if df_gen.empty:
            st.info("No se pudo obtener la generación detallada desde SIMEM en este rango.")
            return

        info_cols = detect_gen_columns(df_gen)
        col_fecha = info_cols["fecha"]
        col_tipo = info_cols["tipo_fuente"]
        col_gen = info_cols["gen"]

        if not col_fecha or not col_tipo or not col_gen:
            st.error(
                "No se pudieron detectar las columnas de fecha, tipo de fuente "
                "y generación en el dataset E17D25."
            )
            st.write("Columnas disponibles:", list(df_gen.columns))
            return

        df_g = df_gen[[col_fecha, col_tipo, col_gen]].copy()
        df_g[col_fecha] = pd.to_datetime(df_g[col_fecha], errors="coerce").dt.date
        df_g[col_gen] = pd.to_numeric(df_g[col_gen], errors="coerce")
        df_g = df_g.dropna(subset=[col_fecha, col_gen])

        if df_g.empty:
            st.info("No hay datos numéricos de generación para el período seleccionado.")
            return

        df_g.rename(
            columns={
                col_fecha: "Fecha",
                col_tipo: "TipoFuente",
                col_gen: "Gen_raw",
            },
            inplace=True,
        )

        # -------- Unidades de E17D25 (convertimos a GWh/día si aplica) --------
        base_unit_raw = get_dataset_unit("E17D25")
        unit_info = parse_unit_info(base_unit_raw)
        kind = unit_info["kind"]
        base_unit_norm = unit_info["base_unit"]

        if kind == "energy":
            target_unit = "GWh"
        elif kind == "power":
            target_unit = "MW"
        else:
            target_unit = None

        df_g["Gen"] = convert_series_numeric(
            df_g["Gen_raw"],
            base_unit=base_unit_norm,
            target_unit=target_unit,
            kind=kind,
        )
        label_unit = target_unit or base_unit_norm or "unid."

        # -------- Clasificación de fuente en categorías --------
        def clasificar_fuente(txt: str) -> str:
            if not isinstance(txt, str):
                return "Otras"
            t = txt.lower()

            # Hidráulica
            if "hidra" in t:
                return "Hidráulica"

            # Térmica: gas, carbón, diésel, fuel, etc.
            if any(w in t for w in ["term", "carb", "gas", "diesel", "diésel", "fuel", "mazut", "petcoke"]):
                return "Térmica"

            # Menores y FERNC: solares, eólicas, biomasa, menores, etc.
            if any(
                w in t
                for w in [
                    "solar",
                    "fotovol",
                    "eolic",
                    "eólico",
                    "viento",
                    "bioma",
                    "biog",
                    "fernc",
                    "no convencional",
                    "menor",
                    "pch",
                ]
            ):
                return "Menores y FERNC"

            return "Otras"

        df_g["Categoria"] = df_g["TipoFuente"].astype(str).apply(clasificar_fuente)

        # -------- Agregación diaria por categoría --------
        df_daily_gen = (
            df_g.groupby(["Fecha", "Categoria"], as_index=False)["Gen"]
            .sum()
            .sort_values("Fecha")
        )
        df_daily_gen["Fecha"] = pd.to_datetime(df_daily_gen["Fecha"])

        if df_daily_gen.empty:
            st.info("No hay datos agregados por categoría para este rango.")
            return

        # Pivot a formato ancho: una columna por categoría
        df_gen_wide = df_daily_gen.pivot(
            index="Fecha",
            columns="Categoria",
            values="Gen",
        ).reset_index()
        df_gen_wide.columns.name = None  # quitar nombre del índice de columnas

        # -------- Unimos Precio de Bolsa + generación por categoría --------
        df_pb_gen = pd.merge(df_pbe, df_gen_wide, on="Fecha", how="inner")

        if df_pb_gen.empty:
            st.info(
                "No hay fechas comunes entre Precio de Bolsa y la generación por fuente "
                "en el rango seleccionado."
            )
            return

        # Generación total (suma de todas las categorías)
        cat_all = ["Hidráulica", "Térmica", "Menores y FERNC", "Otras"]
        gen_cols = [c for c in cat_all if c in df_pb_gen.columns]
        df_pb_gen["Gen_total"] = df_pb_gen[gen_cols].sum(axis=1)

        c1, c2 = st.columns(2)
        c1.metric(
            "Precio de Bolsa (último día)",
            f"{df_pb_gen['precio_bolsa'].iloc[-1]:,.0f}",
        )
        c2.metric(
            "Generación total (último día)",
            f"{df_pb_gen['Gen_total'].iloc[-1]:,.1f} {label_unit}/día",
        )

        # -------- Colores para categorías --------
        color_map = {
            "Hidráulica": "#64B5F6",
            "Térmica": "#FF8A65",
            "Menores y FERNC": "#81C784",
            "Otras": "#B39DDB",
        }

        # -------- Serie temporal: Precio de Bolsa vs generación total --------
        fig_ts = go.Figure()
        fig_ts.add_trace(
            go.Scatter(
                x=df_pb_gen["Fecha"],
                y=df_pb_gen["Gen_total"],
                name=f"Generación total [{label_unit}/día]",
                line=dict(color=PASTEL_GREEN),
                yaxis="y1",
            )
        )
        fig_ts.add_trace(
            go.Scatter(
                x=df_pb_gen["Fecha"],
                y=df_pb_gen["precio_bolsa"],
                name="Precio de Bolsa",
                line=dict(color=PASTEL_RED),
                yaxis="y2",
            )
        )
        fig_ts.update_layout(
            template="plotly_dark",
            title="Generación total diaria vs Precio de Bolsa de Energía",
            xaxis_title="Fecha",
            yaxis=dict(title=f"Generación total [{label_unit}/día]"),
            yaxis2=dict(
                title="Precio de Bolsa",
                overlaying="y",
                side="right",
            ),
            hovermode="x unified",
        )
        st.plotly_chart(fig_ts, use_container_width=True)

        # -------- Serie temporal: Precio de Bolsa vs generación por tipo --------
        st.markdown("#### Generación diaria por tipo de fuente vs Precio de Bolsa")

        fig_cat_ts = go.Figure()

        for cat in gen_cols:
            fig_cat_ts.add_trace(
                go.Scatter(
                    x=df_pb_gen["Fecha"],
                    y=df_pb_gen[cat],
                    name=f"Gen {cat} [{label_unit}/día]",
                    yaxis="y1",
                    line=dict(color=color_map.get(cat, None)),
                )
            )

        fig_cat_ts.add_trace(
            go.Scatter(
                x=df_pb_gen["Fecha"],
                y=df_pb_gen["precio_bolsa"],
                name="Precio de Bolsa",
                yaxis="y2",
                line=dict(color=PASTEL_RED, dash="dash"),
            )
        )

        fig_cat_ts.update_layout(
            template="plotly_dark",
            title="Generación diaria por tipo de fuente y Precio de Bolsa",
            xaxis_title="Fecha",
            yaxis=dict(title=f"Generación [{label_unit}/día]"),
            yaxis2=dict(
                title="Precio de Bolsa",
                overlaying="y",
                side="right",
            ),
            hovermode="x unified",
        )
        st.plotly_chart(fig_cat_ts, use_container_width=True)

        # -------- Correlación Precio de Bolsa vs cada tipo de generación --------
        st.markdown("#### Correlación Precio de Bolsa – Generación diaria por tipo de fuente")

        corrs = []
        for cat in gen_cols:
            serie_gen = df_pb_gen[cat]
            if serie_gen.notna().sum() > 10:
                r = df_pb_gen["precio_bolsa"].corr(serie_gen)
                if pd.notna(r):
                    corrs.append(
                        {
                            "Categoria": cat,
                            "Correlación Precio–Generación": r,
                        }
                    )

        if not corrs:
            st.info(
                "No hay suficientes datos para calcular correlaciones por categoría."
            )
        else:
            df_corr = pd.DataFrame(corrs).sort_values(
                "Correlación Precio–Generación", ascending=False
            )

            st.dataframe(df_corr, use_container_width=True)

            fig_corr = px.bar(
                df_corr,
                x="Categoria",
                y="Correlación Precio–Generación",
                title="Correlación Precio de Bolsa – Generación diaria por tipo de fuente",
                labels={"Categoria": "Tipo de fuente"},
            )
            fig_corr.update_layout(
                template="plotly_dark",
                yaxis_range=[-1, 1],
            )
            st.plotly_chart(fig_corr, use_container_width=True)

        # -------- Diagramas de dispersión Precio vs generación por tipo --------
        st.markdown("#### Diagramas de dispersión Precio de Bolsa vs generación diaria por tipo de fuente")

        if len(gen_cols) == 0:
            st.info("No hay columnas de generación por categoría para graficar la dispersión.")
        else:
            subplot_titles = [f"Precio vs {cat}" for cat in gen_cols]

            fig_sc_all = make_subplots(
                rows=1,
                cols=len(gen_cols),
                shared_yaxes=True,
                subplot_titles=subplot_titles,
                horizontal_spacing=0.04,
            )

            for i, cat in enumerate(gen_cols, start=1):
                fig_sc_all.add_trace(
                    go.Scatter(
                        x=df_pb_gen[cat],
                        y=df_pb_gen["precio_bolsa"],
                        mode="markers",
                        name=cat,
                        marker=dict(size=5),
                        showlegend=False,
                    ),
                    row=1,
                    col=i,
                )
                fig_sc_all.update_xaxes(
                    title_text=f"{cat} [{label_unit}/día]",
                    row=1,
                    col=i,
                )

            fig_sc_all.update_yaxes(
                title_text="Precio de Bolsa",
                row=1,
                col=1,
            )
            fig_sc_all.update_layout(
                title_text="Dispersión Precio de Bolsa vs generación diaria por tipo de fuente",
                template="plotly_dark",
                height=450,
            )
            st.plotly_chart(fig_sc_all, use_container_width=True)

        # -------- Tabla y descarga --------
        with st.expander("Ver tabla Precio de Bolsa + generación y descargar CSV"):
            st.dataframe(df_pb_gen, use_container_width=True, height=350)
            csv_pb = df_pb_gen.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Descargar CSV Precio de Bolsa + generación por tipo",
                data=csv_pb,
                file_name=f"precio_bolsa_generacion_{start_date}_{end_date}.csv",
                mime="text/csv",
            )

    # ======================================================================
    # 2) TRM Y DEMANDA (diario)
    # ======================================================================
    with tab_trm:
        if trm.empty:
            st.info("No se encontró TRM.csv en la carpeta del proyecto.")
        else:
            # Filtrar TRM al rango seleccionado
            trm_f = trm[
                (trm["Fecha"] >= start_date) & (trm["Fecha"] <= end_date)
            ].copy()
            if trm_f.empty:
                st.info("No hay datos de TRM en el rango seleccionado.")
            else:
                df_dem = df[["Fecha", "demanda_gwh"]].dropna().copy()
                trm_f["Fecha"] = pd.to_datetime(trm_f["Fecha"])
                df_trm_join = pd.merge(df_dem, trm_f, on="Fecha", how="inner")
                if df_trm_join.empty:
                    st.info("No hay fechas comunes entre demanda y TRM.")
                else:
                    corr = df_trm_join["demanda_gwh"].corr(df_trm_join["TRM"])
                    last = df_trm_join.iloc[-1]

                    c1, c2, c3 = st.columns(3)
                    c1.metric(
                        "TRM última fecha",
                        f"{last['TRM']:,.0f} $/USD",
                    )
                    c2.metric(
                        "Demanda misma fecha",
                        f"{last['demanda_gwh']:.1f} GWh/día",
                    )
                    c3.metric(
                        "Correlación demanda–TRM",
                        f"{corr:.2f}" if not np.isnan(corr) else "N/A",
                    )

                    # Serie temporal conjunta (dos ejes)
                    fig_trm = go.Figure()
                    fig_trm.add_trace(
                        go.Scatter(
                            x=df_trm_join["Fecha"],
                            y=df_trm_join["demanda_gwh"],
                            name="Demanda [GWh/día]",
                            line=dict(color=PASTEL_YELLOW),
                            yaxis="y1",
                        )
                    )
                    fig_trm.add_trace(
                        go.Scatter(
                            x=df_trm_join["Fecha"],
                            y=df_trm_join["TRM"],
                            name="TRM [COP/USD]",
                            line=dict(color=PASTEL_BLUE),
                            yaxis="y2",
                        )
                    )
                    fig_trm.update_layout(
                        template="plotly_dark",
                        title="Demanda diaria del SIN vs TRM",
                        xaxis_title="Fecha",
                        yaxis=dict(title="Demanda [GWh/día]"),
                        yaxis2=dict(
                            title="TRM [COP/USD]",
                            overlaying="y",
                            side="right",
                        ),
                        hovermode="x unified",
                    )
                    st.plotly_chart(fig_trm, use_container_width=True)

                    # Dispersión demanda vs TRM
                    fig_trm_sc = px.scatter(
                        df_trm_join,
                        x="TRM",
                        y="demanda_gwh",
                        labels={
                            "TRM": "TRM [COP/USD]",
                            "demanda_gwh": "Demanda [GWh/día]",
                        },
                        title="Dispersión demanda diaria vs TRM",
                    )
                    fig_trm_sc.update_layout(template="plotly_dark")
                    st.plotly_chart(fig_trm_sc, use_container_width=True)

                    with st.expander("Ver datos demanda–TRM y descargar CSV"):
                        st.dataframe(df_trm_join, use_container_width=True, height=350)
                        csv_trm = df_trm_join.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "Descargar CSV demanda–TRM",
                            data=csv_trm,
                            file_name=f"demanda_trm_{start_date}_{end_date}.csv",
                            mime="text/csv",
                        )

    # ======================================================================
    # 3) IPC MENSUAL vs DEMANDA
    # ======================================================================
    with tab_ipc:
        if ipc.empty:
            st.info("No se encontró IPC.csv en la carpeta del proyecto.")
        else:
            df_m = df.copy()
            df_m["Periodo"] = df_m["Fecha"].dt.to_period("M")
            df_m = (
                df_m.groupby("Periodo", as_index=False)
                .agg(
                    demanda_gwh=("demanda_gwh", "mean"),
                    dias_respaldo=("dias_respaldo", "mean"),
                )
            )
            if df_m.empty:
                st.info("No se pudo construir serie mensual de demanda.")
            else:
                p_min = df_m["Periodo"].min()
                p_max = df_m["Periodo"].max()
                ipc_f = ipc[
                    (ipc["Periodo"] >= p_min) & (ipc["Periodo"] <= p_max)
                ].copy()
                if ipc_f.empty:
                    st.info("No hay datos de IPC para el período analizado.")
                else:
                    df_ipc = pd.merge(df_m, ipc_f, on="Periodo", how="inner")
                    df_ipc["Fecha"] = df_ipc["Periodo"].dt.to_timestamp("M")

                    corr_ipc_dem = df_ipc["demanda_gwh"].corr(df_ipc["IPC"])
                    corr_ipc_res = df_ipc["dias_respaldo"].corr(df_ipc["IPC"])

                    last = df_ipc.iloc[-1]
                    c1, c2, c3 = st.columns(3)
                    c1.metric("IPC último dato", f"{last['IPC']:.2f}")
                    c2.metric(
                        "Demanda media mes último",
                        f"{last['demanda_gwh']:.1f} GWh/día",
                    )
                    c3.metric(
                        "Corr. IPC–demanda (mensual)",
                        f"{corr_ipc_dem:.2f}" if not np.isnan(corr_ipc_dem) else "N/A",
                    )
                    st.caption(
                        f"Correlación IPC–días de respaldo: "
                        f"{corr_ipc_res:.2f}" if not np.isnan(corr_ipc_res) else
                        "Correlación IPC–días de respaldo: N/A"
                    )

                    # Serie temporal conjunta
                    fig_ipc = go.Figure()
                    fig_ipc.add_trace(
                        go.Scatter(
                            x=df_ipc["Fecha"],
                            y=df_ipc["demanda_gwh"],
                            name="Demanda media mensual [GWh/día]",
                            line=dict(color=PASTEL_YELLOW),
                            yaxis="y1",
                        )
                    )
                    fig_ipc.add_trace(
                        go.Scatter(
                            x=df_ipc["Fecha"],
                            y=df_ipc["IPC"],
                            name="IPC (índice)",
                            line=dict(color=PASTEL_BLUE),
                            yaxis="y2",
                        )
                    )
                    fig_ipc.update_layout(
                        template="plotly_dark",
                        title="Demanda media mensual vs IPC",
                        xaxis_title="Fecha",
                        yaxis=dict(title="Demanda [GWh/día]"),
                        yaxis2=dict(
                            title="IPC (índice)",
                            overlaying="y",
                            side="right",
                        ),
                        hovermode="x unified",
                    )
                    st.plotly_chart(fig_ipc, use_container_width=True)

                    # Dispersión
                    fig_ipc_sc = px.scatter(
                        df_ipc,
                        x="IPC",
                        y="demanda_gwh",
                        labels={
                            "IPC": "IPC (índice)",
                            "demanda_gwh": "Demanda media mensual [GWh/día]",
                        },
                        title="Demanda mensual vs IPC",
                    )
                    fig_ipc_sc.update_layout(template="plotly_dark")
                    st.plotly_chart(fig_ipc_sc, use_container_width=True)

                    with st.expander("Ver datos mensuales y descargar CSV"):
                        st.dataframe(df_ipc, use_container_width=True, height=350)
                        csv_ipc = df_ipc.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "Descargar CSV demanda–IPC (mensual)",
                            data=csv_ipc,
                            file_name=f"demanda_ipc_mensual_{start_date}_{end_date}.csv",
                            mime="text/csv",
                        )

    # ======================================================================
    # 4) PIB TRIMESTRAL vs DEMANDA
    # ======================================================================
    with tab_pib:
        if pib.empty:
            st.info("No se encontró PIB.csv en la carpeta del proyecto.")
        else:
            df_q = df.copy()
            df_q["Periodo"] = df_q["Fecha"].dt.to_period("Q")
            df_q = (
                df_q.groupby("Periodo", as_index=False)
                .agg(
                    demanda_gwh=("demanda_gwh", "mean"),
                    dias_respaldo=("dias_respaldo", "mean"),
                )
            )
            if df_q.empty:
                st.info("No se pudo construir serie trimestral de demanda.")
            else:
                p_min = df_q["Periodo"].min()
                p_max = df_q["Periodo"].max()
                pib_f = pib[
                    (pib["Periodo"] >= p_min) & (pib["Periodo"] <= p_max)
                ].copy()
                if pib_f.empty:
                    st.info("No hay datos de PIB real para el período analizado.")
                else:
                    df_pib = pd.merge(df_q, pib_f, on="Periodo", how="inner")
                    df_pib["Fecha"] = df_pib["Periodo"].dt.to_timestamp("Q")

                    corr_pib_dem = df_pib["demanda_gwh"].corr(df_pib["PIB"])
                    corr_pib_res = df_pib["dias_respaldo"].corr(df_pib["PIB"])

                    last = df_pib.iloc[-1]
                    c1, c2, c3 = st.columns(3)
                    c1.metric("PIB real último dato", f"{last['PIB']:,.2f}")
                    c2.metric(
                        "Demanda media trimestre último",
                        f"{last['demanda_gwh']:.1f} GWh/día",
                    )
                    c3.metric(
                        "Corr. PIB–demanda (trimestral)",
                        f"{corr_pib_dem:.2f}" if not np.isnan(corr_pib_dem) else "N/A",
                    )
                    st.caption(
                        f"Correlación PIB–días de respaldo: "
                        f"{corr_pib_res:.2f}" if not np.isnan(corr_pib_res) else
                        "Correlación PIB–días de respaldo: N/A"
                    )

                    # Serie temporal conjunta
                    fig_pib = go.Figure()
                    fig_pib.add_trace(
                        go.Scatter(
                            x=df_pib["Fecha"],
                            y=df_pib["demanda_gwh"],
                            name="Demanda media trimestral [GWh/día]",
                            line=dict(color=PASTEL_YELLOW),
                            yaxis="y1",
                        )
                    )
                    fig_pib.add_trace(
                        go.Scatter(
                            x=df_pib["Fecha"],
                            y=df_pib["PIB"],
                            name="PIB real (índice, base 2015)",
                            line=dict(color=PASTEL_BLUE),
                            yaxis="y2",
                        )
                    )
                    fig_pib.update_layout(
                        template="plotly_dark",
                        title="Demanda eléctrica vs PIB real (trimestral)",
                        xaxis_title="Fecha",
                        yaxis=dict(title="Demanda [GWh/día]"),
                        yaxis2=dict(
                            title="PIB real (índice, base 2015)",
                            overlaying="y",
                            side="right",
                        ),
                        hovermode="x unified",
                    )
                    st.plotly_chart(fig_pib, use_container_width=True)

                    fig_pib_sc = px.scatter(
                        df_pib,
                        x="PIB",
                        y="demanda_gwh",
                        labels={
                            "PIB": "PIB real (índice, base 2015)",
                            "demanda_gwh": "Demanda media trimestral [GWh/día]",
                        },
                        title="Demanda trimestral vs PIB real",
                    )
                    fig_pib_sc.update_layout(template="plotly_dark")
                    st.plotly_chart(fig_pib_sc, use_container_width=True)

                    with st.expander("Ver datos trimestrales y descargar CSV"):
                        st.dataframe(df_pib, use_container_width=True, height=350)
                        csv_pib = df_pib.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "Descargar CSV demanda–PIB (trimestral)",
                            data=csv_pib,
                            file_name=f"demanda_pib_trimestral_{start_date}_{end_date}.csv",
                            mime="text/csv",
                        )

    
        
def forecast_serie_fecha(
    df: pd.DataFrame,
    col: str,
    horizon_days: int,
    use_log: bool = True,
) -> Tuple[pd.DataFrame, pd.DataFrame, str]:
    """
    Pronostica una serie diaria con columna 'Fecha' y valor en `col`.

    Devuelve:
        df_hist_out:  Fecha, col  (histórico usado)
        df_pred_out:  Fecha, col  (pronóstico futuro)
        modelo_desc:  descripción corta del modelo

    Modelo principal:
        - Regresión Ridge con:
            t, t^2, senos/cosenos anuales, dummies de día de semana y mes.
        - Opcionalmente transforma el target con log(1 + y) para series siempre positivas.
    """

    # ------------------ LIMPIEZA Y RE-INDEXACIÓN DIARIA ------------------
    if "Fecha" not in df.columns:
        raise ValueError("El DataFrame debe tener una columna 'Fecha'.")

    df_local = df[["Fecha", col]].copy()
    df_local["Fecha"] = pd.to_datetime(df_local["Fecha"], errors="coerce")
    df_local = df_local.dropna(subset=["Fecha", col]).sort_values("Fecha")

    if df_local.empty:
        return pd.DataFrame(), pd.DataFrame(), "Sin datos"

    # Índice diario completo
    idx_hist = pd.date_range(
        df_local["Fecha"].min(),
        df_local["Fecha"].max(),
        freq="D",
    )
    df_ts = df_local.set_index("Fecha").reindex(idx_hist)
    df_ts.index.name = "Fecha"

    # Interpolar huecos y rellenar extremos
    df_ts[col] = (
        df_ts[col]
        .interpolate(method="time")
        .fillna(method="bfill")
        .fillna(method="ffill")
    )

    # Por si queda algo raro
    df_ts[col] = pd.to_numeric(df_ts[col], errors="coerce")
    df_ts = df_ts.dropna(subset=[col])

    if len(df_ts) < 60:
        # Muy pocos datos → pronóstico burdo
        mean_val = float(df_ts[col].tail(30).mean())
        future_idx = pd.date_range(
            df_ts.index[-1] + pd.Timedelta(days=1),
            periods=horizon_days,
            freq="D",
        )
        df_hist_out = df_ts.reset_index()[["Fecha", col]]
        df_pred_out = pd.DataFrame({"Fecha": future_idx, [col]: mean_val})
        return df_hist_out, df_pred_out, "Promedio móvil (datos insuficientes)"
        # ------------------ FEATURES Y TRANSFORMACIÓN ------------------
    y = df_ts[col].astype(float).values

    if use_log:
        y_clipped = np.clip(y, a_min=1e-6, a_max=None)
        y_tr = np.log1p(y_clipped)
    else:
        y_tr = y

    # t empieza en 0 para el histórico
    X_hist = build_time_features(df_ts.index, t_start=0)

    # ------------------ INTENTAR RIDGE (scikit-learn) ------------------
    try:
        from sklearn.linear_model import Ridge

        model = Ridge(alpha=1.0)
        model.fit(X_hist.values, y_tr)

        modelo_desc = "Ridge (tendencia + estacionalidad anual)"
        use_baseline = False

    except Exception:
        use_baseline = True
        modelo_desc = "Promedio móvil 30 días (sin scikit-learn)"

    # ------------------ PRONÓSTICO FUTURO ------------------
    future_idx = pd.date_range(
        df_ts.index[-1] + pd.Timedelta(days=1),
        periods=horizon_days,
        freq="D",
    )

    if use_baseline:
        mean_val = float(df_ts[col].tail(30).mean())
        y_future = np.full(shape=len(future_idx), fill_value=mean_val, dtype=float)
    else:
        # t continúa a partir del último valor del histórico
        last_t = float(X_hist["t"].iloc[-1])
        X_future = build_time_features(
            future_idx,
            t_start=int(last_t) + 1,
        )

        # aseguramos mismas columnas entre hist y futuro
        feat_cols = X_hist.columns
        for c in feat_cols:
            if c not in X_future.columns:
                X_future[c] = 0.0
        X_future = X_future[feat_cols]

        y_hat_future = model.predict(X_future.values)

        if use_log:
            y_future = np.expm1(y_hat_future)
        else:
            y_future = y_hat_future

        y_future = np.maximum(y_future, 0.0)
    # ------------------ SALIDAS ------------------
    df_hist_out = df_ts.reset_index()[["Fecha", col]]
    df_pred_out = pd.DataFrame(
        {
            "Fecha": future_idx,
            col: y_future,
        }
    )

    return df_hist_out, df_pred_out, modelo_desc
   
# -------------------------------------------------------------------
# HELPERS PARA PREDICCIONES
# -------------------------------------------------------------------
import numpy as np  # ⚠️ asegúrate de tener ESTO arriba del archivo
def go_to(section: str) -> None:
    """Cambia la sección activa y recarga la app."""
    st.session_state["seccion"] = section
    try:
        st.rerun()          # Streamlit >= 1.27
    except Exception:
        st.experimental_rerun()   # compatibilidad con versiones viejas

def _normalize_embalse_name(name: str) -> str:
    """
    Normaliza nombres de embalses para poder empatar entre XM y nuestra tabla:
    - Quita tildes
    - Pone todo en mayúsculas
    - Quita espacios, guiones, etc.
    """
    if not isinstance(name, str):
        return ""
    # quitar tildes
    nfkd = unicodedata.normalize("NFKD", name)
    s = "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    # mayúsculas y sin separadores
    s = s.upper()
    for ch in [" ", "-", "_", ".", ","]:
        s = s.replace(ch, "")
    return s
# -------------------------------------------------------------------
# META – EMBALSES PRINCIPALES CONECTADOS AL SIN
EMBALSES_SIN_META = pd.DataFrame(
    [
        # Nombre_Visor,  Lat,    Lon,      Río / sistema                                    , Cuenca             , Hidro_asociada                              , Operador                                   , Región_hidrológica, Estación_hidrológica
        ("PENOL",         6.242, -75.169, "Río Nare (Embalse Peñol-Guatapé)"               , "Magdalena–Cauca" , "Guatapé"                                   , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "NARE"),
        ("AGREGADO BOGOTA",4.800,-73.950, "Río Bogotá (Tominé–Sisga–Neusa–Muña)"           , "Magdalena"       , "Pagua (Paraíso–Guaca)"                     , "ENEL COLOMBIA SA ESP"                    , "Centro"          , "NEUSA / SISGA / TOMINÉ"),
        ("GUAVIO",        4.745, -73.479, "Río Guavio"                                     , "Orinoco"         , "Guavio"                                    , "ENEL COLOMBIA SA ESP"                    , "Oriente"         , "GUAVIO / BATATAS / CHIVOR / GUAVIO2 / DESVIACIÓN RÍO BATATAS / DESVIACIÓN RÍO CHIVOR"),
        ("TOPOCORO",      6.823, -73.152, "Río Sogamoso (Embalse Topocoro)"                , "Magdalena"       , "Sogamosos (Hidrosogamoso)"                 , "ISAGEN S.A. E.S.P."                      , "Centro"          , "SOGAMOSO"),
        ("EL QUIMBO",     2.250, -75.700, "Río Magdalena"                                  , "Magdalena"       , "El Quimbo"                                 , "ENEL COLOMBIA SA ESP"                    , "Centro"          , "EL QUIMBO"),

        # *** CORREGIDO ***
        ("ESMERALDA",     4.950, -73.327, "Río Garagoa/Batá (Valle de Tenza)"              , "Orinoco"         , "Chivor"                                    , "AES COLOMBIA & CIA. S.C.A. E.S.P."       , "Oriente"         , "BATÁ / DESVIACIÓN RUCIO + NEGRO / DESVIACIÓN RÍO TUNITA / RUCIO / TUNITA / NEGRO"),
        ("CHUZA",         4.640, -73.720, "Río Chuza (afluente del Guatiquía)"             , "Orinoco"         , "Pagua"                                     , "ENEL COLOMBIA SA ESP"                    , "Centro"          , "BOGOTÁ / MUÑA"),
        ("RIOGRANDE2",    6.580, -75.557, "Río Grande"                                     , "Magdalena–Cauca" , "La Tasajera"                               , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "RÍO GRANDE"),

        # *** AJUSTE río ***
        ("SAN LORENZO",   6.370, -75.040, "Ríos Nare y San Lorenzo"                        , "Magdalena–Cauca" , "Jaguas"                                    , "ISAGEN S.A. E.S.P."                      , "Antioquia"       , "SAN LORENZO"),
        ("SALVAJINA",     3.030, -76.685, "Río Cauca"                                      , "Magdalena–Cauca" , "Salvajina"                                 , "CELSIA COLOMBIA S.A. E.S.P."             , "Valle"           , "SALVAJINA"),
        ("AMANI",         5.525, -74.880, "Río La Miel (Embalse Amaní)"                    , "Magdalena–Cauca" , "Miel I"                                    , "ISAGEN S.A. E.S.P."                      , "Caldas"          , "GUARINÓ / MIEL / MANSO / DESVIACIÓN GUARINÓ / DESVIACIÓN MANSO"),
        ("CALIMA1",       3.907, -76.554, "Río Calima"                                     , "Pacífico (río San Juan)", "Calima"                      , "CELSIA COLOMBIA S.A. E.S.P."             , "Valle"           , "ALTOANCHICAYÁ / CALIMA / BRAVO / DESVIACIÓN RÍO BRAVO"),

        # *** AJUSTADO ***
        ("MIRAFLORES",    6.767, -75.318, "Río Tenche (complejo río Guadalupe)"            , "Magdalena–Cauca" , "Guatrón (Guadalupe–Troneras)"             , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "TENCHE"),
        ("URRA1",         8.026, -76.189, "Río Sinú"                                       , "Caribe"          , "Urrá I"                                    , "EMPRESA URRA S.A. E.S.P."                , "Caribe"          , "URRÁ"),
        ("PLAYAS",        6.166, -75.044, "Ríos Guatapé y Nare (descargas Guatapé y Jaguas)", "Magdalena–Cauca", "Playas"                      , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "GUATAPÉ"),
        ("BETANIA",       2.515, -75.469, "Río Magdalena (sector Yaguará)"                 , "Magdalena"       , "Betania"                                  , "ENEL COLOMBIA SA ESP"                    , "Centro"          , "BETANIA CP"),
        ("PORCE II",      6.980, -75.090, "Río Porce"                                      , "Magdalena–Cauca" , "Porce II"                                 , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "PORCE2 CP"),

        # *** CORREGIDO ***
        ("PORCE III",     6.939, -75.139, "Río Porce"                                      , "Magdalena–Cauca" , "Porce III"                                , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "PORCE 3"),

        # *** AJUSTADO ***
        ("TRONERAS",      6.751, -75.254, "Río Concepción y río Guadalupe (Miraflores→Troneras)",
                                                                                           "Magdalena–Cauca" , "Guatrón (Troneras/Guadalupe III)"        , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "DOLORES / PAJARITO / CONCEPCIÓN / NECHÍ / GUADALUPE / DESVIACIONES EPM / DESVIACIÓN TENCHE (TÚNEL TENCHE TRONERAS)"),
        ("ITUANGO",       7.050, -75.650, "Río Cauca"                                      , "Magdalena–Cauca" , "Ituango (Hidroituango)"                   , "EMPRESAS PUBLICAS DE MEDELLIN E.S.P."    , "Antioquia"       , "ITUANGO"),
        ("MUNA",          4.500, -74.250, "Río Bogotá (embalse del Muña)"                  , "Magdalena"       , "Pagua (Paraíso–Guaca)"                    , "ENEL COLOMBIA SA ESP"                    , "Centro"          , "BOGOTÁ / MUÑA"),

        # *** CORREGIDO ***
        ("PUNCHINA",      6.186, -74.883, "Río Guatapé"                                    , "Magdalena–Cauca" , "San Carlos"                               , "ISAGEN S.A. E.S.P."                      , "Antioquia"       , "SAN CARLOS"),
        ("PRADO",         3.735, -74.930, "Río Prado (ríos Negro y Cunday)"                , "Magdalena"       , "Prado"                                    , "CELSIA COLOMBIA S.A. E.S.P."             , "Centro"          , "PRADO"),
        ("ALTOANCHICAYA", 3.554, -76.873, "Río Anchicayá (y ríos Verde y Murrapal)"        , "Pacífico"        , "Alto Anchicayá (Albán)"                   , "CELSIA COLOMBIA S.A. E.S.P."             , "Valle"           , "ALTOANCHICAYÁ"),
    ],
    columns=[
        "Nombre_Visor",
        "Lat",
        "Lon",
        "Rio",
        "Cuenca",
        "Hidro_asociada",
        "Operador",
        "Region_hidrologica",
        "Estacion_hidrologica",
    ],
)

# Renombrar para que el resto del código siga funcionando igual
EMBALSES_SIN_META.rename(columns={"Nombre_Visor": "Embalse_label"}, inplace=True)
EMBALSES_SIN_META["Embalse_norm"] = EMBALSES_SIN_META["Embalse_label"].apply(_normalize_embalse_name)


# -------------------------------------------------------------------
# HELPERS PREDICCIONES – FEATURES Y MODELO RIDGE
# -------------------------------------------------------------------
def _make_calendar_features(idx, with_trend: bool = True) -> pd.DataFrame:
    """
    Crea variables de calendario para series diarias:

      - Día del año (doy), día de la semana y mes
      - Armónicos de Fourier anuales (varios órdenes) para capturar estacionalidad fuerte
      - Armónicos semanales
      - Dummies de mes y día de semana
      - Tendencia lineal opcional

    Esta combinación permite reproducir bien patrones como los picos
    recurrentes de enero y otros meses “especiales”.
    """
    idx = pd.to_datetime(idx)
    if not isinstance(idx, pd.DatetimeIndex):
        idx = pd.DatetimeIndex(idx)

    df_feat = pd.DataFrame(index=idx)

    # Componentes básicos
    df_feat["doy"] = idx.dayofyear.astype(int)     # 1..365
    df_feat["dow"] = idx.dayofweek.astype(int)     # 0..6
    df_feat["month"] = idx.month.astype(int)       # 1..12

    # ---------- Estacionalidad anual: varios armónicos de Fourier ----------
    # Con más armónicos podemos tener picos más marcados (ene, dic, etc.)
    for k in range(1, FOURIER_ORDER_YEARLY + 1):
        df_feat[f"sin_year_{k}"] = np.sin(
            2.0 * np.pi * k * df_feat["doy"] / 365.25
        )
        df_feat[f"cos_year_{k}"] = np.cos(
            2.0 * np.pi * k * df_feat["doy"] / 365.25
        )

    # ---------- Estacionalidad semanal ----------
    df_feat["sin_week"] = np.sin(2.0 * np.pi * df_feat["dow"] / 7.0)
    df_feat["cos_week"] = np.cos(2.0 * np.pi * df_feat["dow"] / 7.0)

    # ---------- Tendencia ----------
    if with_trend:
        # días desde el inicio de la ventana de entrenamiento
        trend = (idx - idx.min()).days.astype(float)
        df_feat["trend"] = trend

    # ---------- Dummies de mes y día de semana ----------
    # Esto permite, por ejemplo, que enero tenga un nivel medio distinto de julio, etc.
    df_feat = pd.get_dummies(
        df_feat,
        columns=["month", "dow"],
        drop_first=True,   # evitamos multicolinealidad perfecta
    )

    return df_feat


def _ridge_forecast_df(
    df_source: pd.DataFrame,
    value_col: str,
    horizon_days: int,
    train_start: dt.date,
    train_end: dt.date,
    allow_trend: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, Optional[object]]:
    """
    Entrena un modelo Ridge con features de calendario dentro del rango
    [train_start, train_end] y genera pronóstico.

    Devuelve:
      df_hist  -> Fecha, y   (histórico USADO para el modelo)
      df_fut   -> Fecha, y   (pronóstico)
      model    -> modelo entrenado (o None si no hay sklearn)
    """
    df = df_source.copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    df = df.sort_values("Fecha")
    df = df.dropna(subset=[value_col])

    if df.empty:
        return df, pd.DataFrame(), None

    # --- acotar histórico explícitamente al rango de entrenamiento ---
    fecha_min = df["Fecha"].min()
    fecha_max = df["Fecha"].max()

    ts_start = max(pd.to_datetime(train_start), fecha_min)
    ts_end = min(pd.to_datetime(train_end), fecha_max)
    if ts_end < ts_start:
        ts_end = ts_start

    # ESTE es el histórico que se usará para entrenar y mostrar
    mask_hist = (df["Fecha"] >= ts_start) & (df["Fecha"] <= ts_end)
    df_hist = df.loc[mask_hist].copy()

    if df_hist.empty:
        return df_hist, pd.DataFrame(), None

    # --- features y target sobre el histórico recortado ---
    y = df_hist[value_col].astype(float).values
    X = _make_calendar_features(df_hist["Fecha"], with_trend=allow_trend)

    if SKLEARN_AVAILABLE:
        model = Pipeline(
            [
                ("scaler", StandardScaler()),
                ("ridge", Ridge(alpha=1.0)),
            ]
        )
        model.fit(X, y)
    else:
        model = None

    # --- fechas futuras: a partir del ÚLTIMO DÍA DEL HISTÓRICO USADO ---
    last_date = df_hist["Fecha"].max()
    future_dates = pd.date_range(
        last_date + pd.Timedelta(days=1),
        periods=horizon_days,
        freq="D",
    )
    X_future = _make_calendar_features(future_dates, with_trend=allow_trend)

    if SKLEARN_AVAILABLE:
        y_future = model.predict(X_future)
    else:
        # fallback: promedio de últimos 30 días DEL HISTÓRICO USADO
        mean_last = df_hist[value_col].tail(30).mean()
        y_future = np.full(horizon_days, float(mean_last))

    # NO hacemos anclaje aquí; lo hacemos luego en la UI
    df_hist_out = df_hist[["Fecha", value_col]].rename(columns={value_col: "y"})
    df_future = pd.DataFrame({"Fecha": future_dates, "y": y_future})

    return df_hist_out, df_future, model


def _plot_forecast(
    df_hist: pd.DataFrame,
    df_future: pd.DataFrame,
    title: str,
    y_label: str,
    mean_value: Optional[float] = None,
) -> go.Figure:
    """Grafica histórico + pronóstico, con línea vertical y media histórica."""
    fig = go.Figure()

    # Histórico
    fig.add_trace(
        go.Scatter(
            x=df_hist["Fecha"],
            y=df_hist["y"],
            mode="lines",
            name="Histórico",
            line=dict(color=PASTEL_BLUE),
        )
    )

    min_date = df_hist["Fecha"].min()
    max_date = df_hist["Fecha"].max()

    if df_future is not None and not df_future.empty:
        # Pronóstico
        fig.add_trace(
            go.Scatter(
                x=df_future["Fecha"],
                y=df_future["y"],
                mode="lines",
                name="Pronóstico",
                line=dict(color=PASTEL_YELLOW, dash="dash"),
            )
        )

        max_date = max(max_date, df_future["Fecha"].max())

        # Rango para la línea vertical
        y_min = min(df_hist["y"].min(), df_future["y"].min())
        y_max = max(df_hist["y"].max(), df_future["y"].max())
        if mean_value is not None:
            y_min = min(y_min, mean_value)
            y_max = max(y_max, mean_value)

        # Línea vertical (inicio del pronóstico) como un Scatter
        split_date = df_hist["Fecha"].max()
        split_x = np.datetime64(pd.to_datetime(split_date))

        fig.add_trace(
            go.Scatter(
                x=[split_x, split_x],
                y=[y_min, y_max],
                mode="lines",
                name="Inicio pronóstico",
                line=dict(color="white", dash="dot"),
                showlegend=True,
                hoverinfo="skip",
            )
        )

    # Línea de media histórica
    if mean_value is not None:
        fig.add_trace(
            go.Scatter(
                x=[min_date, max_date],
                y=[mean_value, mean_value],
                mode="lines",
                name="Media histórica",
                line=dict(color="rgba(255,255,255,0.4)", dash="dot"),
            )
        )

    fig.update_layout(
        template="plotly_dark",
        title=title,
        yaxis_title=y_label,
        xaxis_title="Fecha",
        hovermode="x unified",
    )
    return fig


def _show_series_diagnostics(df_hist: pd.DataFrame, target_col: str, units_label: str):
    """
    Pequeño panel de diagnóstico:
    - promedio anual
    - patrón mensual medio
    - test ADF (si statsmodels está instalado)
    """
    st.markdown("#### Diagnóstico rápido de la serie (período de entrenamiento)")
    with st.expander("Ver diagnóstico de tendencia y estacionalidad"):
        df = df_hist.copy()
        df = df.dropna(subset=["Fecha", target_col])
        if df.empty:
            st.info("No hay datos suficientes para diagnóstico.")
            return

        df["Fecha"] = pd.to_datetime(df["Fecha"])
        df = df.sort_values("Fecha").set_index("Fecha")
        ts = df[target_col].asfreq("D")

        col1, col2 = st.columns(2)

        # Promedio anual
        anuales = ts.resample("A").mean()
        if not anuales.dropna().empty:
            fig_year = px.line(
                x=anuales.index,
                y=anuales.values,
                title="Promedio anual",
                labels={"x": "Año", "y": units_label},
            )
            fig_year.update_layout(template="plotly_dark")
            col1.plotly_chart(fig_year, use_container_width=True)

        # Patrón mensual
        if len(ts.dropna()) > 0:
            df_month = ts.groupby(ts.index.month).mean()
            fig_month = px.bar(
                x=df_month.index,
                y=df_month.values,
                title="Patrón medio mensual",
                labels={"x": "Mes", "y": units_label},
            )
            fig_month.update_layout(template="plotly_dark")
            col2.plotly_chart(fig_month, use_container_width=True)

        # Test ADF
        if adfuller is not None:
            try:
                result = adfuller(ts.dropna())
                st.markdown(
                    f"**ADF:** {result[0]:.3f}  —  **p-value:** {result[1]:.3f}  "
                    "(p-value < 0.05 suele indicar estacionariedad)."
                )
            except Exception as e:
                st.info(f"No se pudo calcular la prueba ADF: {e}")
        else:
            st.info(
                "Instala el paquete `statsmodels` para ver la prueba de estacionariedad ADF "
                "(ej: `pip install statsmodels`)."
            )
def ui_home() -> None:
    st.subheader("Panel principal")
    st.caption("Resumen rápido y navegación a los módulos del visor.")

    st.markdown("---")

    col1, col2, col3 = st.columns(3)

    # ----- Tarjeta Hidrología -----
    with col1:
        st.markdown("### 💧 Hidrología")
        st.write(
            "Niveles de embalses, energía embalsada, aportes hídricos "
            "y días de respaldo."
        )
        if st.button("Ir a Hidrología", key="btn_hidro"):
            go_to("Hidrología")

    # ----- Tarjeta Fuentes de generación -----
    with col2:
        st.markdown("### ⚡ Fuentes de generación")
        st.write(
            "Participación hidráulica, térmica, menores y FERNC, "
            "y detalle por tecnología renovable."
        )
        if st.button("Ir a Fuentes de generación", key="btn_fuentes"):
            go_to("Fuentes de generación")

    # ----- Tarjeta Mercado -----
    with col3:
        st.markdown("### 📈 Mercado")
        st.write(
            "Evolución de la demanda y relación con días de respaldo "
            "hidrológico."
        )
        if st.button("Ir a Mercado", key="btn_mercado"):
            go_to("Mercado")

    st.markdown("---")

    col4, col5 = st.columns(2)

    # ----- Tarjeta Demanda -----
    with col4:
        st.markdown("### 🔌 Demanda")
        st.write("Análisis específico de la demanda diaria del SIN.")
        if st.button("Ir a Demanda", key="btn_demanda"):
            go_to("Demanda")

    # ----- Tarjeta Predicciones -----
    with col5:
        st.markdown("### 🔮 Predicciones")
        st.write(
            "Pronósticos de demanda, nivel de embalses y energía embalsada "
            "con modelo de regresión Ridge."
        )
        if st.button("Ir a Predicciones", key="btn_pred"):
            go_to("Predicciones")

def ui_predicciones() -> None:
    st.subheader("Predicciones")
    st.caption(
        "Pronósticos diarios a partir de series históricas de XM. "
        "Modelo base: Ridge con tendencia y efectos de calendario."
    )

    hoy = dt.date.today()

    # Serie hidrológica completa
    with st.spinner("Construyendo series históricas..."):
        df_all = build_hidro_diaria(GLOBAL_EARLIEST_DATE, hoy)

    if df_all.empty:
        st.warning("No se pudieron obtener series históricas para generar pronósticos.")
        return

    df_all["Fecha"] = pd.to_datetime(df_all["Fecha"])

    # --------------------------- variable ---------------------------
    vars_disp = {
        "Demanda diaria del SIN [GWh/día]": ("demanda_gwh", "GWh/día"),
        "Energía embalsada total [GWh]": ("energia_embalsada_gwh", "GWh"),
        "Nivel útil de embalses [% volumen útil]": ("nivel_embalse_pct", "% volumen útil"),
    }

    nombre_var = st.selectbox("Variable a pronosticar", list(vars_disp.keys()), index=0)
    col_var, unidad_var = vars_disp[nombre_var]

    if col_var not in df_all.columns:
        st.warning(f"La serie '{col_var}' no está disponible en build_hidro_diaria.")
        return

    df_var = df_all[["Fecha", col_var]].dropna()
    if df_var.empty:
        st.warning("No hay datos suficientes de la variable seleccionada.")
        return

    # ---------------------- diagnóstico rápido ----------------------
    with st.expander("Ver diagnóstico de tendencia y estacionalidad"):
        df_diag = df_var.copy()
        df_diag["Año"] = df_diag["Fecha"].dt.year
        df_diag["Mes"] = df_diag["Fecha"].dt.month

        anual = df_diag.groupby("Año")[col_var].mean().reset_index()
        mensual = df_diag.groupby("Mes")[col_var].mean().reset_index()

        c1, c2 = st.columns(2)

        with c1:
            fig_a = px.line(
                anual,
                x="Año",
                y=col_var,
                title="Promedio anual",
                labels={col_var: unidad_var},
            )
            fig_a.update_layout(template="plotly_dark", xaxis_title="Año", yaxis_title=unidad_var)
            st.plotly_chart(fig_a, use_container_width=True)

        with c2:
            fig_m = px.bar(
                mensual,
                x="Mes",
                y=col_var,
                title="Patrón medio mensual",
                labels={col_var: unidad_var},
            )
            fig_m.update_layout(template="plotly_dark", xaxis_title="Mes", yaxis_title=unidad_var)
            st.plotly_chart(fig_m, use_container_width=True)

    # ---------------------- horizonte de forecast -------------------
    st.markdown("### Horizonte de pronóstico")

    horizons = {
        "7 días": 7,
        "30 días": 30,
        "90 días": 90,
        "6 meses (~180 días)": 180,
        "1 año (~365 días)": 365,
        "2 años (~730 días)": 730,
        "5 años (~1825 días)": 1825,
        "10 años (~3650 días)": 3650,
    }

    nombre_h = st.selectbox("Horizonte", list(horizons.keys()), index=5)
    horizonte_dias = horizons[nombre_h]

    # ------------------- periodo de entrenamiento -------------------
    st.markdown("### Período de entrenamiento")

    fecha_min = df_var["Fecha"].min().date()
    fecha_max = df_var["Fecha"].max().date()

    opciones_train = [
        "Últimos 2 años",
        "Últimos 5 años",
        "Últimos 10 años",
        "Todo el histórico",
        "Personalizado",
    ]
    opt_train = st.selectbox("Período base", opciones_train, index=2)

    if opt_train == "Últimos 2 años":
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 2))
    elif opt_train == "Últimos 5 años":
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 5))
    elif opt_train == "Últimos 10 años":
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 10))
    elif opt_train == "Todo el histórico":
        base_start = fecha_min
    else:
        # Personalizado: sugerimos 5 años hacia atrás
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 5))

    c1, c2 = st.columns(2)
    with c1:
        train_start = st.date_input(
            "Fecha inicio entrenamiento",
            value=base_start,
            min_value=fecha_min,
            max_value=fecha_max,
        )
    with c2:
        train_end = st.date_input(
            "Fecha fin entrenamiento",
            value=fecha_max,
            min_value=train_start,
            max_value=fecha_max,
        )

    dias_entrenamiento = (train_end - train_start).days + 1
    st.caption(
        f"Período de entrenamiento: desde **{train_start}** hasta **{train_end}** "
        f"({dias_entrenamiento} días)."
    )

    # ------------------ entrenar y pronosticar ----------------------
    if st.button("Entrenar modelo y generar pronóstico", type="primary"):
        with st.spinner("Entrenando modelo y calculando pronóstico..."):
            df_hist, df_future, model = _ridge_forecast_df(
                df_source=df_var,
                value_col=col_var,
                horizon_days=horizonte_dias,
                train_start=train_start,
                train_end=train_end,
                allow_trend=True,
            )

        if df_future.empty:
            st.warning("No se pudo generar pronóstico con los datos disponibles.")
            return

        # ------------------------------------------------------------------
        # 1) Re-escalar si la variable está en GWh pero numéricamente < 10
        #    (caso típico: viene en cientos de MWh pero la etiquetamos como GWh)
        # ------------------------------------------------------------------
        scale_factor = 1.0
        if "GWh" in unidad_var and df_hist["y"].mean() < 10:
            scale_factor = 1000.0
            df_hist["y"] = df_hist["y"] * scale_factor
            df_future["y"] = df_future["y"] * scale_factor

        # ------------------------------------------------------------------
        # 2) ANCLAJE: que el pronóstico arranque al mismo nivel que el último dato
        # ------------------------------------------------------------------
        first_f = float(df_future["y"].iloc[0])
        last_h = float(df_hist["y"].iloc[-1])

        if np.isfinite(first_f) and np.isfinite(last_h):
            if first_f != 0.0:
                factor = last_h / first_f
                # Si el factor de corrección no es absurdo, usamos anclaje multiplicativo
                if 0.2 < factor < 5.0:
                    df_future["y"] = df_future["y"] * factor
                else:
                    # Si el factor es muy raro, usamos un desplazamiento aditivo
                    df_future["y"] = df_future["y"] + (last_h - first_f)
            else:
                # first_f == 0, usamos sólo desplazamiento aditivo
                df_future["y"] = df_future["y"] + (last_h - first_f)

        # ------------------------------------------------------------------
        # 3) KPIs usando valores ya escalados y anclados
        # ------------------------------------------------------------------
        y_actual = float(df_hist["y"].iloc[-1])
        y_final = float(df_future["y"].iloc[-1])
        cambio_rel = (y_final - y_actual) / y_actual * 100 if y_actual != 0 else 0.0
        mean_hist = float(df_hist["y"].mean())

        c1, c2, c3 = st.columns(3)
        c1.metric("Valor actual", f"{y_actual:,.2f} {unidad_var}")
        c2.metric(
            f"Valor al final del horizonte ({nombre_h})",
            f"{y_final:,.2f} {unidad_var}",
            f"{cambio_rel:+.2f} %",
        )

        modelo_texto = (
            "Ridge (tendencia + estacionalidad calendario)" if SKLEARN_AVAILABLE
            else "Promedio móvil simple (sin scikit-learn)"
        )
        c3.metric("Modelo utilizado", modelo_texto)

        # ------------------------------------------------------------------
        # 4) Gráfica principal con línea de media histórica
        # ------------------------------------------------------------------
        titulo = f"{nombre_var} – histórico y pronóstico"
        fig = _plot_forecast(
            df_hist,
            df_future,
            titulo,
            unidad_var,
            mean_value=mean_hist,
        )
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Predicciones")
    st.caption(
        "Pronósticos diarios a partir de series históricas de XM. "
        "Modelo base: Ridge con tendencia y efectos de calendario."
    )

    hoy = dt.date.today()

    # Serie hidrológica completa
    with st.spinner("Construyendo series históricas..."):
        df_all = build_hidro_diaria(GLOBAL_EARLIEST_DATE, hoy)

    if df_all.empty:
        st.warning("No se pudieron obtener series históricas para generar pronósticos.")
        return

    df_all["Fecha"] = pd.to_datetime(df_all["Fecha"])

    # --------------------------- variable ---------------------------
    vars_disp = {
        "Demanda diaria del SIN [GWh/día]": ("demanda_gwh", "GWh/día"),
        "Energía embalsada total [GWh]": ("energia_embalsada_gwh", "GWh"),
        "Nivel útil de embalses [% volumen útil]": ("nivel_embalse_pct", "% volumen útil"),
    }

    nombre_var = st.selectbox("Variable a pronosticar", list(vars_disp.keys()), index=0)
    col_var, unidad_var = vars_disp[nombre_var]

    if col_var not in df_all.columns:
        st.warning(f"La serie '{col_var}' no está disponible en build_hidro_diaria.")
        return

    df_var = df_all[["Fecha", col_var]].dropna()
    if df_var.empty:
        st.warning("No hay datos suficientes de la variable seleccionada.")
        return

    # ---------------------- diagnóstico rápido ----------------------
    with st.expander("Ver diagnóstico de tendencia y estacionalidad"):
        df_diag = df_var.copy()
        df_diag["Año"] = df_diag["Fecha"].dt.year
        df_diag["Mes"] = df_diag["Fecha"].dt.month

        anual = df_diag.groupby("Año")[col_var].mean().reset_index()
        mensual = df_diag.groupby("Mes")[col_var].mean().reset_index()

        c1, c2 = st.columns(2)

        with c1:
            fig_a = px.line(
                anual,
                x="Año",
                y=col_var,
                title="Promedio anual",
                labels={col_var: unidad_var},
            )
            fig_a.update_layout(template="plotly_dark", xaxis_title="Año", yaxis_title=unidad_var)
            st.plotly_chart(fig_a, use_container_width=True)

        with c2:
            fig_m = px.bar(
                mensual,
                x="Mes",
                y=col_var,
                title="Patrón medio mensual",
                labels={col_var: unidad_var},
            )
            fig_m.update_layout(template="plotly_dark", xaxis_title="Mes", yaxis_title=unidad_var)
            st.plotly_chart(fig_m, use_container_width=True)

    # ---------------------- horizonte de forecast -------------------
    st.markdown("### Horizonte de pronóstico")

    horizons = {
        "7 días": 7,
        "30 días": 30,
        "90 días": 90,
        "6 meses (~180 días)": 180,
        "1 año (~365 días)": 365,
        "2 años (~730 días)": 730,
        "5 años (~1825 días)": 1825,
        "10 años (~3650 días)": 3650,
    }

    nombre_h = st.selectbox("Horizonte", list(horizons.keys()), index=5)
    horizonte_dias = horizons[nombre_h]

    # ------------------- periodo de entrenamiento -------------------
    st.markdown("### Período de entrenamiento")

    fecha_min = df_var["Fecha"].min().date()
    fecha_max = df_var["Fecha"].max().date()

    opciones_train = [
        "Últimos 2 años",
        "Últimos 5 años",
        "Últimos 10 años",
        "Todo el histórico",
        "Personalizado",
    ]
    opt_train = st.selectbox("Período base", opciones_train, index=2)

    if opt_train == "Últimos 2 años":
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 2))
    elif opt_train == "Últimos 5 años":
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 5))
    elif opt_train == "Últimos 10 años":
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 10))
    elif opt_train == "Todo el histórico":
        base_start = fecha_min
    else:
        base_start = max(fecha_min, hoy - dt.timedelta(days=365 * 5))

    # Ahora sí: usuario puede fijar inicio y fin explícitos
    c1, c2 = st.columns(2)
    with c1:
        train_start = st.date_input(
            "Fecha inicio entrenamiento",
            value=base_start,
            min_value=fecha_min,
            max_value=fecha_max,
        )
    with c2:
        train_end = st.date_input(
            "Fecha fin entrenamiento",
            value=fecha_max,
            min_value=train_start,
            max_value=fecha_max,
        )

    dias_entrenamiento = (train_end - train_start).days + 1
    st.caption(
        f"Período de entrenamiento: desde **{train_start}** hasta **{train_end}** "
        f"({dias_entrenamiento} días)."
    )
        # ------------------ entrenar y pronosticar ----------------------
    if st.button("Entrenar modelo y generar pronóstico", type="primary"):
        with st.spinner("Entrenando modelo y calculando pronóstico..."):
            df_hist, df_future, model = _ridge_forecast_df(
                df_source=df_var,
                value_col=col_var,
                horizon_days=horizonte_dias,
                train_start=train_start,
                train_end=train_end,
                allow_trend=True,
            )

        if df_future.empty:
            st.warning("No se pudo generar pronóstico con los datos disponibles.")
            return

        # ------------------------------------------------------------------
        # 1) Re-escalar si la variable está en GWh pero numéricamente < 10
        #    (caso típico: viene en cientos de MWh pero la etiquetamos como GWh)
        # ------------------------------------------------------------------
        scale_factor = 1.0
        if "GWh" in unidad_var and df_hist["y"].mean() < 10:
            scale_factor = 1000.0
            df_hist["y"] = df_hist["y"] * scale_factor
            df_future["y"] = df_future["y"] * scale_factor

        # ------------------------------------------------------------------
        # 2) ANCLAJE: que el pronóstico arranque al mismo nivel que el último dato
        #    Lo hacemos AQUÍ, ya en la escala en la que vamos a graficar.
        # ------------------------------------------------------------------
        first_f = float(df_future["y"].iloc[0])
        last_h = float(df_hist["y"].iloc[-1])

        if np.isfinite(first_f) and np.isfinite(last_h):
            if first_f != 0.0:
                factor = last_h / first_f
                # Si el factor de corrección no es absurdo, usamos anclaje multiplicativo
                if 0.2 < factor < 5.0:
                    df_future["y"] = df_future["y"] * factor
                else:
                    # Si el factor es muy raro, usamos un desplazamiento aditivo
                    df_future["y"] = df_future["y"] + (last_h - first_f)
            else:
                # first_f == 0, usamos sólo desplazamiento aditivo
                df_future["y"] = df_future["y"] + (last_h - first_f)

        # ------------------------------------------------------------------
        # 3) KPIs usando valores ya escalados y anclados
        # ------------------------------------------------------------------
        y_actual = float(df_hist["y"].iloc[-1])
        y_final = float(df_future["y"].iloc[-1])
        cambio_rel = (y_final - y_actual) / y_actual * 100 if y_actual != 0 else 0.0
        mean_hist = float(df_hist["y"].mean())

        c1, c2, c3 = st.columns(3)
        c1.metric("Valor actual", f"{y_actual:,.2f} {unidad_var}")
        c2.metric(
            f"Valor al final del horizonte ({nombre_h})",
            f"{y_final:,.2f} {unidad_var}",
            f"{cambio_rel:+.2f} %",
        )

        modelo_texto = (
            "Ridge (tendencia + estacionalidad calendario)" if SKLEARN_AVAILABLE
            else "Promedio móvil simple (sin scikit-learn)"
        )
        c3.metric("Modelo utilizado", modelo_texto)

        # ------------------------------------------------------------------
        # 4) Gráfica principal con línea de media histórica
        # ------------------------------------------------------------------
        titulo = f"{nombre_var} – histórico y pronóstico"
        fig = _plot_forecast(
            df_hist,
            df_future,
            titulo,
            unidad_var,
            mean_value=mean_hist,
        )
        st.plotly_chart(fig, use_container_width=True)
# -------------------------------------------------------------------
# UI – DEMANDA (enfocado)
# -------------------------------------------------------------------
def ui_demanda() -> None:
    st.subheader("Demanda")
    st.caption("Análisis de la demanda diaria del SIN (DemaSIN)")

    hoy = dt.date.today()
    default_start = hoy - dt.timedelta(days=365)

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Fecha inicio",
            value=default_start,
            format="YYYY-MM-DD",
            key="demanda_start",
        )
    with col2:
        end_date = st.date_input(
            "Fecha fin",
            value=hoy,
            format="YYYY-MM-DD",
            key="demanda_end",
        )

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    df = build_hidro_diaria(start_date, end_date)
    if df.empty or "demanda_gwh" not in df.columns:
        st.info("No se pudo obtener la demanda diaria DemaSIN en este rango.")
        return

    df = df.sort_values("Fecha")

    c1, c2, c3 = st.columns(3)
    c1.metric("Demanda promedio", f"{df['demanda_gwh'].mean():.1f} GWh/día")
    c2.metric("Demanda máxima", f"{df['demanda_gwh'].max():.1f} GWh/día")
    c3.metric("Demanda mínima", f"{df['demanda_gwh'].min():.1f} GWh/día")

    fig = px.line(
        df,
        x="Fecha",
        y="demanda_gwh",
        title="Demanda diaria SIN [GWh/día]",
    )
    fig.update_traces(line_color=PASTEL_YELLOW)
    fig.update_layout(template="plotly_dark", yaxis_title="GWh/día")
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Ver tabla y descargar CSV"):
        st.dataframe(df, use_container_width=True, height=350)
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar CSV demanda",
            data=csv,
            file_name=f"demanda_{start_date}_{end_date}.csv",
            mime="text/csv",
        )


# -------------------------------------------------------------------
# MAIN
# -------------------------------------------------------------------
def main() -> None:
    st.set_page_config(
        page_title="DEMO VISOR",
        page_icon="logo1.png",
        layout="wide",
    )

    set_global_style()
    render_top_header()

    # Sección activa en la sesión (para que go_to() funcione bien)
    if "seccion" not in st.session_state:
        st.session_state["seccion"] = "Panel principal"

    opciones = [
        "Panel principal",
        "Hidrología",
        "Fuentes de generación",
        "Zonas No Interconectadas (ZNI)",
        "Mercado",
        "Demanda",
        "Predicciones",
    ]


    opcion = st.sidebar.radio(
        "Sección",
        opciones,
        index=opciones.index(st.session_state["seccion"]),
    )

    st.session_state["seccion"] = opcion

    if opcion == "Panel principal":
        ui_home()
    elif opcion == "Hidrología":
        ui_hidrologia()
    elif opcion == "Fuentes de generación":
        ui_fuentes_generacion()
    elif opcion == "Zonas No Interconectadas (ZNI)":
        ui_zni()
    elif opcion == "Mercado":
        ui_mercado()
    elif opcion == "Demanda":
        ui_demanda()
    elif opcion == "Predicciones":
        ui_predicciones()



if __name__ == "__main__":
    main()

