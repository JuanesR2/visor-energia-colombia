# hidrologia.py
from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from pydataxm.pydataxm import ReadDB

# -------------------------------------------------------------------
# Rutas y colores (paleta estilo Ministerio de Minas y Energía)
# -------------------------------------------------------------------
XM_EXCEL = Path("Consulta_API_XM.xlsm")

PRIMARY_PURPLE = "#5A2D82"   # Morado institucional aproximado
ACCENT_YELLOW = "#FFC400"    # Amarillo acento
BG_PANEL = "#F7F4FF"         # Lila muy claro
TEXT_DARK = "#222222"


# -------------------------------------------------------------------
# Catálogo XM y descarga en trozos
# -------------------------------------------------------------------
@st.cache_data
def load_xm_catalog() -> pd.DataFrame:
    """Carga el catálogo de variables XM desde el Excel de apoyo."""
    if not XM_EXCEL.exists():
        st.warning(
            f"No se encontró {XM_EXCEL}. "
            "Pon Consulta_API_XM.xlsm en la misma carpeta que app.py."
        )
        return pd.DataFrame()

    df = pd.read_excel(XM_EXCEL, sheet_name="Parametros")
    df.columns = [c.strip() for c in df.columns]
    return df


def _sanitize_max_dias(raw_val) -> int:
    """Convierte el 'Máximo Días' del catálogo a un int seguro (>0)."""
    try:
        val = int(raw_val)
    except Exception:
        val = 365
    if val <= 0:
        val = 365
    return val


@st.cache_data
def get_xm_meta(codigo_api: str) -> Dict[str, Optional[object]]:
    """
    Devuelve metadatos básicos de una variable XM:
    - metrica (granularidad, ej. 'Sistema')
    - max_dias (máximo de días por llamada)
    """
    catalog = load_xm_catalog()
    if catalog.empty:
        return {"metrica": "Sistema", "max_dias": 365}

    mask = (
        (catalog.get("Código API", "") == codigo_api)
        | (catalog.get("Codigo API", "") == codigo_api)
    )
    if not mask.any():
        return {"metrica": "Sistema", "max_dias": 365}

    row = catalog[mask].iloc[0]
    metrica = row.get("Granularidad", row.get("Metrica", "Sistema"))
    max_raw = row.get("Máximo Días", row.get("MaxDias", 365))
    return {"metrica": str(metrica), "max_dias": _sanitize_max_dias(max_raw)}


@st.cache_data
def fetch_xm_data_chunked(
    coleccion: str,
    metrica: str,
    start_date: dt.date,
    end_date: dt.date,
    max_dias: int,
) -> pd.DataFrame:
    """
    Igual a la lógica de app.py: consulta la API XM en trozos de
    'max_dias' y concatena resultados.
    """
    api = ReadDB()
    dfs = []

    max_dias = _sanitize_max_dias(max_dias)
    current_start = start_date

    while current_start <= end_date:
        current_end = min(
            current_start + dt.timedelta(days=max_dias - 1),
            end_date,
        )
        df_tmp = api.request_data(
            coleccion=coleccion,
            metrica=metrica,
            start_date=str(current_start),
            end_date=str(current_end),
            filtros=None,
        )
        if not df_tmp.empty:
            df_tmp.columns = [c.strip() for c in df_tmp.columns]
            dfs.append(df_tmp)

        current_start = current_end + dt.timedelta(days=1)

    if not dfs:
        return pd.DataFrame()

    df = pd.concat(dfs, ignore_index=True).drop_duplicates()
    df.columns = [c.strip() for c in df.columns]
    return df


# -------------------------------------------------------------------
# Helpers para detectar columnas de fecha / valor
# -------------------------------------------------------------------
def detect_date_col(df: pd.DataFrame) -> str:
    for c in df.columns:
        cl = c.lower()
        if "fecha" in cl or "date" in cl:
            return c
    return df.columns[0]


def detect_value_col(df: pd.DataFrame) -> Optional[str]:
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if not num_cols:
        return None

    # Evitar columnas tipo año/mes/día
    for c in num_cols:
        cl = c.lower()
        if not any(k in cl for k in ["anio", "año", "mes", "dia", "día"]):
            return c
    return num_cols[0]


# -------------------------------------------------------------------
# Construcción de tabla diaria hidrológica (SIN)
# -------------------------------------------------------------------
@st.cache_data
def build_hydro_daily(
    start_date: dt.date,
    end_date: dt.date,
) -> pd.DataFrame:
    """
    Descarga y arma una tabla diaria con las variables hidrológicas clave
    (a nivel SIN agregando por Sistema).

    Columnas principales:
    - Fecha
    - nivel_embalse_pct
    - energia_embalsada_gwh
    - aportes_gwh
    - aportes_hist_gwh
    - aportes_pct_hist
    - demanda_gwh
    - dias_respaldo
    """
    # Código API XM -> cómo agregar
    variables = {
        "nivel_embalse_pct": {"codigo": "PorcVoluUtilDiar", "agg": "mean"},
        "energia_embalsada_gwh": {"codigo": "VoluUtilDiarEner", "agg": "sum"},
        "aportes_gwh": {"codigo": "AporEner", "agg": "sum"},
        "aportes_hist_gwh": {"codigo": "AporEnerMediHist", "agg": "sum"},
        "demanda_gwh": {"codigo": "DemaSIN", "agg": "sum"},
    }

    series = []

    for nombre, spec in variables.items():
        meta = get_xm_meta(spec["codigo"])
        df_raw = fetch_xm_data_chunked(
            coleccion=spec["codigo"],
            metrica=meta["metrica"],
            start_date=start_date,
            end_date=end_date,
            max_dias=meta["max_dias"],
        )
        if df_raw.empty:
            continue

        fecha_col = detect_date_col(df_raw)
        valor_col = detect_value_col(df_raw)
        if valor_col is None:
            continue

        df_tmp = df_raw[[fecha_col, valor_col]].copy()
        df_tmp[fecha_col] = pd.to_datetime(df_tmp[fecha_col], errors="coerce").dt.date
        df_tmp = df_tmp.dropna(subset=[fecha_col])

        if spec["agg"] == "mean":
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

    # Derivadas útiles
    if {"aportes_gwh", "aportes_hist_gwh"}.issubset(df_all.columns):
        df_all["aportes_pct_hist"] = (
            100.0 * df_all["aportes_gwh"] / df_all["aportes_hist_gwh"]
        )

    if {"energia_embalsada_gwh", "demanda_gwh"}.issubset(df_all.columns):
        df_all["dias_respaldo"] = (
            df_all["energia_embalsada_gwh"] / df_all["demanda_gwh"]
        )

    return df_all


# -------------------------------------------------------------------
# UI principal del tablero de Hidrología
# -------------------------------------------------------------------
def ui_hidrologia() -> None:
    """Tablero de Embalses y Aportes, estilo MME."""

    # Estilos ligeros con paleta MME
    st.markdown(
        f"""
        <style>
        .hydro-header {{
            background: linear-gradient(90deg, {PRIMARY_PURPLE}, #7B3FB8);
            color: white;
            padding: 0.9rem 1.4rem;
            border-radius: 0.8rem;
            margin-bottom: 1rem;
        }}
        .hydro-header h2 {{
            margin: 0;
            font-size: 1.5rem;
        }}
        .hydro-header p {{
            margin: 0.15rem 0 0 0;
            font-size: 0.9rem;
            opacity: 0.9;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="hydro-header">
            <h2>Embalses y Aportes</h2>
            <p>Dirección de Energía Eléctrica – Seguimiento hidrológico del SIN</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ------------------------ Rango de fechas ------------------------
    hoy = dt.date.today()
    default_start = hoy - dt.timedelta(days=365 * 3)

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Fecha inicio",
            value=default_start,
            format="YYYY-MM-DD",
            key="hidro_start",
        )
    with col2:
        end_date = st.date_input(
            "Fecha fin",
            value=hoy,
            format="YYYY-MM-DD",
            key="hidro_end",
        )

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    with st.spinner("Consultando XM y construyendo serie hidrológica..."):
        df = build_hydro_daily(start_date, end_date)

    if df.empty:
        st.warning("No se obtuvieron datos hidrológicos para el rango seleccionado.")
        return

    df = df.sort_values("Fecha")
    last_row = df.iloc[-1]
    first_row = df.iloc[0]

    # ------------------------ KPIs principales -----------------------
    k1, k2, k3, k4 = st.columns(4)

    # Nivel de embalses
    if "nivel_embalse_pct" in df.columns:
        delta_nivel = last_row["nivel_embalse_pct"] - first_row["nivel_embalse_pct"]
        with k1:
            st.metric(
                "Nivel de embalses",
                f"{last_row['nivel_embalse_pct']:.2f} %",
                f"{delta_nivel:+.2f} pp",
            )

    # Energía embalsada
    if "energia_embalsada_gwh" in df.columns:
        with k2:
            st.metric(
                "Energía embalsada",
                f"{last_row['energia_embalsada_gwh']:,.0f} GWh",
            )

    # % Aportes vs media histórica
    if "aportes_pct_hist" in df.columns:
        delta_aportes = last_row["aportes_pct_hist"] - df["aportes_pct_hist"].mean()
        with k3:
            st.metric(
                "% Aportes vs histórico",
                f"{last_row['aportes_pct_hist']:.1f} %",
                f"{delta_aportes:+.1f} pts",
            )

    # Días de respaldo
    if "dias_respaldo" in df.columns:
        with k4:
            st.metric(
                "Días de respaldo hidro",
                f"{last_row['dias_respaldo']:.1f} días",
            )

    st.markdown("---")

    # ------------------------ Gráficos principales -------------------
    c1, c2 = st.columns(2)

    # Nivel de embalses (%)
    if "nivel_embalse_pct" in df.columns:
        with c1:
            fig1 = px.area(
                df,
                x="Fecha",
                y="nivel_embalse_pct",
                title="Nivel de embalses (%)",
            )
            fig1.update_traces(
                line_color=PRIMARY_PURPLE,
                fillcolor="rgba(90,45,130,0.30)",
            )
            fig1.update_layout(
                template="plotly_white",
                title_font_color=PRIMARY_PURPLE,
                yaxis_title="% volumen útil",
            )
            st.plotly_chart(fig1, use_container_width=True)

    # Aportes vs media histórica
    if {"aportes_gwh", "aportes_hist_gwh"}.issubset(df.columns):
        with c2:
            fig2 = go.Figure()
            fig2.add_trace(
                go.Scatter(
                    x=df["Fecha"],
                    y=df["aportes_gwh"],
                    mode="lines",
                    name="Aportes hídricos [GWh/día]",
                    line=dict(color=PRIMARY_PURPLE, width=2),
                    fill="tozeroy",
                    fillcolor="rgba(90,45,130,0.25)",
                )
            )
            fig2.add_trace(
                go.Scatter(
                    x=df["Fecha"],
                    y=df["aportes_hist_gwh"],
                    mode="lines",
                    name="Media histórica [GWh/día]",
                    line=dict(color=ACCENT_YELLOW, width=3),
                )
            )
            fig2.update_layout(
                template="plotly_white",
                title="Aportes hídricos vs media histórica",
                title_font_color=PRIMARY_PURPLE,
                yaxis_title="GWh/día",
            )
            st.plotly_chart(fig2, use_container_width=True)

    # Demanda vs nivel de embalses
    if {"demanda_gwh", "nivel_embalse_pct"}.issubset(df.columns):
        st.markdown("### Demanda vs nivel de embalses")

        fig3 = go.Figure()
        fig3.add_trace(
            go.Scatter(
                x=df["Fecha"],
                y=df["demanda_gwh"],
                name="Demanda SIN [GWh/día]",
                line=dict(color="#444444"),
            )
        )
        fig3.add_trace(
            go.Scatter(
                x=df["Fecha"],
                y=df["nivel_embalse_pct"],
                name="Nivel embalses (%)",
                line=dict(color=PRIMARY_PURPLE, dash="dot"),
                yaxis="y2",
            )
        )
        fig3.update_layout(
            template="plotly_white",
            title="Demanda vs nivel de embalses",
            title_font_color=PRIMARY_PURPLE,
            yaxis=dict(title="Demanda [GWh/día]"),
            yaxis2=dict(
                title="Nivel embalses (%)",
                overlaying="y",
                side="right",
            ),
        )
        st.plotly_chart(fig3, use_container_width=True)

    # Tabla y descarga
    with st.expander("Ver tabla hidrológica y descargar CSV"):
        st.dataframe(df, use_container_width=True, height=380)
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar CSV hidrológico",
            data=csv,
            file_name=f"hidrologia_{start_date}_{end_date}.csv",
            mime="text/csv",
        )
