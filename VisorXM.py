# app.py
from __future__ import annotations

import base64
import datetime as dt
from pathlib import Path
from typing import List, Optional, Dict

import pandas as pd
import streamlit as st

from pydataxm.pydatasimem import ReadSIMEM
from pydataxm.pydataxm import ReadDB


# -------------------------------------------------------------------
# RUTAS A ARCHIVOS DE APOYO
# -------------------------------------------------------------------
SIMEM_EXCEL = Path("Consulta_API_SIMEM.xlsm")
XM_EXCEL = Path("Consulta_API_XM.xlsm")
LOGO_PATH = Path("logo.png")


# -------------------------------------------------------------------
# UTIL PARA CLAVES (session_state, widgets)
# -------------------------------------------------------------------
def slug(text: str) -> str:
    s = text.lower()
    for a, b in [
        (" ", "_"),
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
        ("ñ", "n"),
    ]:
        s = s.replace(a, b)
    return s


def make_key(prefix: str, section: str) -> str:
    return f"{prefix}_{slug(section)}"


# -------------------------------------------------------------------
# LOGO (CENTRO SUPERIOR, SIN TAPAR CONTENIDO)
# -------------------------------------------------------------------
def add_logo_center(logo_path: Path = LOGO_PATH, height: int = 64) -> None:
    """
    Muestra un logo centrado arriba, pero como parte del layout
    (no fixed) para que no tape nada.
    """
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
# UNIDADES PARA DATASETS SIMEM
# -------------------------------------------------------------------
@st.cache_data
def load_simem_catalog() -> pd.DataFrame:
    """Catálogo de variables SIMEM."""
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
def get_dataset_unit(dataset_id: str) -> Optional[str]:
    """
    Devuelve la unidad declarada en el catálogo SiMEM para un IdDataset.
    Busca en columnas típicas: Unidad, UnidadMedida, etc.
    """
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
    """
    A partir del texto de unidad del catálogo (ej. 'MWh', 'MW', 'GWh')
    decide si es potencia o energía y normaliza a k/M/G.
    """
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

    # Potencia (evitamos confusión con MWh/GWh)
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
    """
    Convierte una serie numérica desde base_unit a target_unit.
    kind = 'power' -> kW/MW/GW, kind = 'energy' -> kWh/MWh/GWh.
    Si no se puede convertir, devuelve la serie numérica tal cual.
    """
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


@st.cache_data
def load_xm_catalog() -> pd.DataFrame:
    """Catálogo de variables de la API XM."""
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
def load_recursos_xm() -> pd.DataFrame:
    """
    Catálogo de recursos (plantas) desde la API XM.

    Métrica: ListadoRecursos / Sistema (ListsEntities)
    Incluye: códigos, nombres, agente, municipio, departamento, etc.
    """
    api = ReadDB()
    dummy_date = "2020-01-01"

    df = api.request_data(
        coleccion="ListadoRecursos",
        metrica="Sistema",
        start_date=dummy_date,
        end_date=dummy_date,
        filtros=None,
    )
    df.columns = [c.strip() for c in df.columns]
    return df


# -------------------------------------------------------------------
# HELPERS PARA GRÁFICAS
# -------------------------------------------------------------------
def numeric_columns(df: pd.DataFrame) -> List[str]:
    """Devuelve columnas numéricas candidatas a graficar."""
    return df.select_dtypes(include=["number"]).columns.tolist()


def plot_dynamic(
    df: pd.DataFrame,
    x_col: Optional[str],
    y_cols: List[str],
    chart_type: str = "Línea",
    aggregate: bool = False,
) -> None:
    """
    Gráfica dinámica:
    - Eje X configurable (cualquier columna)
    - Varias columnas Y
    - Tipo: Línea o Barras
    - Opcional: agrupar por X y sumar las Y
    """
    if not x_col or not y_cols:
        st.info("Selecciona una columna para el eje X y al menos una para el eje Y.")
        return

    if x_col not in df.columns:
        st.warning("La columna seleccionada para el eje X no existe en el DataFrame.")
        return

    for y in list(y_cols):
        if y not in df.columns:
            st.warning(f"La columna '{y}' ya no existe en el DataFrame.")
            return

    df_plot = df[[x_col] + y_cols].copy()

    # ---- convertir X a algo razonable (datetime, numérico o texto) ----
    x_series = df_plot[x_col]
    is_datetime = False

    # 1) intento datetime
    try:
        x_dt = pd.to_datetime(x_series, errors="coerce")
        if x_dt.notna().sum() > 0:
            df_plot[x_col] = x_dt
            is_datetime = True
    except Exception:
        is_datetime = False

    # 2) si no es datetime, intento numérico
    if not is_datetime:
        try:
            x_num = pd.to_numeric(x_series, errors="coerce")
            if x_num.notna().sum() > 0:
                df_plot[x_col] = x_num
            else:
                df_plot[x_col] = x_series.astype(str)
        except Exception:
            df_plot[x_col] = x_series.astype(str)

    # ---- agregación opcional por X ----
    if aggregate:
        df_plot = df_plot.groupby(x_col, as_index=False)[y_cols].sum()

    # ordenar por X cuando se pueda
    try:
        df_plot = df_plot.sort_values(x_col)
    except Exception:
        pass

    df_plot = df_plot.set_index(x_col)

    if chart_type == "Barras":
        st.bar_chart(df_plot[y_cols])
    else:
        st.line_chart(df_plot[y_cols])


# -------------------------------------------------------------------
# API XM – CONSULTA EN TROZOS (PERÍODOS LARGOS)
# -------------------------------------------------------------------
def _sanitize_max_dias(raw_val) -> int:
    """Convierte el Máximo Días del catálogo a un int seguro (>0)."""
    try:
        val = int(raw_val)
    except Exception:
        val = 365
    if val <= 0:
        val = 365
    return val


def fetch_xm_data_chunked(
    coleccion: str,
    metrica: str,
    start_date: dt.date,
    end_date: dt.date,
    max_dias: int,
) -> pd.DataFrame:
    """
    Hace varias llamadas a la API XM respetando max_dias
    y concatena los resultados.
    """
    api = ReadDB()
    dfs = []

    # seguridad
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
            dfs.append(df_tmp)

        current_start = current_end + dt.timedelta(days=1)

    if not dfs:
        return pd.DataFrame()

    df = pd.concat(dfs, ignore_index=True).drop_duplicates()
    df.columns = [c.strip() for c in df.columns]
    return df


# -------------------------------------------------------------------
# SIMEM – CONSULTA GENERAL (CUALQUIER DATASET) EN TROZOS
# -------------------------------------------------------------------
def fetch_simem_data_chunked(
    dataset_id: str,
    start_date: dt.date,
    end_date: dt.date,
    chunk_days: int = 90,
    filter_column: Optional[str] = None,
    filter_values: Optional[str] = None,
    use_filter: bool = False,
) -> pd.DataFrame:
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
            filter_column=filter_column if use_filter else None,
            filter_values=filter_values if use_filter else None,
        )
        df_tmp = simem.main(filter=use_filter)
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
# SIMEM – GENERACIÓN DETALLADA E17D25 EN TROZOS + CATÁLOGO RECURSOS
# -------------------------------------------------------------------
@st.cache_data
def consultar_generacion_detallada_simem(
    start_date: dt.date,
    end_date: dt.date,
    chunk_days: int = 90,
) -> pd.DataFrame:
    dataset_id = "E17D25"
    return fetch_simem_data_chunked(
        dataset_id=dataset_id,
        start_date=start_date,
        end_date=end_date,
        chunk_days=chunk_days,
        use_filter=False,
    )


@st.cache_data
def generar_tabla_generacion_enriquecida(
    start_date: dt.date,
    end_date: dt.date,
) -> pd.DataFrame:
    """
    Une:
    - Generación detallada de SIMEM (E17D25)
    - Catálogo de recursos de XM (ListadoRecursos)
    para agregar nombre de planta, municipio, departamento, agente, etc.
    """
    df_gen = consultar_generacion_detallada_simem(start_date, end_date)
    if df_gen.empty:
        return df_gen

    df_rec = load_recursos_xm()
    if df_rec.empty:
        return df_gen

    df_gen.columns = [c.strip() for c in df_gen.columns]
    df_rec.columns = [c.strip() for c in df_rec.columns]

    # código de planta/recurso en generación
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

    # código en catálogo
    col_codigo_recursos = None
    for c in df_rec.columns:
        cl = c.lower()
        if "codigoplanta" in cl or "codigorecurso" in cl or "idrecurso" in cl:
            col_codigo_recursos = c
            break
    if col_codigo_recursos is None:
        return df_gen

    # columnas descriptivas del catálogo
    cols_info = [col_codigo_recursos]
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

    df_rec_sub = df_rec[cols_info].drop_duplicates(subset=[col_codigo_recursos])

    # merge
    df_merged = df_gen.merge(
        df_rec_sub,
        left_on=col_codigo_gen,
        right_on=col_codigo_recursos,
        how="left",
        suffixes=("", "_cat"),
    )

    # ordenar columnas: info descriptiva primero
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


# -------------------------------------------------------------------
# FILTROS TIPO SINERGOX PARA GENERACIÓN DETALLADA
# -------------------------------------------------------------------
def detect_gen_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Detecta nombres de columnas relevantes en la tabla de generación."""
    cols = [c for c in df.columns]

    def find(*keywords) -> Optional[str]:
        for c in cols:
            cl = c.lower()
            if all(k in cl for k in keywords):
                return c
        return None

    col_fecha = "Fecha" if "Fecha" in df.columns else find("fecha")
    col_planta = None
    for k in ["planta", "nombreplanta", "recurso", "central"]:
        c = find(k)
        if c:
            col_planta = c
            break

    col_tipo_fuente = find("tipogener") or find("tipofuente")
    col_tipo_despacho = find("tipodespacho")
    col_gd = find("gd")
    gen_cols = [c for c in cols if "generacion" in c.lower()]
    col_gen = gen_cols[0] if gen_cols else None

    return {
        "fecha": col_fecha,
        "planta": col_planta,
        "tipo_fuente": col_tipo_fuente,
        "tipo_despacho": col_tipo_despacho,
        "gd": col_gd,
        "gen": col_gen,
    }


def renewable_mask(tipo_fuente_serie: pd.Series) -> pd.Series:
    """Heurística: qué filas son renovables según el texto de tipo de fuente."""
    if tipo_fuente_serie is None:
        return pd.Series([], dtype=bool)

    txt = tipo_fuente_serie.astype(str).str.lower()
    renovables = ["solar", "eolica", "eólico", "hidraul", "biomasa", "geoterm", "ocea", "mareo", "viento"]

    mask = pd.Series(False, index=tipo_fuente_serie.index)
    for k in renovables:
        mask = mask | txt.str.contains(k)
    return mask


# -------------------------------------------------------------------
# UI – GENERACIÓN DETALLADA (SIMEM)
# -------------------------------------------------------------------
def ui_generacion_detallada() -> None:
    """UI específica para Generación detallada (SIMEM)."""
    st.subheader("Generación detallada por planta (SIMEM + catálogo XM)")

    key_prefix = "simem_gen"

    hoy = dt.date.today()
    col1, col2 = st.columns(2)
    with col1:
        ini = st.date_input(
            "Fecha inicio",
            st.session_state.get(f"{key_prefix}_ini", hoy - dt.timedelta(days=365)),
            format="YYYY-MM-DD",
            key=f"{key_prefix}_ini_widget",
        )
    with col2:
        fin = st.date_input(
            "Fecha fin",
            st.session_state.get(f"{key_prefix}_fin", hoy),
            format="YYYY-MM-DD",
            key=f"{key_prefix}_fin_widget",
        )

    st.session_state[f"{key_prefix}_ini"] = ini
    st.session_state[f"{key_prefix}_fin"] = fin

    if ini > fin:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    if st.button("Consultar generación detallada", key=f"{key_prefix}_btn"):
        with st.spinner("Consultando SIMEM (E17D25) y catálogo de recursos XM..."):
            df = generar_tabla_generacion_enriquecida(ini, fin)
        st.session_state[f"{key_prefix}_df"] = df

    df_res: pd.DataFrame = st.session_state.get(f"{key_prefix}_df")

    if df_res is None or df_res.empty:
        st.info("Realiza una consulta para ver los datos de generación detallada.")
        return

    st.success(f"Filas: {df_res.shape[0]} · Columnas: {df_res.shape[1]}")

    info_cols = detect_gen_columns(df_res)
    col_planta = info_cols["planta"]
    col_tipo_fuente = info_cols["tipo_fuente"]
    col_tipo_despacho = info_cols["tipo_despacho"]
    col_gd = info_cols["gd"]
    col_fecha = info_cols["fecha"]
    col_gen = info_cols["gen"]

    # ------------------------------------------------------------------
    # Selección de unidad para la generación (kW/MW/GW o kWh/MWh/GWh)
    # ------------------------------------------------------------------
    base_unit_raw = get_dataset_unit("E17D25")
    unit_info = parse_unit_info(base_unit_raw)
    kind = unit_info["kind"]
    base_unit_norm = unit_info["base_unit"]

    display_unit = None
    if kind in ("power", "energy"):
        if kind == "power":
            options_units = ["kW", "MW", "GW"]
        else:
            options_units = ["kWh", "MWh", "GWh"]

        default_unit = (
            base_unit_norm
            if base_unit_norm in options_units
            else ("MW" if kind == "power" else "MWh")
        )

        display_unit = st.selectbox(
            "Unidad de salida para la generación",
            options_units,
            index=options_units.index(default_unit),
            key=f"{key_prefix}_unit",
        )
        if base_unit_raw:
            st.caption(f"Unidad base de E17D25 en SiMEM: {base_unit_raw}")
    else:
        if base_unit_raw:
            st.caption(f"Unidad registrada para E17D25: {base_unit_raw}")
        else:
            st.caption("No se encontró unidad para E17D25 en el catálogo SiMEM.")

    # ------------------------------------------------------------------
    # Filtros estilo Sinergox
    # ------------------------------------------------------------------
    with st.expander("Filtros de detalle", expanded=True):
        fcol1, fcol2, fcol3, fcol4 = st.columns(4)

        df_filt = df_res.copy()

        if col_planta:
            opciones_planta = ["Todas"] + sorted(
                df_res[col_planta].dropna().astype(str).unique().tolist()
            )
            with fcol1:
                planta_sel = st.selectbox("Planta", opciones_planta, index=0)
            if planta_sel != "Todas":
                df_filt = df_filt[df_filt[col_planta].astype(str) == planta_sel]

        if col_tipo_fuente:
            opciones_tipo = ["Todas"] + sorted(
                df_res[col_tipo_fuente].dropna().astype(str).unique().tolist()
            )
            with fcol2:
                tipo_sel = st.selectbox("Tipo Fuente", opciones_tipo, index=0)
            if tipo_sel != "Todas":
                df_filt = df_filt[df_filt[col_tipo_fuente].astype(str) == tipo_sel]

        if col_tipo_despacho:
            opciones_desp = ["Todas"] + sorted(
                df_res[col_tipo_despacho].dropna().astype(str).unique().tolist()
            )
            with fcol3:
                desp_sel = st.selectbox("Tipo Despacho", opciones_desp, index=0)
            if desp_sel != "Todas":
                df_filt = df_filt[df_filt[col_tipo_despacho].astype(str) == desp_sel]

        if col_gd:
            opciones_gd = ["Todas"] + sorted(
                df_res[col_gd].dropna().astype(str).unique().tolist()
            )
            with fcol4:
                gd_sel = st.selectbox("GD", opciones_gd, index=0)
            if gd_sel != "Todas":
                df_filt = df_filt[df_filt[col_gd].astype(str) == gd_sel]

    # ------------------------------------------------------------------
    # Métricas: Renovable / No Renovable / Total (en unidad elegida)
    # ------------------------------------------------------------------
    if col_gen and col_fecha:
        df_agg = df_filt.copy()

        # Conversión numérica rigurosa
        df_agg[col_gen] = pd.to_numeric(df_agg[col_gen], errors="coerce")
        df_agg[col_fecha] = pd.to_datetime(df_agg[col_fecha], errors="coerce")
        df_agg = df_agg.dropna(subset=[col_fecha, col_gen])

        if not df_agg.empty:
            serie_scaled = convert_series_numeric(
                df_agg[col_gen],
                base_unit=base_unit_norm,
                target_unit=display_unit,
                kind=kind,
            )

            total_gen = serie_scaled.sum()

            if col_tipo_fuente:
                mask_ren = renewable_mask(df_agg[col_tipo_fuente])
                gen_ren = serie_scaled[mask_ren].sum()
                gen_no_ren = serie_scaled[~mask_ren].sum()
            else:
                gen_ren = None
                gen_no_ren = None

            label_unit = display_unit or base_unit_norm or "unid."

            mcol1, mcol2, mcol3 = st.columns(3)
            with mcol1:
                if gen_ren is not None and total_gen:
                    pct_ren = 100 * gen_ren / total_gen
                    st.metric("Generación Renovable [%]", f"{pct_ren:,.2f}")
                else:
                    st.metric("Generación Renovable [%]", "N/A")
            with mcol2:
                if gen_no_ren is not None and total_gen:
                    pct_no = 100 * gen_no_ren / total_gen
                    st.metric("Generación No Renovable [%]", f"{pct_no:,.2f}")
                else:
                    st.metric("Generación No Renovable [%]", "N/A")
            with mcol3:
                st.metric(
                    f"Generación Total [{label_unit}]",
                    f"{total_gen:,.2f}",
                )

    # ------------------------------------------------------------------
    # Tabla + Gráfica dinámica
    # ------------------------------------------------------------------
    tabs = st.tabs(["Tabla", "Gráfica"])

    with tabs[0]:
        st.dataframe(df_filt, use_container_width=True, height=480)
        csv = df_filt.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar CSV (filtrado)",
            data=csv,
            file_name=f"generacion_detallada_filtrada_{ini}_{fin}.csv",
            mime="text/csv",
            key=f"{key_prefix}_download",
        )

    with tabs[1]:
        st.markdown("### Gráfica configurable")
        cols = list(df_filt.columns)
        if not cols:
            st.info("No hay columnas para graficar.")
            return

        # eje X por defecto: Fecha si existe
        default_x = cols.index(col_fecha) if col_fecha in cols else 0
        x_col = st.selectbox(
            "Eje X",
            cols,
            index=default_x,
            key=f"{key_prefix}_xcol",
        )

        # Aseguramos que la columna de generación esté numérica
        # y en la misma unidad elegida (kW/MW/GW o kWh/MWh/GWh)
        if col_gen in df_filt.columns:
            df_filt[col_gen] = convert_series_numeric(
                pd.to_numeric(df_filt[col_gen], errors="coerce"),
                base_unit=base_unit_norm,
                target_unit=display_unit,
                kind=kind,
            )

        num_cols = numeric_columns(df_filt)
        if not num_cols:
            st.info("No hay columnas numéricas para graficar.")
            return

        if col_gen in num_cols:
            default_y = [col_gen]
        else:
            default_y = num_cols[:1]

        y_cols = st.multiselect(
            "Columnas eje Y (numéricas)",
            num_cols,
            default=default_y,
            key=f"{key_prefix}_ycols",
        )

        chart_type = st.radio(
            "Tipo de gráfica",
            ["Línea", "Barras"],
            index=0,
            horizontal=True,
            key=f"{key_prefix}_chart_type",
        )

        aggregate = st.checkbox(
            "Agrupar por eje X (sumar Y)",
            value=(x_col == col_fecha),
            key=f"{key_prefix}_agg",
        )

        plot_dynamic(df_filt, x_col, y_cols, chart_type, aggregate)


# -------------------------------------------------------------------
# UI – CONSULTAS GENERALES SIMEM (TODOS LOS DATASETS)
# -------------------------------------------------------------------
def ui_simem_general() -> None:
    st.subheader("Consultas generales SiMEM (todos los datasets)")

    key_prefix = "simem_general"
    catalog = load_simem_catalog()
    if catalog.empty:
        st.stop()

    # Mostramos directamente el catálogo y luego dejamos escoger el dataset
    resultados = catalog.copy()
    st.write(f"Variables disponibles: {len(resultados)}")
    st.dataframe(resultados.head(200), use_container_width=True, height=260)

    if resultados.empty:
        st.info("No hay variables en el catálogo de SiMEM.")
        return

    def label_row(i) -> str:
        row = resultados.loc[i]
        return f"{row.get('CodigoVariable','')} – {row.get('NombreVariable','')} ({row.get('IdDataset','')})"

    idx = st.selectbox(
        "Selecciona la variable / dataset a consultar:",
        list(resultados.index),
        format_func=label_row,
        key=f"{key_prefix}_var",
    )

    row = resultados.loc[idx]
    dataset_id = str(row.get("IdDataset"))
    nombre_dataset = str(row.get("NombreDataset", ""))
    unidad = row.get("Unidad", row.get("UnidadMedida", row.get("UnidadVariable", "")))
    if pd.isna(unidad):
        unidad = ""

    st.markdown(f"**Dataset seleccionado:** `{dataset_id}`")
    st.caption(f"Nombre dataset: {nombre_dataset}")
    if unidad:
        st.caption(f"Unidad reportada en el catálogo: {unidad}")

    hoy = dt.date.today()
    default_start = st.session_state.get(
        f"{key_prefix}_ini", hoy - dt.timedelta(days=30)
    )
    default_end = st.session_state.get(f"{key_prefix}_fin", hoy)

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Fecha inicio",
            default_start,
            format="YYYY-MM-DD",
            key=f"{key_prefix}_ini_widget",
        )
    with col2:
        end_date = st.date_input(
            "Fecha fin",
            default_end,
            format="YYYY-MM-DD",
            key=f"{key_prefix}_fin_widget",
        )

    st.session_state[f"{key_prefix}_ini"] = start_date
    st.session_state[f"{key_prefix}_fin"] = end_date

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    # filtro opcional por columna/valor(es)
    with st.expander("Filtro opcional por columna ", expanded=False):
        filter_column = st.text_input(
            "Nombre de la columna a filtrar (ej. AGENTE, RECURSO, AREA):",
            "",
            key=f"{key_prefix}_filter_col",
        )
        filter_values_raw = st.text_input(
            "Valores (separados por coma). Si lo dejas vacío, no se aplica filtro:",
            "",
            key=f"{key_prefix}_filter_vals",
        )
        filter_values = (
            ",".join([v.strip() for v in filter_values_raw.split(",") if v.strip()])
            if filter_values_raw
            else None
        )
        use_filter = bool(filter_column and filter_values)

    if st.button("Consultar dataset SiMEM", key=f"{key_prefix}_btn"):
        with st.spinner("Consultando SiMEM (puede tardar para rangos largos)..."):
            df = fetch_simem_data_chunked(
                dataset_id=dataset_id,
                start_date=start_date,
                end_date=end_date,
                chunk_days=90,
                filter_column=filter_column,
                filter_values=filter_values,
                use_filter=use_filter,
            )
        st.session_state[f"{key_prefix}_df"] = df
        st.session_state[f"{key_prefix}_meta"] = {
            "dataset": dataset_id,
            "start_date": str(start_date),
            "end_date": str(end_date),
        }

    df_res: pd.DataFrame = st.session_state.get(f"{key_prefix}_df")
    meta = st.session_state.get(f"{key_prefix}_meta", {})

    if df_res is None or df_res.empty:
        st.info("Realiza una consulta para ver datos del dataset seleccionado.")
        return

    st.success(
        f"Filas: {df_res.shape[0]} · Columnas: {df_res.shape[1]} · "
        f"Período: {meta.get('start_date', start_date)} a {meta.get('end_date', end_date)}"
    )

    tabs = st.tabs(["Tabla", "Gráfica"])

    with tabs[0]:
        st.dataframe(df_res, use_container_width=True, height=480)
        csv = df_res.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar CSV",
            data=csv,
            file_name=f"simem_{meta.get('dataset', dataset_id)}_{meta.get('start_date', start_date)}_{meta.get('end_date', end_date)}.csv",
            mime="text/csv",
            key=f"{key_prefix}_download",
        )

    with tabs[1]:
        st.markdown("### Gráfica configurable")
        cols = list(df_res.columns)
        if not cols:
            st.info("No hay columnas para graficar.")
            return

        x_col = st.selectbox(
            "Eje X",
            cols,
            index=0,
            key=f"{key_prefix}_xcol",
        )

        num_cols = numeric_columns(df_res)
        if not num_cols:
            st.info("No hay columnas numéricas para graficar.")
            return

        y_cols = st.multiselect(
            "Columnas eje Y (numéricas)",
            num_cols,
            default=num_cols[:1],
            key=f"{key_prefix}_ycols",
        )

        chart_type = st.radio(
            "Tipo de gráfica",
            ["Línea", "Barras"],
            index=0,
            horizontal=True,
            key=f"{key_prefix}_chart_type",
        )

        aggregate = st.checkbox(
            "Agrupar por eje X (sumar Y)",
            value=False,
            key=f"{key_prefix}_agg",
        )

        plot_dynamic(df_res, x_col, y_cols, chart_type, aggregate)


# -------------------------------------------------------------------
# UI – CONSULTAS XM
# -------------------------------------------------------------------
def ui_consultas_xm() -> None:
    st.subheader("Consultas XM (API Sinergox – XM)")

    key_prefix = "xm"

    catalog = load_xm_catalog()
    if catalog.empty:
        st.stop()

    # selección de sección (Mercado, Demanda, etc.) si existe
    seccion_col = None
    for c in catalog.columns:
        if c.lower() in ["seccion", "sección", "grupo", "categoria", "categoría"]:
            seccion_col = c
            break

    secciones = ["Todas"]
    if seccion_col is not None:
        secciones += sorted(catalog[seccion_col].dropna().unique().tolist())

    seccion_sel = st.selectbox(
        "Sección",
        secciones,
        index=st.session_state.get(f"{key_prefix}_sec_idx", 0),
        key=f"{key_prefix}_sec",
    )
    st.session_state[f"{key_prefix}_sec_idx"] = secciones.index(seccion_sel)

    df_cat = catalog.copy()
    if seccion_col is not None and seccion_sel != "Todas":
        df_cat = df_cat[df_cat[seccion_col] == seccion_sel]

    # búsqueda
    search = st.text_input(
        "Buscar variable (por nombre, código API o descripción):",
        st.session_state.get(f"{key_prefix}_search", ""),
        placeholder="Ej: Precio Bolsa, Demanda SIN, Embalses...",
        key=f"{key_prefix}_search_widget",
    )
    st.session_state[f"{key_prefix}_search"] = search

    mask = pd.Series(True, index=df_cat.index)
    if search:
        s = search.lower()
        for col in ["Nombre Variable", "Código API", "Codigo API", "Descripción", "Descripcion"]:
            if col in df_cat.columns:
                mask &= df_cat[col].astype(str).str.lower().str.contains(s)

    resultados = df_cat[mask].copy()
    st.write(f"Variables encontradas: {len(resultados)}")
    st.dataframe(resultados.head(200), use_container_width=True, height=260)

    if resultados.empty:
        st.info("Ajusta el filtro de búsqueda o la sección para ver variables.")
        return

    def label_var(i) -> str:
        nv = resultados.loc[i].get("Nombre Variable", "")
        ca = resultados.loc[i].get("Código API", resultados.loc[i].get("Codigo API", ""))
        gr = resultados.loc[i].get("Granularidad", "")
        return f"{nv} ({ca} – {gr})"

    idx = st.selectbox(
        "Selecciona la variable a consultar:",
        list(resultados.index),
        format_func=label_var,
        key=f"{key_prefix}_var",
    )

    fila = resultados.loc[idx]

    coleccion = str(fila.get("Código API", fila.get("Codigo API")))
    metrica = str(fila.get("Granularidad", fila.get("Metrica", "Sistema")))
    max_dias = _sanitize_max_dias(fila.get("Máximo Días", fila.get("MaxDias", 365)))
    tipo = str(fila.get("Desagregación", fila.get("Type", "")))
    unidad = fila.get("Unidad", fila.get("Unidad Medida", ""))

    st.markdown(f"**Código API (colección):** `{coleccion}`")
    st.markdown(f"**Métrica / entidad (granularidad):** `{metrica}`")
    st.caption(f"Tipo interno: {tipo} · Máximo días por llamada: {max_dias}")
    if pd.notna(unidad) and str(unidad).strip():
        st.caption(f"Unidad principal (según catálogo): {unidad}")

    hoy = dt.date.today()
    default_start = st.session_state.get(
        f"{key_prefix}_ini", hoy - dt.timedelta(days=min(365, max_dias * 3))
    )
    default_end = st.session_state.get(f"{key_prefix}_fin", hoy)

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Fecha inicio",
            default_start,
            format="YYYY-MM-DD",
            key=f"{key_prefix}_ini_widget",
        )
    with col2:
        end_date = st.date_input(
            "Fecha fin",
            default_end,
            format="YYYY-MM-DD",
            key=f"{key_prefix}_fin_widget",
        )

    st.session_state[f"{key_prefix}_ini"] = start_date
    st.session_state[f"{key_prefix}_fin"] = end_date

    if start_date > end_date:
        st.error("La fecha de inicio no puede ser mayor que la fecha fin.")
        return

    delta = (end_date - start_date).days + 1
    if delta > max_dias:
        st.info(
            f"La app dividirá automáticamente el período seleccionado "
            f"en trozos de {max_dias} días para descargar todo el rango."
        )

    if st.button("Consultar API XM", key=f"{key_prefix}_btn"):
        with st.spinner("Consultando API XM (puede tardar para rangos largos)..."):
            df = fetch_xm_data_chunked(
                coleccion=coleccion,
                metrica=metrica,
                start_date=start_date,
                end_date=end_date,
                max_dias=max_dias,
            )
        st.session_state[f"{key_prefix}_df"] = df
        st.session_state[f"{key_prefix}_meta"] = {
            "coleccion": coleccion,
            "metrica": metrica,
            "start_date": str(start_date),
            "end_date": str(end_date),
        }

    df_res: pd.DataFrame = st.session_state.get(f"{key_prefix}_df")
    meta = st.session_state.get(f"{key_prefix}_meta", {})

    if df_res is None or df_res.empty:
        st.info("Realiza una consulta para ver datos de la API XM.")
        return

    st.success(
        f"Filas: {df_res.shape[0]} · Columnas: {df_res.shape[1]} · "
        f"Período: {meta.get('start_date', start_date)} a {meta.get('end_date', end_date)}"
    )

    tabs = st.tabs(["Tabla", "Gráfica"])

    with tabs[0]:
        st.dataframe(df_res, use_container_width=True, height=480)
        csv = df_res.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar CSV",
            data=csv,
            file_name=f"xm_{meta.get('coleccion', coleccion)}_{meta.get('metrica', metrica)}_{meta.get('start_date', start_date)}_{meta.get('end_date', end_date)}.csv",
            mime="text/csv",
            key=f"{key_prefix}_download",
        )

    with tabs[1]:
        st.markdown("### Gráfica configurable")

        cols = list(df_res.columns)
        if not cols:
            st.info("No hay columnas para graficar.")
            return

        # intentar sugerir una columna X de tiempo
        x_default = 0
        for c in cols:
            if any(k in c.lower() for k in ["fecha", "date", "hora", "datetime"]):
                x_default = cols.index(c)
                break

        x_col = st.selectbox(
            "Eje X",
            cols,
            index=x_default,
            key=f"{key_prefix}_xcol",
        )

        num_cols = numeric_columns(df_res)
        if not num_cols:
            st.info("No hay columnas numéricas para graficar.")
            return

        y_cols = st.multiselect(
            "Columnas eje Y (numéricas)",
            num_cols,
            default=num_cols[:1],
            key=f"{key_prefix}_ycols",
        )

        chart_type = st.radio(
            "Tipo de gráfica",
            ["Línea", "Barras"],
            index=0,
            horizontal=True,
            key=f"{key_prefix}_chart_type",
        )

        aggregate = st.checkbox(
            "Agrupar por eje X (sumar Y)",
            value=("fecha" in x_col.lower() or "date" in x_col.lower()),
            key=f"{key_prefix}_agg",
        )

        plot_dynamic(df_res, x_col, y_cols, chart_type, aggregate)


# -------------------------------------------------------------------
# MAIN
# -------------------------------------------------------------------
def main() -> None:
    st.set_page_config(
        page_title="Visor XM - SiMEM via Python",
        layout="wide",
    )

    add_logo_center()

    st.title("Visor XM - SiMEM via Python")
    st.caption(
        "Consultas profesionales sobre las APIs de XM y SiMEM, con tablas, filtros, "
        "gráficas dinámicas y descargas CSV."
    )
    st.caption("Juan Esteban Rodríguez Villada - Noviembre 2025")

    opcion = st.sidebar.radio(
        "Modo de consulta",
        ["Consultas XM", "Consultas SiMEM"],
        index=0,
    )

    if opcion == "Consultas XM":
        ui_consultas_xm()
    else:
        tabs = st.tabs(["General (todos los datasets)", "Generación detallada (E17D25)"])
        with tabs[0]:
            ui_simem_general()
        with tabs[1]:
            ui_generacion_detallada()


if __name__ == "__main__":
    main()