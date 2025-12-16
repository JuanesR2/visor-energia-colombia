
![Logo MME](logo.png)

# Visor XM – SiMEM vía Python  
**Prototipo de tablero de datos para el mercado eléctrico colombiano**

Autor: **Juan Esteban Rodríguez Villada**  
Versión: **0.1 (demo/prototipo)** – Noviembre 2025  

Aplicación en **Python + Streamlit** para explorar datos públicos de:

- **API XM / Sinergox**
- **API SiMEM**

La idea es ofrecer una alternativa a las macros de Excel: consultas más flexibles, filtros tipo tablero, gráficas rápidas y descarga a CSV.  
> Nota: es un **prototipo interno**, no una herramienta oficial de XM.

---

## 1. Requisitos

### 1.1 Dependencias

- Python 3.10+
- Paquetes:

```bash
pip install streamlit pandas pydataxm
````

### 1.2 Archivos necesarios (en la misma carpeta que `app.py`)

* `Consulta_API_XM.xlsm` → catálogo de variables de la API XM (hoja `Parametros`).
* `Consulta_API_SIMEM.xlsm` → catálogo de datasets SiMEM (hoja `ListadoVariables`).
* `logo.png` → imagen que se muestra centrada en la parte superior.

Ejecutar:
```bash
streamlit run app.py
````

## 2. Qué hace el prototipo

La app tiene dos modos principales (selección en el sidebar):

1. **Consultas XM**

   * Usa el catálogo de `Consulta_API_XM.xlsm`.
   * Permite buscar una variable, elegir rango de fechas y consultar la API XM.
   * Si el período es mayor al **Máximo Días** de la variable, la app lo trocea en varios llamados.
   * Muestra:

     * **Tabla** con descarga CSV.
     * **Gráfica configurable**:

       * Eje X: cualquier columna.
       * Eje Y: una o varias columnas numéricas.
       * Tipo: línea o barras.
       * Opción *“Agrupar por eje X (sumar Y)”*.

2. **Consultas SiMEM**
   Dos pestañas:

   * **General (todos los datasets)**

     * Usa `Consulta_API_SIMEM.xlsm` para buscar `IdDataset`.
     * Se puede filtrar por columna (ej. AGENTE, AREA).
     * Descarga el rango completo dividiéndolo en ventanas de 90 días.
     * Tabla con descarga CSV y gráfica dinámica (mismo esquema que XM).

   * **Generación detallada (E17D25)**

     * Descarga el dataset E17D25 de SiMEM.
     * Lo enriquece con el catálogo de recursos de XM (`ListadoRecursos / Sistema`):
       nombre de planta, agente, municipio, departamento, etc.
     * Detecta automáticamente columnas clave: fecha, planta, tipo de fuente, tipo de despacho, GD y columna principal de generación.
     * Permite filtrar por:

       * Planta
       * Tipo Fuente
       * Tipo Despacho
       * GD

---

## 3. Unidades y métricas de generación

Para E17D25:

* Se consulta al catálogo de SiMEM la **unidad base** (ej. MWh o MW).
* Se clasifica como:

  * **Energía:** kWh / MWh / GWh
  * **Potencia:** kW / MW / GW
* El usuario elige la **unidad de salida**, y la app convierte numéricamente la serie de generación a esa unidad.

Se calculan tres métricas principales:

1. **Generación Renovable [%]**
2. **Generación No Renovable [%]**
3. **Generación Total [unidad seleccionada]**

La clasificación renovable/no renovable se hace por texto en el tipo de fuente (solar, eólica, hidráulica, biomasa, etc.).

---

## 4. Estructura interna (resumen)

En `app.py`:

* Utilidades:

  * `add_logo_center` → muestra `logo.png`.
  * `numeric_columns`, `plot_dynamic` → soporte general para gráficas.
* Carga de catálogos:

  * `load_simem_catalog`, `load_xm_catalog`, `load_recursos_xm`.
* Tratamiento de unidades:

  * `get_dataset_unit`, `parse_unit_info`, `convert_series_numeric`.
* Descarga de datos:

  * `fetch_xm_data_chunked` (API XM con troceo de fechas).
  * `fetch_simem_data_chunked` (API SiMEM).
* Paneles de UI:

  * `ui_consultas_xm`
  * `ui_simem_general`
  * `ui_generacion_detallada`
* Entrada principal:

  * `main()` → configura Streamlit (`layout="wide"`), muestra el logo, título y selector de modo.

---

## 5. Limitaciones y futuras mejoras

**Limitaciones actuales**

* Prototipo pensado para uso local; sin autenticación ni control de acceso.
* El rendimiento puede bajar con volúmenes muy grandes de datos (todo va a memoria con pandas).
* La interpretación de unidades depende de lo que venga en los catálogos de Excel.

**Ideas de mejora**

* Paneles predefinidos (demanda SIN, precios bolsa, etc.).
* Cache local o base de datos para no repetir descargas.
* Mapas con generación por departamento.
* Exportar reportes automáticos (PDF/PowerPoint).

---

```
::contentReference[oaicite:0]{index=0}
```
