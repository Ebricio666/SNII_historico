"""
SNII Insight Dashboard
======================

Esta aplicación NO descarga Google Drive y NO procesa PDF.
Únicamente lee el archivo ya construido:

    data/SNII_MASTER.parquet

Esto permite que la aplicación cargue rápidamente.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
MASTER_PATH = DATA_DIR / "SNII_MASTER.parquet"


st.set_page_config(
    page_title="SNII Insight",
    page_icon="📊",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def cargar_master(ruta: str) -> pd.DataFrame:
    return pd.read_parquet(ruta)


def mostrar_error_master() -> None:
    st.error(
        "No se encontró `data/SNII_MASTER.parquet`."
    )

    st.markdown(
        """
Ejecuta primero el constructor:

```bash
python builder.py
```

Después sube a GitHub el archivo generado:

```text
data/SNII_MASTER.parquet
```
"""
    )


st.title("📊 SNII Insight")
st.caption(
    "Sistema para la integración histórica y análisis evolutivo "
    "del Sistema Nacional de Investigadoras e Investigadores"
)

if not MASTER_PATH.exists():
    mostrar_error_master()
    st.stop()

try:
    master = cargar_master(str(MASTER_PATH))
except Exception as error:
    st.error("No fue posible abrir el archivo maestro.")
    st.exception(error)
    st.stop()

# ============================================================
# RESUMEN GENERAL
# ============================================================

st.success("Base histórica cargada correctamente.")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Registros",
    f"{len(master):,}",
)

col2.metric(
    "Años",
    (
        master["ANIO"].nunique()
        if "ANIO" in master.columns
        else 0
    ),
)

col3.metric(
    "Archivos fuente",
    (
        master["ORIGEN_ARCHIVO"].nunique()
        if "ORIGEN_ARCHIVO" in master.columns
        else 0
    ),
)

col4.metric(
    "Variables",
    len(master.columns),
)

# ============================================================
# FILTROS
# ============================================================

st.subheader("Exploración de la base")

filtro1, filtro2, filtro3 = st.columns(3)

df_filtrado = master.copy()

with filtro1:
    if "ANIO" in master.columns:
        anios = sorted(
            master["ANIO"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        anios_seleccionados = st.multiselect(
            "Año",
            options=anios,
            default=anios,
        )

        if anios_seleccionados:
            df_filtrado = df_filtrado[
                df_filtrado["ANIO"]
                .astype(str)
                .isin(anios_seleccionados)
            ]

with filtro2:
    if "NIVEL" in master.columns:
        niveles = sorted(
            master["NIVEL"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        niveles_seleccionados = st.multiselect(
            "Nivel",
            options=niveles,
        )

        if niveles_seleccionados:
            df_filtrado = df_filtrado[
                df_filtrado["NIVEL"]
                .astype(str)
                .isin(niveles_seleccionados)
            ]

with filtro3:
    texto_busqueda = st.text_input(
        "Buscar persona, institución o CVU",
        placeholder="Escribe un término...",
    )

if texto_busqueda:
    columnas_texto = [
        columna
        for columna in df_filtrado.columns
        if (
            pd.api.types.is_object_dtype(
                df_filtrado[columna]
            )
            or pd.api.types.is_string_dtype(
                df_filtrado[columna]
            )
        )
    ]

    mascara = pd.Series(
        False,
        index=df_filtrado.index,
    )

    for columna in columnas_texto:
        mascara = mascara | (
            df_filtrado[columna]
            .astype(str)
            .str.contains(
                texto_busqueda,
                case=False,
                na=False,
                regex=False,
            )
        )

    df_filtrado = df_filtrado[mascara]

st.caption(
    f"Registros mostrados: {len(df_filtrado):,}"
)

st.dataframe(
    df_filtrado,
    use_container_width=True,
    hide_index=True,
    height=520,
)

# ============================================================
# DESCARGAS
# ============================================================

st.subheader("Descargar datos filtrados")

csv_filtrado = df_filtrado.to_csv(
    index=False
).encode("utf-8-sig")

st.download_button(
    "⬇️ Descargar selección en CSV",
    data=csv_filtrado,
    file_name="SNII_Insight_filtrado.csv",
    mime="text/csv",
    type="primary",
)

st.caption(
    "El dashboard sólo lee el archivo Parquet ya procesado. "
    "No vuelve a descargar ni reconstruir la base."
)
