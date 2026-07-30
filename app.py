import os
import tempfile
from pathlib import Path

import gdown
import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1UQ_sPApThDd3xHMQPz0GvMuIfLTcd_4x?usp=sharing"
)

EXTENSIONES_PERMITIDAS = {
    ".xlsx",
    ".xls",
    ".pdf",
    ".csv",
}


# ============================================================
# FUNCIONES
# ============================================================

def descargar_carpeta_drive(url: str, destino: str) -> list[str]:
    """
    Descarga los archivos visibles de una carpeta pública de Google Drive.

    Parámetros
    ----------
    url:
        Enlace público de la carpeta de Google Drive.
    destino:
        Carpeta temporal donde se guardarán los archivos.

    Retorna
    -------
    list[str]
        Lista de rutas descargadas.
    """
    archivos = gdown.download_folder(
        url=url,
        output=destino,
        quiet=True,
        use_cookies=False,
        remaining_ok=True,
    )

    return archivos or []


def obtener_archivos_compatibles(carpeta: str) -> pd.DataFrame:
    """
    Busca archivos compatibles dentro de la carpeta descargada.
    """
    registros = []

    for ruta in Path(carpeta).rglob("*"):
        if not ruta.is_file():
            continue

        extension = ruta.suffix.lower()

        if extension not in EXTENSIONES_PERMITIDAS:
            continue

        registros.append(
            {
                "Archivo": ruta.name,
                "Tipo": extension.replace(".", "").upper(),
                "Tamaño (MB)": round(ruta.stat().st_size / (1024 * 1024), 2),
                "Ruta relativa": str(ruta.relative_to(carpeta)),
            }
        )

    if not registros:
        return pd.DataFrame(
            columns=["Archivo", "Tipo", "Tamaño (MB)", "Ruta relativa"]
        )

    return (
        pd.DataFrame(registros)
        .sort_values(["Tipo", "Archivo"])
        .reset_index(drop=True)
    )


def contar_por_tipo(df_archivos: pd.DataFrame) -> dict[str, int]:
    """
    Cuenta cuántos archivos existen por extensión.
    """
    if df_archivos.empty:
        return {}

    return df_archivos["Tipo"].value_counts().to_dict()


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.set_page_config(
    page_title="SNII Insight",
    page_icon="📊",
    layout="wide",
)

st.title("📊 SNII Insight")
st.subheader("Integración histórica del Sistema Nacional de Investigadoras e Investigadores")

st.write(
    "Esta primera versión se conecta con la carpeta pública de Google Drive, "
    "descarga temporalmente sus archivos y muestra cuáles fueron encontrados."
)

st.info(
    "La carpeta de Google Drive debe tener permiso de acceso: "
    "“Cualquier persona con el enlace”."
)

with st.expander("Repositorio configurado"):
    st.code(DRIVE_FOLDER_URL)

if st.button("🔄 Sincronizar archivos", type="primary", use_container_width=True):
    with st.spinner("Conectando con Google Drive y descargando archivos..."):
        try:
            with tempfile.TemporaryDirectory() as carpeta_temporal:
                archivos_descargados = descargar_carpeta_drive(
                    DRIVE_FOLDER_URL,
                    carpeta_temporal,
                )

                if not archivos_descargados:
                    st.error(
                        "No se descargaron archivos. Verifica que la carpeta sea pública "
                        "y que el enlace sea correcto."
                    )
                    st.stop()

                df_archivos = obtener_archivos_compatibles(carpeta_temporal)

                if df_archivos.empty:
                    st.warning(
                        "La carpeta fue consultada, pero no se encontraron archivos "
                        "Excel, PDF o CSV compatibles."
                    )
                    st.stop()

                conteos = contar_por_tipo(df_archivos)

                st.success("Sincronización terminada correctamente.")

                col1, col2, col3, col4 = st.columns(4)

                col1.metric("Archivos compatibles", len(df_archivos))
                col2.metric("Excel", conteos.get("XLSX", 0) + conteos.get("XLS", 0))
                col3.metric("PDF", conteos.get("PDF", 0))
                col4.metric("CSV", conteos.get("CSV", 0))

                st.dataframe(
                    df_archivos,
                    use_container_width=True,
                    hide_index=True,
                )

                st.caption(
                    "En esta etapa los archivos sólo se descargan temporalmente. "
                    "Todavía no se modifican ni se genera el archivo maestro."
                )

        except Exception as error:
            st.error("No fue posible completar la sincronización.")
            st.exception(error)
