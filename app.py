import tempfile
from pathlib import Path

import gdown
import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN
# ============================================================

DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1UQ_sPApThDd3xHMQPz0GvMuIfLTcd_4x?usp=sharing"
)

EXTENSIONES_PERMITIDAS = {".xlsx", ".xls", ".pdf", ".csv"}


# ============================================================
# FUNCIONES
# ============================================================

def descargar_carpeta_drive(url: str, destino: str) -> list[str]:
    """
    Descarga una carpeta pública de Google Drive.

    Se usan únicamente los argumentos compatibles con las versiones
    actuales de gdown.
    """
    archivos = gdown.download_folder(
        url=url,
        output=destino,
        quiet=True,
    )

    return archivos or []


def obtener_archivos_compatibles(carpeta: str) -> pd.DataFrame:
    """Obtiene la lista de archivos Excel, PDF y CSV descargados."""
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
                "Tipo": extension[1:].upper(),
                "Tamaño (MB)": round(ruta.stat().st_size / (1024 * 1024), 2),
                "Ruta relativa": str(ruta.relative_to(carpeta)),
            }
        )

    columnas = ["Archivo", "Tipo", "Tamaño (MB)", "Ruta relativa"]

    if not registros:
        return pd.DataFrame(columns=columnas)

    return (
        pd.DataFrame(registros, columns=columnas)
        .sort_values(["Tipo", "Archivo"])
        .reset_index(drop=True)
    )


# ============================================================
# INTERFAZ
# ============================================================

st.set_page_config(
    page_title="SNII Insight",
    page_icon="📊",
    layout="wide",
)

st.title("📊 SNII Insight")
st.subheader(
    "Integración histórica del Sistema Nacional de "
    "Investigadoras e Investigadores"
)

st.write(
    "Esta primera versión consulta la carpeta pública de Google Drive "
    "y muestra los archivos compatibles encontrados."
)

st.info(
    "La carpeta debe estar compartida como "
    "“Cualquier persona con el enlace”."
)

with st.expander("Repositorio configurado"):
    st.code(DRIVE_FOLDER_URL)

if st.button(
    "🔄 Sincronizar archivos",
    type="primary",
    use_container_width=True,
):
    with st.spinner("Descargando archivos desde Google Drive..."):
        try:
            with tempfile.TemporaryDirectory() as carpeta_temporal:
                archivos_descargados = descargar_carpeta_drive(
                    DRIVE_FOLDER_URL,
                    carpeta_temporal,
                )

                if not archivos_descargados:
                    st.error(
                        "Google Drive no devolvió archivos. Verifica que la "
                        "carpeta sea pública y que contenga archivos descargables."
                    )
                    st.stop()

                df_archivos = obtener_archivos_compatibles(
                    carpeta_temporal
                )

                if df_archivos.empty:
                    st.warning(
                        "La carpeta se descargó, pero no se encontraron "
                        "archivos Excel, PDF o CSV."
                    )
                    st.stop()

                conteos = df_archivos["Tipo"].value_counts().to_dict()

                st.success("Sincronización terminada correctamente.")

                col1, col2, col3, col4 = st.columns(4)

                col1.metric("Archivos", len(df_archivos))
                col2.metric(
                    "Excel",
                    conteos.get("XLSX", 0) + conteos.get("XLS", 0),
                )
                col3.metric("PDF", conteos.get("PDF", 0))
                col4.metric("CSV", conteos.get("CSV", 0))

                st.dataframe(
                    df_archivos,
                    use_container_width=True,
                    hide_index=True,
                )

                st.caption(
                    "Los archivos se descargan temporalmente y se eliminan "
                    "al terminar esta ejecución."
                )

        except TypeError as error:
            st.error(
                "La versión instalada de gdown no es compatible con el código."
            )
            st.code(str(error))
            st.write(
                "Sustituye también el archivo requirements.txt por la "
                "versión corregida."
            )

        except Exception as error:
            st.error("No fue posible sincronizar la carpeta.")
            st.exception(error)
