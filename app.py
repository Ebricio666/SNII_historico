import re
import tempfile
from io import BytesIO
from pathlib import Path

import gdown
import pandas as pd
import streamlit as st

DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1UQ_sPApThDd3xHMQPz0GvMuIfLTcd_4x?usp=sharing"
)

EXTENSIONES_PERMITIDAS = {".xlsx", ".xls", ".pdf", ".csv"}


def descargar_carpeta_drive(url: str, destino: str) -> list[str]:
    archivos = gdown.download_folder(url=url, output=destino, quiet=True)
    return archivos or []


def obtener_archivos_compatibles(carpeta: str) -> pd.DataFrame:
    registros = []
    for ruta in Path(carpeta).rglob("*"):
        if ruta.is_file() and ruta.suffix.lower() in EXTENSIONES_PERMITIDAS:
            registros.append(
                {
                    "Archivo": ruta.name,
                    "Tipo": ruta.suffix[1:].upper(),
                    "Tamaño (MB)": round(ruta.stat().st_size / (1024 * 1024), 2),
                    "Ruta relativa": str(ruta.relative_to(carpeta)),
                    "Ruta completa": str(ruta),
                }
            )

    columnas = ["Archivo", "Tipo", "Tamaño (MB)", "Ruta relativa", "Ruta completa"]
    if not registros:
        return pd.DataFrame(columns=columnas)

    return (
        pd.DataFrame(registros, columns=columnas)
        .sort_values(["Tipo", "Archivo"])
        .reset_index(drop=True)
    )


def detectar_anio(texto: str):
    coincidencia = re.search(r"(20\d{2})", texto)
    return int(coincidencia.group(1)) if coincidencia else pd.NA


def limpiar_encabezados(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    columnas_limpias = []
    usados = {}

    for columna in df.columns:
        nombre = re.sub(r"\s+", " ", str(columna).strip())
        if not nombre or nombre.lower().startswith("unnamed"):
            nombre = "COLUMNA_SIN_NOMBRE"

        contador = usados.get(nombre, 0)
        usados[nombre] = contador + 1
        if contador:
            nombre = f"{nombre}_{contador + 1}"

        columnas_limpias.append(nombre)

    df.columns = columnas_limpias
    return df


def leer_excel_completo(ruta: Path):
    tablas = []
    incidencias = []

    try:
        hojas = pd.read_excel(ruta, sheet_name=None, dtype=object)
        for nombre_hoja, df in hojas.items():
            if df.empty:
                continue

            df = limpiar_encabezados(df)
            df.insert(0, "ORIGEN_HOJA", str(nombre_hoja))
            df.insert(0, "ORIGEN_ARCHIVO", ruta.name)

            if "ANIO" not in df.columns:
                df.insert(0, "ANIO", detectar_anio(ruta.name))

            tablas.append(df)

    except Exception as error:
        incidencias.append(
            {
                "Archivo": ruta.name,
                "Tipo": ruta.suffix.upper().replace(".", ""),
                "Estado": "No procesado",
                "Detalle": str(error),
            }
        )

    return tablas, incidencias


def leer_csv(ruta: Path):
    ultimo_error = None

    for codificacion in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            df = pd.read_csv(ruta, dtype=object, encoding=codificacion)
            df = limpiar_encabezados(df)
            df.insert(0, "ORIGEN_HOJA", "CSV")
            df.insert(0, "ORIGEN_ARCHIVO", ruta.name)

            if "ANIO" not in df.columns:
                df.insert(0, "ANIO", detectar_anio(ruta.name))

            return df, None

        except Exception as error:
            ultimo_error = error

    return None, {
        "Archivo": ruta.name,
        "Tipo": "CSV",
        "Estado": "No procesado",
        "Detalle": str(ultimo_error),
    }


def construir_master(df_archivos: pd.DataFrame):
    tablas = []
    incidencias = []

    for _, archivo in df_archivos.iterrows():
        ruta = Path(archivo["Ruta completa"])
        extension = ruta.suffix.lower()

        if extension in {".xlsx", ".xls"}:
            nuevas_tablas, nuevos_errores = leer_excel_completo(ruta)
            tablas.extend(nuevas_tablas)
            incidencias.extend(nuevos_errores)

        elif extension == ".csv":
            tabla, error = leer_csv(ruta)
            if tabla is not None:
                tablas.append(tabla)
            if error is not None:
                incidencias.append(error)

    if not tablas:
        raise ValueError("No fue posible leer ningún archivo Excel o CSV.")

    master = pd.concat(tablas, ignore_index=True, sort=False)

    columnas_iniciales = [
        columna
        for columna in ["ANIO", "ORIGEN_ARCHIVO", "ORIGEN_HOJA"]
        if columna in master.columns
    ]
    columnas_restantes = [
        columna for columna in master.columns if columna not in columnas_iniciales
    ]
    master = master[columnas_iniciales + columnas_restantes]

    pdfs = df_archivos[df_archivos["Tipo"] == "PDF"][
        ["Archivo", "Tamaño (MB)", "Ruta relativa"]
    ].copy()
    if not pdfs.empty:
        pdfs["ESTADO_EXTRACCION"] = "Pendiente de integrar extractor PDF"

    control_archivos = df_archivos[
        ["Archivo", "Tipo", "Tamaño (MB)", "Ruta relativa"]
    ].copy()

    incidencias_df = pd.DataFrame(
        incidencias,
        columns=["Archivo", "Tipo", "Estado", "Detalle"],
    )

    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="MASTER", index=False)
        control_archivos.to_excel(writer, sheet_name="CONTROL_ARCHIVOS", index=False)

        if not pdfs.empty:
            pdfs.to_excel(writer, sheet_name="CONTROL_PDF", index=False)

        if not incidencias_df.empty:
            incidencias_df.to_excel(writer, sheet_name="INCIDENCIAS", index=False)

    resumen = {
        "registros": len(master),
        "columnas": len(master.columns),
        "archivos_datos": master["ORIGEN_ARCHIVO"].nunique(),
        "pdf_pendientes": len(pdfs),
        "incidencias": len(incidencias_df),
    }

    return salida.getvalue(), resumen


st.set_page_config(page_title="SNII Insight", page_icon="📊", layout="wide")
st.title("📊 SNII Insight")
st.subheader("Integración histórica del Sistema Nacional de Investigadoras e Investigadores")
st.write(
    "Sincroniza el repositorio de Google Drive y genera un archivo maestro "
    "descargable para comenzar la visualización en HTML."
)
st.info("La carpeta debe estar compartida como “Cualquier persona con el enlace”.")

with st.expander("Repositorio configurado"):
    st.code(DRIVE_FOLDER_URL)

if "master_bytes" not in st.session_state:
    st.session_state.master_bytes = None
if "resumen_master" not in st.session_state:
    st.session_state.resumen_master = None

if st.button("🔄 Sincronizar y preparar MASTER", type="primary", use_container_width=True):
    st.session_state.master_bytes = None
    st.session_state.resumen_master = None

    with st.spinner("Descargando archivos y construyendo SNII_MASTER.xlsx..."):
        try:
            with tempfile.TemporaryDirectory() as carpeta_temporal:
                archivos_descargados = descargar_carpeta_drive(
                    DRIVE_FOLDER_URL,
                    carpeta_temporal,
                )

                if not archivos_descargados:
                    st.error(
                        "Google Drive no devolvió archivos. Verifica que la carpeta "
                        "sea pública y contenga archivos descargables."
                    )
                    st.stop()

                df_archivos = obtener_archivos_compatibles(carpeta_temporal)

                if df_archivos.empty:
                    st.warning(
                        "La carpeta se descargó, pero no se encontraron archivos "
                        "Excel, PDF o CSV."
                    )
                    st.stop()

                master_bytes, resumen = construir_master(df_archivos)
                st.session_state.master_bytes = master_bytes
                st.session_state.resumen_master = resumen

                conteos = df_archivos["Tipo"].value_counts().to_dict()
                st.success("El archivo maestro fue preparado correctamente.")

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Archivos", len(df_archivos))
                col2.metric("Excel", conteos.get("XLSX", 0) + conteos.get("XLS", 0))
                col3.metric("PDF", conteos.get("PDF", 0))
                col4.metric("CSV", conteos.get("CSV", 0))

                st.dataframe(
                    df_archivos.drop(columns=["Ruta completa"]),
                    use_container_width=True,
                    hide_index=True,
                )

        except Exception as error:
            st.error("No fue posible construir el archivo maestro.")
            st.exception(error)

if st.session_state.master_bytes is not None:
    resumen = st.session_state.resumen_master

    st.divider()
    st.subheader("📦 Archivo maestro listo")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Registros integrados", f"{resumen['registros']:,}")
    col2.metric("Columnas", resumen["columnas"])
    col3.metric("Archivos integrados", resumen["archivos_datos"])
    col4.metric("PDF pendientes", resumen["pdf_pendientes"])

    st.download_button(
        label="⬇️ Descargar SNII_MASTER.xlsx",
        data=st.session_state.master_bytes,
        file_name="SNII_MASTER.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    if resumen["pdf_pendientes"] > 0:
        st.warning(
            "Esta primera versión integra los archivos Excel y CSV. Los PDF ya "
            "se detectan y aparecen en CONTROL_PDF, pero sus registros todavía "
            "no se incorporan a MASTER."
        )

    if resumen["incidencias"] > 0:
        st.warning(
            f"Se registraron {resumen['incidencias']} incidencias de lectura. "
            "Revísalas en la hoja INCIDENCIAS."
        )
