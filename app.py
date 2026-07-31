import re
import tempfile
import unicodedata
from io import BytesIO
from pathlib import Path

import gdown
import pandas as pd
import pdfplumber
import streamlit as st


# ============================================================
# CONFIGURACIÓN
# ============================================================

DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1UQ_sPApThDd3xHMQPz0GvMuIfLTcd_4x?usp=sharing"
)

EXTENSIONES_PERMITIDAS = {".xlsx", ".xls", ".pdf", ".csv"}

COLUMNAS_PDF_2025 = [
    "CVU",
    "APELLIDO_PATERNO",
    "APELLIDO_MATERNO",
    "NOMBRES",
    "NIVEL",
]


# ============================================================
# UTILIDADES
# ============================================================

def normalizar_texto(valor) -> str:
    """Convierte texto a mayúsculas, sin acentos y con espacios simples."""
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )
    texto = re.sub(r"\s+", " ", texto)
    return texto.upper().strip()


def detectar_anio(texto: str):
    """Busca un año entre 2000 y 2099 dentro de un texto."""
    coincidencia = re.search(r"(20\d{2})", texto)
    return int(coincidencia.group(1)) if coincidencia else pd.NA


def detectar_origen_pdf(nombre_archivo: str, texto_inicial: str = "") -> str:
    """Clasifica el PDF por nombre o contenido."""
    texto = normalizar_texto(f"{nombre_archivo} {texto_inicial}")

    if "EMERIT" in texto:
        return "Emérito"
    if "RECONSIDER" in texto:
        return "Reconsideración"
    return "Primera ronda"


def limpiar_encabezados(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia encabezados sin modificar los valores originales."""
    df = df.copy()
    columnas = []
    usados = {}

    for columna in df.columns:
        nombre = str(columna).strip()
        nombre = re.sub(r"\s+", " ", nombre)

        if not nombre or nombre.lower().startswith("unnamed"):
            nombre = "COLUMNA_SIN_NOMBRE"

        repeticion = usados.get(nombre, 0)
        usados[nombre] = repeticion + 1

        if repeticion:
            nombre = f"{nombre}_{repeticion + 1}"

        columnas.append(nombre)

    df.columns = columnas
    return df


def descargar_carpeta_drive(url: str, destino: str) -> list[str]:
    """Descarga una carpeta pública de Google Drive."""
    archivos = gdown.download_folder(
        url=url,
        output=destino,
        quiet=True,
    )
    return archivos or []


def obtener_archivos_compatibles(carpeta: str) -> pd.DataFrame:
    """Lista archivos compatibles descargados."""
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
                "Ruta completa": str(ruta),
            }
        )

    if not registros:
        return pd.DataFrame(
            columns=[
                "Archivo",
                "Tipo",
                "Tamaño (MB)",
                "Ruta relativa",
                "Ruta completa",
            ]
        )

    return (
        pd.DataFrame(registros)
        .sort_values(["Tipo", "Archivo"])
        .reset_index(drop=True)
    )


# ============================================================
# LECTURA DE EXCEL Y CSV
# ============================================================

def leer_excel_completo(ruta: Path) -> tuple[list[pd.DataFrame], list[dict]]:
    """Lee todas las hojas con información de un archivo Excel."""
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
                "Tipo": "EXCEL",
                "Estado": "No procesado",
                "Detalle": str(error),
            }
        )

    return tablas, incidencias


def leer_csv(ruta: Path) -> tuple[pd.DataFrame | None, dict | None]:
    """Lee CSV con varias codificaciones posibles."""
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


# ============================================================
# EXTRACCIÓN DE PDF 2025
# ============================================================

def fila_pdf_es_valida(valores: list[str]) -> bool:
    """Valida que una fila parezca un registro real del SNII."""
    if len(valores) < 5:
        return False

    cvu = re.sub(r"\D", "", valores[0])

    if not cvu:
        return False

    texto_fila = normalizar_texto(" ".join(valores))

    encabezados = (
        "APELLIDO PATERNO",
        "APELLIDO MATERNO",
        "NOMBRES",
        "NIVEL OTORGADO",
    )

    return not any(encabezado in texto_fila for encabezado in encabezados)


def convertir_fila_pdf(valores: list[str]) -> dict | None:
    """
    Convierte una fila de tabla PDF al formato estándar 2025.

    El formato esperado es:
    CVU | Apellido paterno | Apellido materno | Nombres | Nivel otorgado
    """
    valores = [
        re.sub(r"\s+", " ", str(valor or "")).strip()
        for valor in valores
    ]

    if not fila_pdf_es_valida(valores):
        return None

    return {
        "CVU": re.sub(r"\D", "", valores[0]),
        "APELLIDO_PATERNO": valores[1],
        "APELLIDO_MATERNO": valores[2],
        "NOMBRES": valores[3],
        "NIVEL": valores[4],
    }


def extraer_registros_desde_tablas(
    pagina,
) -> list[dict]:
    """Extrae registros usando las tablas detectadas por pdfplumber."""
    registros = []

    configuraciones = [
        {},
        {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
        },
        {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "intersection_tolerance": 5,
            "snap_tolerance": 4,
            "join_tolerance": 4,
        },
    ]

    for configuracion in configuraciones:
        try:
            tablas = pagina.extract_tables(
                table_settings=configuracion or None
            )
        except Exception:
            continue

        for tabla in tablas or []:
            for fila in tabla or []:
                registro = convertir_fila_pdf(fila or [])
                if registro:
                    registros.append(registro)

        if registros:
            break

    return registros


def extraer_registros_desde_texto(texto: str) -> list[dict]:
    """
    Método alternativo para PDFs donde las tablas no son reconocidas.

    Busca líneas que comienzan con CVU y terminan con un nivel SNII.
    """
    registros = []

    niveles = (
        r"(?:CANDIDAT[OA]|NIVEL\s*[123I]{1,3}|"
        r"INVESTIGADOR(?:A)?\s+NACIONAL\s+(?:NIVEL\s*)?[123I]{1,3}|"
        r"EMERIT[OA])"
    )

    for linea in (texto or "").splitlines():
        linea = re.sub(r"\s+", " ", linea).strip()

        patron = re.compile(
            rf"^(?P<cvu>\d{{4,10}})\s+"
            rf"(?P<contenido>.+?)\s+"
            rf"(?P<nivel>{niveles})$",
            flags=re.IGNORECASE,
        )

        coincidencia = patron.match(linea)
        if not coincidencia:
            continue

        contenido = coincidencia.group("contenido").strip()
        partes = contenido.split()

        # Este método es de respaldo. Se requieren al menos:
        # apellido paterno, apellido materno y un nombre.
        if len(partes) < 3:
            continue

        registros.append(
            {
                "CVU": coincidencia.group("cvu"),
                "APELLIDO_PATERNO": partes[0],
                "APELLIDO_MATERNO": partes[1],
                "NOMBRES": " ".join(partes[2:]),
                "NIVEL": coincidencia.group("nivel"),
            }
        )

    return registros


def extraer_pdf_2025(ruta: Path) -> tuple[pd.DataFrame, list[dict]]:
    """Extrae los registros de un PDF de resultados 2025."""
    registros = []
    incidencias = []
    texto_inicial = ""

    try:
        with pdfplumber.open(ruta) as pdf:
            total_paginas = len(pdf.pages)

            for numero_pagina, pagina in enumerate(pdf.pages, start=1):
                texto = pagina.extract_text() or ""

                if numero_pagina <= 2:
                    texto_inicial += f" {texto}"

                registros_pagina = extraer_registros_desde_tablas(pagina)

                if not registros_pagina:
                    registros_pagina = extraer_registros_desde_texto(texto)

                for registro in registros_pagina:
                    registro["PAGINA_PDF"] = numero_pagina
                    registros.append(registro)

                if not registros_pagina:
                    incidencias.append(
                        {
                            "Archivo": ruta.name,
                            "Tipo": "PDF",
                            "Estado": "Página sin registros",
                            "Detalle": (
                                f"No se identificaron registros en la página "
                                f"{numero_pagina} de {total_paginas}."
                            ),
                        }
                    )

    except Exception as error:
        incidencias.append(
            {
                "Archivo": ruta.name,
                "Tipo": "PDF",
                "Estado": "No procesado",
                "Detalle": str(error),
            }
        )
        return pd.DataFrame(), incidencias

    if not registros:
        incidencias.append(
            {
                "Archivo": ruta.name,
                "Tipo": "PDF",
                "Estado": "Sin registros",
                "Detalle": (
                    "El PDF fue abierto, pero no se pudo reconocer su tabla. "
                    "Puede requerir un ajuste específico del extractor."
                ),
            }
        )
        return pd.DataFrame(), incidencias

    origen = detectar_origen_pdf(ruta.name, texto_inicial)

    df = pd.DataFrame(registros)
    df = df.drop_duplicates(
        subset=["CVU", "APELLIDO_PATERNO", "APELLIDO_MATERNO", "NOMBRES", "NIVEL"]
    )

    df.insert(0, "ANIO", 2025)
    df.insert(1, "ORIGEN_ARCHIVO", ruta.name)
    df.insert(2, "ORIGEN_HOJA", "PDF")
    df.insert(3, "ORIGEN_2025", origen)

    return df.reset_index(drop=True), incidencias


# ============================================================
# GENERACIÓN DE ARCHIVOS DE SALIDA
# ============================================================

def construir_master(
    df_archivos: pd.DataFrame,
) -> tuple[bytes, bytes, dict]:
    """Integra Excel, CSV y PDF en un archivo maestro."""
    tablas = []
    incidencias = []
    pdf_extraidos = 0
    registros_pdf = 0

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

        elif extension == ".pdf":
            tabla_pdf, errores_pdf = extraer_pdf_2025(ruta)
            incidencias.extend(errores_pdf)

            if not tabla_pdf.empty:
                tablas.append(tabla_pdf)
                pdf_extraidos += 1
                registros_pdf += len(tabla_pdf)

    if not tablas:
        raise ValueError(
            "No fue posible leer datos de los archivos sincronizados."
        )

    master = pd.concat(tablas, ignore_index=True, sort=False)

    columnas_iniciales = [
        columna
        for columna in [
            "ANIO",
            "ORIGEN_ARCHIVO",
            "ORIGEN_HOJA",
            "ORIGEN_2025",
            "PAGINA_PDF",
        ]
        if columna in master.columns
    ]

    master = master[
        columnas_iniciales
        + [
            columna
            for columna in master.columns
            if columna not in columnas_iniciales
        ]
    ]

    control_archivos = df_archivos[
        ["Archivo", "Tipo", "Tamaño (MB)", "Ruta relativa"]
    ].copy()

    incidencias_df = pd.DataFrame(
        incidencias,
        columns=["Archivo", "Tipo", "Estado", "Detalle"],
    )

    salida_excel = BytesIO()

    with pd.ExcelWriter(salida_excel, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="MASTER", index=False)
        control_archivos.to_excel(
            writer,
            sheet_name="CONTROL_ARCHIVOS",
            index=False,
        )

        if not incidencias_df.empty:
            incidencias_df.to_excel(
                writer,
                sheet_name="INCIDENCIAS",
                index=False,
            )

    # CSV para alimentar el HTML con mayor facilidad.
    salida_csv = master.to_csv(index=False).encode("utf-8-sig")

    resumen = {
        "registros": len(master),
        "columnas": len(master.columns),
        "archivos_integrados": master["ORIGEN_ARCHIVO"].nunique(),
        "pdf_extraidos": pdf_extraidos,
        "registros_pdf": registros_pdf,
        "incidencias": len(incidencias_df),
    }

    return salida_excel.getvalue(), salida_csv, resumen


# ============================================================
# INTERFAZ STREAMLIT
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
    "Sincroniza el repositorio, integra los Excel históricos, "
    "extrae los resultados publicados en PDF y genera archivos "
    "listos para comenzar el dashboard HTML."
)

st.info(
    "La carpeta de Google Drive debe permitir el acceso mediante "
    "“Cualquier persona con el enlace”."
)

with st.expander("Repositorio configurado"):
    st.code(DRIVE_FOLDER_URL)

for clave in ("master_excel", "master_csv", "resumen_master"):
    if clave not in st.session_state:
        st.session_state[clave] = None

if st.button(
    "🔄 Sincronizar y generar MASTER",
    type="primary",
    use_container_width=True,
):
    st.session_state.master_excel = None
    st.session_state.master_csv = None
    st.session_state.resumen_master = None

    with st.spinner(
        "Descargando archivos, leyendo PDF y construyendo el MASTER..."
    ):
        try:
            with tempfile.TemporaryDirectory() as carpeta_temporal:
                archivos_descargados = descargar_carpeta_drive(
                    DRIVE_FOLDER_URL,
                    carpeta_temporal,
                )

                if not archivos_descargados:
                    st.error(
                        "Google Drive no devolvió archivos. Verifica el enlace "
                        "y los permisos de la carpeta."
                    )
                    st.stop()

                df_archivos = obtener_archivos_compatibles(
                    carpeta_temporal
                )

                if df_archivos.empty:
                    st.warning(
                        "No se encontraron archivos Excel, CSV o PDF."
                    )
                    st.stop()

                excel_bytes, csv_bytes, resumen = construir_master(
                    df_archivos
                )

                st.session_state.master_excel = excel_bytes
                st.session_state.master_csv = csv_bytes
                st.session_state.resumen_master = resumen

                st.success("El archivo maestro fue generado correctamente.")

                st.dataframe(
                    df_archivos.drop(columns=["Ruta completa"]),
                    use_container_width=True,
                    hide_index=True,
                )

        except Exception as error:
            st.error("No fue posible construir el archivo maestro.")
            st.exception(error)

if st.session_state.master_excel is not None:
    resumen = st.session_state.resumen_master

    st.divider()
    st.subheader("📦 Archivos listos")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Registros", f"{resumen['registros']:,}")
    col2.metric("PDF procesados", resumen["pdf_extraidos"])
    col3.metric("Registros de PDF", f"{resumen['registros_pdf']:,}")
    col4.metric("Incidencias", resumen["incidencias"])

    col_excel, col_csv = st.columns(2)

    with col_excel:
        st.download_button(
            label="⬇️ Descargar SNII_MASTER.xlsx",
            data=st.session_state.master_excel,
            file_name="SNII_MASTER.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )

    with col_csv:
        st.download_button(
            label="⬇️ Descargar SNII_MASTER.csv",
            data=st.session_state.master_csv,
            file_name="SNII_MASTER.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.caption(
        "El CSV es la salida más sencilla para comenzar el dashboard HTML. "
        "El Excel conserva además las hojas CONTROL_ARCHIVOS e INCIDENCIAS."
    )

    if resumen["incidencias"] > 0:
        st.warning(
            "Revisa la hoja INCIDENCIAS. Las páginas de portada, avisos o "
            "instrucciones pueden aparecer como páginas sin registros; eso "
            "no necesariamente significa que la extracción haya fallado."
        )
