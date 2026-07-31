import re
import tempfile
import unicodedata
from io import BytesIO
from pathlib import Path

import gdown
import pandas as pd
import streamlit as st
from pypdf import PdfReader


# ============================================================
# CONFIGURACIÓN
# ============================================================

DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1UQ_sPApThDd3xHMQPz0GvMuIfLTcd_4x?usp=sharing"
)

EXTENSIONES = {".xlsx", ".xls", ".csv", ".pdf"}
COLUMNAS_CONTROL = ["Archivo", "Tipo", "Tamaño (MB)", "Ruta relativa"]


# ============================================================
# UTILIDADES
# ============================================================

def normalizar_texto(valor) -> str:
    if valor is None or pd.isna(valor):
        return ""

    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip().upper()


def detectar_anio(texto: str):
    coincidencia = re.search(r"\b(20\d{2})\b", str(texto))
    return int(coincidencia.group(1)) if coincidencia else pd.NA


def detectar_origen_2025(nombre: str, texto: str = "") -> str:
    contenido = normalizar_texto(f"{nombre} {texto}")

    if "EMERIT" in contenido:
        return "Emérito"
    if "RECONSIDER" in contenido:
        return "Reconsideración"
    return "Primera ronda"


def limpiar_encabezados(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    nuevas = []
    repetidas = {}

    for columna in df.columns:
        nombre = re.sub(r"\s+", " ", str(columna).strip())

        if not nombre or nombre.lower().startswith("unnamed"):
            nombre = "COLUMNA_SIN_NOMBRE"

        numero = repetidas.get(nombre, 0) + 1
        repetidas[nombre] = numero

        nuevas.append(nombre if numero == 1 else f"{nombre}_{numero}")

    df.columns = nuevas
    return df


# ============================================================
# GOOGLE DRIVE
# ============================================================

def descargar_drive(destino: str) -> list[str]:
    archivos = gdown.download_folder(
        url=DRIVE_FOLDER_URL,
        output=destino,
        quiet=True,
    )
    return archivos or []


def listar_archivos(carpeta: str) -> pd.DataFrame:
    filas = []

    for ruta in Path(carpeta).rglob("*"):
        if not ruta.is_file() or ruta.suffix.lower() not in EXTENSIONES:
            continue

        filas.append(
            {
                "Archivo": ruta.name,
                "Tipo": ruta.suffix[1:].upper(),
                "Tamaño (MB)": round(ruta.stat().st_size / 1048576, 2),
                "Ruta relativa": str(ruta.relative_to(carpeta)),
                "Ruta completa": str(ruta),
            }
        )

    if not filas:
        return pd.DataFrame(
            columns=COLUMNAS_CONTROL + ["Ruta completa"]
        )

    return (
        pd.DataFrame(filas)
        .sort_values(["Tipo", "Archivo"])
        .reset_index(drop=True)
    )


# ============================================================
# EXCEL Y CSV
# ============================================================

def leer_excel(ruta: Path):
    tablas = []
    incidencias = []

    try:
        hojas = pd.read_excel(ruta, sheet_name=None, dtype=object)

        for hoja, df in hojas.items():
            if df.empty:
                continue

            df = limpiar_encabezados(df)
            df.insert(0, "ORIGEN_HOJA", str(hoja))
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


def leer_csv(ruta: Path):
    ultimo_error = None

    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            df = pd.read_csv(ruta, dtype=object, encoding=encoding)
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
# PDF: EXTRACCIÓN RÁPIDA POR TEXTO
# ============================================================

NIVEL_PATRON = (
    r"(?:"
    r"CANDIDAT[OA](?:\s+A\s+INVESTIGADOR(?:A)?\s+NACIONAL)?|"
    r"INVESTIGADOR(?:A)?\s+NACIONAL\s+EMERIT[OA]|"
    r"EMERIT[OA]|"
    r"NIVEL\s*(?:1|2|3|I|II|III)|"
    r"INVESTIGADOR(?:A)?\s+NACIONAL\s+(?:NIVEL\s*)?(?:1|2|3|I|II|III)"
    r")"
)


def limpiar_linea_pdf(linea: str) -> str:
    return re.sub(r"\s+", " ", str(linea or "")).strip()


def es_encabezado_pdf(linea: str) -> bool:
    texto = normalizar_texto(linea)
    claves = (
        "APELLIDO PATERNO",
        "APELLIDO MATERNO",
        "NOMBRES",
        "NIVEL OTORGADO",
        "RESULTADOS",
        "SISTEMA NACIONAL",
    )
    return any(clave in texto for clave in claves)


def separar_nombre(contenido: str):
    """
    Separación provisional:
    primera palabra = apellido paterno,
    segunda palabra = apellido materno,
    resto = nombres.

    Se conserva NOMBRE_COMPLETO_PDF para validar y corregir después.
    """
    partes = contenido.split()

    if len(partes) < 3:
        return "", "", contenido

    return partes[0], partes[1], " ".join(partes[2:])


def extraer_registros_lineas(lineas: list[str]) -> list[dict]:
    registros = []
    i = 0

    patron_unalinea = re.compile(
        rf"^(?P<cvu>\d{{4,12}})\s+"
        rf"(?P<nombre>.+?)\s+"
        rf"(?P<nivel>{NIVEL_PATRON})$",
        re.IGNORECASE,
    )

    while i < len(lineas):
        linea = limpiar_linea_pdf(lineas[i])

        if not linea or es_encabezado_pdf(linea):
            i += 1
            continue

        # Caso 1: todo el registro quedó en una sola línea.
        coincidencia = patron_unalinea.match(linea)

        if coincidencia:
            nombre_completo = coincidencia.group("nombre").strip()
            paterno, materno, nombres = separar_nombre(nombre_completo)

            registros.append(
                {
                    "CVU": coincidencia.group("cvu"),
                    "APELLIDO_PATERNO": paterno,
                    "APELLIDO_MATERNO": materno,
                    "NOMBRES": nombres,
                    "NOMBRE_COMPLETO_PDF": nombre_completo,
                    "NIVEL": coincidencia.group("nivel").strip(),
                }
            )
            i += 1
            continue

        # Caso 2: CVU en una línea y campos en líneas posteriores.
        if re.fullmatch(r"\d{4,12}", linea):
            bloque = []
            j = i + 1

            while j < len(lineas) and len(bloque) < 8:
                siguiente = limpiar_linea_pdf(lineas[j])

                if re.fullmatch(r"\d{4,12}", siguiente):
                    break

                if siguiente and not es_encabezado_pdf(siguiente):
                    bloque.append(siguiente)

                if re.fullmatch(NIVEL_PATRON, siguiente, re.IGNORECASE):
                    break

                j += 1

            if bloque:
                nivel = bloque[-1]

                if re.fullmatch(NIVEL_PATRON, nivel, re.IGNORECASE):
                    campos_nombre = bloque[:-1]

                    if len(campos_nombre) >= 3:
                        paterno = campos_nombre[0]
                        materno = campos_nombre[1]
                        nombres = " ".join(campos_nombre[2:])
                    else:
                        nombre_completo = " ".join(campos_nombre)
                        paterno, materno, nombres = separar_nombre(
                            nombre_completo
                        )

                    registros.append(
                        {
                            "CVU": linea,
                            "APELLIDO_PATERNO": paterno,
                            "APELLIDO_MATERNO": materno,
                            "NOMBRES": nombres,
                            "NOMBRE_COMPLETO_PDF": " ".join(campos_nombre),
                            "NIVEL": nivel,
                        }
                    )
                    i = j + 1
                    continue

        i += 1

    return registros


def extraer_pdf_pypdf(ruta: Path):
    """Método principal: rápido y sin análisis geométrico de tablas."""
    reader = PdfReader(str(ruta))
    registros = []
    texto_inicial = ""
    paginas_sin_registros = 0

    for numero, page in enumerate(reader.pages, start=1):
        texto = page.extract_text() or ""

        if numero <= 2:
            texto_inicial += " " + texto

        filas = extraer_registros_lineas(texto.splitlines())

        if not filas:
            paginas_sin_registros += 1

        for fila in filas:
            fila["PAGINA_PDF"] = numero
            registros.append(fila)

    return registros, texto_inicial, paginas_sin_registros


def extraer_pdf_pdfplumber_respaldo(ruta: Path):
    """
    Respaldo opcional. Se importa dentro de la función para que
    la aplicación no se caiga al iniciar si la dependencia falla.
    """
    try:
        import pdfplumber
    except ModuleNotFoundError:
        return [], "", 0

    registros = []
    texto_inicial = ""
    paginas_sin_registros = 0

    with pdfplumber.open(ruta) as pdf:
        for numero, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text(
                x_tolerance=2,
                y_tolerance=3,
            ) or ""

            if numero <= 2:
                texto_inicial += " " + texto

            filas = extraer_registros_lineas(texto.splitlines())

            if not filas:
                paginas_sin_registros += 1

            for fila in filas:
                fila["PAGINA_PDF"] = numero
                registros.append(fila)

    return registros, texto_inicial, paginas_sin_registros


def extraer_pdf(ruta: Path):
    incidencias = []

    try:
        registros, texto_inicial, paginas_vacias = extraer_pdf_pypdf(ruta)

        # Sólo usar el método más lento cuando PyPDF no obtuvo nada.
        if not registros:
            registros, texto_inicial, paginas_vacias = (
                extraer_pdf_pdfplumber_respaldo(ruta)
            )

        if not registros:
            incidencias.append(
                {
                    "Archivo": ruta.name,
                    "Tipo": "PDF",
                    "Estado": "Sin registros",
                    "Detalle": (
                        "El PDF abrió correctamente, pero no se reconocieron "
                        "registros. Requiere ajuste específico del formato."
                    ),
                }
            )
            return pd.DataFrame(), incidencias

        df = pd.DataFrame(registros).drop_duplicates(
            subset=["CVU", "NOMBRE_COMPLETO_PDF", "NIVEL"]
        )

        df.insert(0, "ANIO", 2025)
        df.insert(1, "ORIGEN_ARCHIVO", ruta.name)
        df.insert(2, "ORIGEN_HOJA", "PDF")
        df.insert(
            3,
            "ORIGEN_2025",
            detectar_origen_2025(ruta.name, texto_inicial),
        )

        incidencias.append(
            {
                "Archivo": ruta.name,
                "Tipo": "PDF",
                "Estado": "Procesado",
                "Detalle": (
                    f"{len(df)} registros extraídos; "
                    f"{paginas_vacias} páginas sin registros."
                ),
            }
        )

        return df.reset_index(drop=True), incidencias

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


# ============================================================
# MASTER
# ============================================================

def crear_salidas(master, control, incidencias):
    excel = BytesIO()

    with pd.ExcelWriter(excel, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="MASTER", index=False)
        control.to_excel(writer, sheet_name="CONTROL_ARCHIVOS", index=False)

        if not incidencias.empty:
            incidencias.to_excel(
                writer,
                sheet_name="INCIDENCIAS",
                index=False,
            )

    csv = master.to_csv(index=False).encode("utf-8-sig")

    parquet = BytesIO()
    master.to_parquet(parquet, index=False)

    return excel.getvalue(), csv, parquet.getvalue()


@st.cache_data(ttl=3600, show_spinner=False)
def sincronizar_y_procesar():
    """
    Cachea durante una hora la descarga y el procesamiento.
    Al volver a presionar el botón, Streamlit reutiliza el resultado.
    """
    with tempfile.TemporaryDirectory() as carpeta:
        descargados = descargar_drive(carpeta)

        if not descargados:
            raise RuntimeError(
                "Google Drive no devolvió archivos. Verifica los permisos."
            )

        archivos = listar_archivos(carpeta)

        if archivos.empty:
            raise RuntimeError(
                "No se encontraron archivos Excel, CSV o PDF."
            )

        tablas = []
        incidencias = []

        for _, fila in archivos.iterrows():
            ruta = Path(fila["Ruta completa"])
            extension = ruta.suffix.lower()

            if extension in {".xlsx", ".xls"}:
                nuevas, errores = leer_excel(ruta)
                tablas.extend(nuevas)
                incidencias.extend(errores)

            elif extension == ".csv":
                tabla, error = leer_csv(ruta)

                if tabla is not None:
                    tablas.append(tabla)
                if error:
                    incidencias.append(error)

            elif extension == ".pdf":
                tabla, mensajes = extraer_pdf(ruta)
                incidencias.extend(mensajes)

                if not tabla.empty:
                    tablas.append(tabla)

        if not tablas:
            raise RuntimeError(
                "No fue posible integrar ningún registro."
            )

        master = pd.concat(tablas, ignore_index=True, sort=False)

        prioridad = [
            "ANIO",
            "ORIGEN_ARCHIVO",
            "ORIGEN_HOJA",
            "ORIGEN_2025",
            "PAGINA_PDF",
        ]

        primeras = [c for c in prioridad if c in master.columns]
        restantes = [c for c in master.columns if c not in primeras]
        master = master[primeras + restantes]

        control = archivos[COLUMNAS_CONTROL].copy()
        incidencias_df = pd.DataFrame(
            incidencias,
            columns=["Archivo", "Tipo", "Estado", "Detalle"],
        )

        excel, csv, parquet = crear_salidas(
            master,
            control,
            incidencias_df,
        )

        resumen = {
            "registros": len(master),
            "archivos": master["ORIGEN_ARCHIVO"].nunique(),
            "pdf": int(
                master["ORIGEN_HOJA"].eq("PDF").sum()
                if "ORIGEN_HOJA" in master.columns
                else 0
            ),
            "incidencias": len(incidencias_df),
        }

        return (
            excel,
            csv,
            parquet,
            resumen,
            control,
            incidencias_df,
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
st.caption("Integración histórica y preparación de datos para HTML")

st.info(
    "La primera ejecución puede tardar porque descarga y procesa todo. "
    "Durante la siguiente hora, el resultado queda en caché."
)

col_sync, col_clear = st.columns([3, 1])

with col_sync:
    ejecutar = st.button(
        "🔄 Sincronizar y generar MASTER",
        type="primary",
        use_container_width=True,
    )

with col_clear:
    if st.button("🧹 Limpiar caché", use_container_width=True):
        st.cache_data.clear()
        st.session_state.clear()
        st.success("Caché eliminado.")

if ejecutar:
    barra = st.progress(10, text="Conectando con Google Drive...")

    try:
        barra.progress(35, text="Leyendo Excel y CSV...")
        resultado = sincronizar_y_procesar()
        barra.progress(90, text="Preparando archivos de descarga...")

        (
            excel_bytes,
            csv_bytes,
            parquet_bytes,
            resumen,
            control,
            incidencias,
        ) = resultado

        st.session_state["resultado"] = resultado
        barra.progress(100, text="Proceso terminado.")

    except Exception as error:
        barra.empty()
        st.error("No fue posible generar el archivo maestro.")
        st.exception(error)

if "resultado" in st.session_state:
    (
        excel_bytes,
        csv_bytes,
        parquet_bytes,
        resumen,
        control,
        incidencias,
    ) = st.session_state["resultado"]

    st.success("Archivos preparados correctamente.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros", f"{resumen['registros']:,}")
    c2.metric("Archivos integrados", resumen["archivos"])
    c3.metric("Registros PDF", f"{resumen['pdf']:,}")
    c4.metric("Incidencias", resumen["incidencias"])

    with st.expander("Ver archivos sincronizados"):
        st.dataframe(control, use_container_width=True, hide_index=True)

    d1, d2, d3 = st.columns(3)

    with d1:
        st.download_button(
            "⬇️ Descargar Excel",
            data=excel_bytes,
            file_name="SNII_MASTER.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            type="primary",
        )

    with d2:
        st.download_button(
            "⬇️ Descargar CSV",
            data=csv_bytes,
            file_name="SNII_MASTER.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with d3:
        st.download_button(
            "⬇️ Descargar Parquet",
            data=parquet_bytes,
            file_name="SNII_MASTER.parquet",
            mime="application/octet-stream",
            use_container_width=True,
        )

    if not incidencias.empty:
        with st.expander("Ver incidencias"):
            st.dataframe(
                incidencias,
                use_container_width=True,
                hide_index=True,
            )

    st.caption(
        "Para el dashboard HTML conviene usar Parquet o CSV. "
        "Parquet carga más rápido y ocupa menos espacio."
    )
