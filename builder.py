"""
SNII Builder
============

Procesa los archivos históricos desde Google Drive y genera:

- data/SNII_MASTER.parquet
- data/SNII_MASTER.csv
- data/SNII_MASTER.xlsx
- data/CONTROL_ARCHIVOS.csv
- data/INCIDENCIAS.csv

Este script NO se ejecuta al abrir el dashboard.
Se ejecuta solamente cuando cambian los archivos fuente.

Uso local:
    python builder.py
"""

from __future__ import annotations

import re
import shutil
import tempfile
import unicodedata
from pathlib import Path
from typing import Any

import gdown
import pandas as pd
from pypdf import PdfReader


# ============================================================
# CONFIGURACIÓN
# ============================================================

DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1UQ_sPApThDd3xHMQPz0GvMuIfLTcd_4x?usp=sharing"
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

EXTENSIONES_PERMITIDAS = {".xlsx", ".xls", ".csv", ".pdf"}

COLUMNAS_CONTROL = [
    "archivo",
    "extension",
    "tamano_mb",
    "ruta_relativa",
]


# ============================================================
# UTILIDADES GENERALES
# ============================================================

def normalizar_texto(valor: Any) -> str:
    """Convierte un valor a texto limpio, sin acentos y en mayúsculas."""
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip().upper()


def detectar_anio(texto: str):
    """Busca un año entre 2000 y 2099 dentro del texto."""
    coincidencia = re.search(r"\b(20\d{2})\b", str(texto))
    return int(coincidencia.group(1)) if coincidencia else pd.NA


def detectar_origen_2025(nombre_archivo: str, texto: str = "") -> str:
    contenido = normalizar_texto(f"{nombre_archivo} {texto}")

    if "EMERIT" in contenido:
        return "Emérito"

    if "RECONSIDER" in contenido:
        return "Reconsideración"

    return "Primera ronda"


def limpiar_nombre_columna(nombre: Any) -> str:
    texto = normalizar_texto(nombre)
    texto = re.sub(r"[^A-Z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto or "COLUMNA_SIN_NOMBRE"


def limpiar_encabezados(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    columnas = []
    conteo: dict[str, int] = {}

    for columna in df.columns:
        nombre = limpiar_nombre_columna(columna)
        conteo[nombre] = conteo.get(nombre, 0) + 1

        if conteo[nombre] > 1:
            nombre = f"{nombre}_{conteo[nombre]}"

        columnas.append(nombre)

    df.columns = columnas
    return df


def asegurar_directorio_salida() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# DESCARGA DE GOOGLE DRIVE
# ============================================================

def descargar_carpeta_drive(destino: Path) -> list[str]:
    """
    Descarga la carpeta pública de Google Drive.
    """
    archivos = gdown.download_folder(
        url=DRIVE_FOLDER_URL,
        output=str(destino),
        quiet=False,
    )

    return archivos or []


def listar_archivos(carpeta: Path) -> pd.DataFrame:
    filas = []

    for ruta in carpeta.rglob("*"):
        if not ruta.is_file():
            continue

        extension = ruta.suffix.lower()

        if extension not in EXTENSIONES_PERMITIDAS:
            continue

        filas.append(
            {
                "archivo": ruta.name,
                "extension": extension.lstrip(".").upper(),
                "tamano_mb": round(ruta.stat().st_size / 1_048_576, 3),
                "ruta_relativa": str(ruta.relative_to(carpeta)),
                "ruta_completa": str(ruta),
            }
        )

    if not filas:
        return pd.DataFrame(
            columns=COLUMNAS_CONTROL + ["ruta_completa"]
        )

    return (
        pd.DataFrame(filas)
        .sort_values(["extension", "archivo"])
        .reset_index(drop=True)
    )


# ============================================================
# LECTURA DE EXCEL Y CSV
# ============================================================

def leer_excel(ruta: Path):
    tablas: list[pd.DataFrame] = []
    incidencias: list[dict[str, Any]] = []

    try:
        hojas = pd.read_excel(
            ruta,
            sheet_name=None,
            dtype=object,
        )

        for hoja, df in hojas.items():
            if df is None or df.empty:
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
                "archivo": ruta.name,
                "tipo": "EXCEL",
                "estado": "NO_PROCESADO",
                "detalle": repr(error),
            }
        )

    return tablas, incidencias


def leer_csv(ruta: Path):
    ultimo_error = None

    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            df = pd.read_csv(
                ruta,
                dtype=object,
                encoding=encoding,
            )

            df = limpiar_encabezados(df)
            df.insert(0, "ORIGEN_HOJA", "CSV")
            df.insert(0, "ORIGEN_ARCHIVO", ruta.name)

            if "ANIO" not in df.columns:
                df.insert(0, "ANIO", detectar_anio(ruta.name))

            return df, None

        except Exception as error:
            ultimo_error = error

    return None, {
        "archivo": ruta.name,
        "tipo": "CSV",
        "estado": "NO_PROCESADO",
        "detalle": repr(ultimo_error),
    }


# ============================================================
# PDF 2025: EXTRACCIÓN POR TEXTO
# ============================================================

PATRON_NIVEL = (
    r"(?:"
    r"CANDIDAT[OA](?:\s+A\s+INVESTIGADOR(?:A)?\s+NACIONAL)?|"
    r"INVESTIGADOR(?:A)?\s+NACIONAL\s+EMERIT[OA]|"
    r"EMERIT[OA]|"
    r"NIVEL\s*(?:1|2|3|I|II|III)|"
    r"INVESTIGADOR(?:A)?\s+NACIONAL\s+"
    r"(?:NIVEL\s*)?(?:1|2|3|I|II|III)"
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
        "CONAHCYT",
        "CONACYT",
    )

    return any(clave in texto for clave in claves)


def separar_nombre_provisional(nombre_completo: str):
    """
    Separación provisional:
    - primera palabra: apellido paterno
    - segunda palabra: apellido materno
    - resto: nombres

    Se conserva NOMBRE_COMPLETO_PDF para futuras correcciones.
    """
    partes = nombre_completo.split()

    if len(partes) < 3:
        return "", "", nombre_completo

    return partes[0], partes[1], " ".join(partes[2:])


def extraer_registros_desde_lineas(lineas: list[str]) -> list[dict[str, Any]]:
    registros: list[dict[str, Any]] = []
    indice = 0

    patron_linea_completa = re.compile(
        rf"^(?P<cvu>\d{{4,12}})\s+"
        rf"(?P<nombre>.+?)\s+"
        rf"(?P<nivel>{PATRON_NIVEL})$",
        flags=re.IGNORECASE,
    )

    while indice < len(lineas):
        linea = limpiar_linea_pdf(lineas[indice])

        if not linea or es_encabezado_pdf(linea):
            indice += 1
            continue

        # Caso 1: registro en una sola línea.
        coincidencia = patron_linea_completa.match(linea)

        if coincidencia:
            nombre_completo = coincidencia.group("nombre").strip()

            paterno, materno, nombres = separar_nombre_provisional(
                nombre_completo
            )

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

            indice += 1
            continue

        # Caso 2: CVU en línea independiente.
        if re.fullmatch(r"\d{4,12}", linea):
            bloque: list[str] = []
            siguiente_indice = indice + 1

            while siguiente_indice < len(lineas) and len(bloque) < 10:
                siguiente = limpiar_linea_pdf(
                    lineas[siguiente_indice]
                )

                if re.fullmatch(r"\d{4,12}", siguiente):
                    break

                if siguiente and not es_encabezado_pdf(siguiente):
                    bloque.append(siguiente)

                if re.fullmatch(
                    PATRON_NIVEL,
                    siguiente,
                    flags=re.IGNORECASE,
                ):
                    break

                siguiente_indice += 1

            if bloque:
                nivel = bloque[-1]

                if re.fullmatch(
                    PATRON_NIVEL,
                    nivel,
                    flags=re.IGNORECASE,
                ):
                    campos_nombre = bloque[:-1]

                    if len(campos_nombre) >= 3:
                        paterno = campos_nombre[0]
                        materno = campos_nombre[1]
                        nombres = " ".join(campos_nombre[2:])
                    else:
                        nombre_completo = " ".join(campos_nombre)

                        (
                            paterno,
                            materno,
                            nombres,
                        ) = separar_nombre_provisional(nombre_completo)

                    registros.append(
                        {
                            "CVU": linea,
                            "APELLIDO_PATERNO": paterno,
                            "APELLIDO_MATERNO": materno,
                            "NOMBRES": nombres,
                            "NOMBRE_COMPLETO_PDF": " ".join(
                                campos_nombre
                            ),
                            "NIVEL": nivel,
                        }
                    )

                    indice = siguiente_indice + 1
                    continue

        indice += 1

    return registros


def extraer_pdf(ruta: Path):
    incidencias: list[dict[str, Any]] = []

    try:
        lector = PdfReader(str(ruta))
        registros: list[dict[str, Any]] = []
        texto_inicial = ""
        paginas_sin_registros = 0

        for numero_pagina, pagina in enumerate(
            lector.pages,
            start=1,
        ):
            texto = pagina.extract_text() or ""

            if numero_pagina <= 2:
                texto_inicial += " " + texto

            registros_pagina = extraer_registros_desde_lineas(
                texto.splitlines()
            )

            if not registros_pagina:
                paginas_sin_registros += 1

            for registro in registros_pagina:
                registro["PAGINA_PDF"] = numero_pagina
                registros.append(registro)

        if not registros:
            incidencias.append(
                {
                    "archivo": ruta.name,
                    "tipo": "PDF",
                    "estado": "SIN_REGISTROS",
                    "detalle": (
                        "El PDF abrió correctamente, pero no se "
                        "reconocieron registros con el patrón actual."
                    ),
                }
            )

            return pd.DataFrame(), incidencias

        df = pd.DataFrame(registros)

        df = df.drop_duplicates(
            subset=["CVU", "NOMBRE_COMPLETO_PDF", "NIVEL"]
        )

        df.insert(0, "ANIO", 2025)
        df.insert(1, "ORIGEN_ARCHIVO", ruta.name)
        df.insert(2, "ORIGEN_HOJA", "PDF")
        df.insert(
            3,
            "ORIGEN_2025",
            detectar_origen_2025(
                ruta.name,
                texto_inicial,
            ),
        )

        incidencias.append(
            {
                "archivo": ruta.name,
                "tipo": "PDF",
                "estado": "PROCESADO",
                "detalle": (
                    f"{len(df)} registros; "
                    f"{paginas_sin_registros} páginas sin registros."
                ),
            }
        )

        return df.reset_index(drop=True), incidencias

    except Exception as error:
        incidencias.append(
            {
                "archivo": ruta.name,
                "tipo": "PDF",
                "estado": "NO_PROCESADO",
                "detalle": repr(error),
            }
        )

        return pd.DataFrame(), incidencias


# ============================================================
# INTEGRACIÓN Y SALIDA
# ============================================================

def ordenar_columnas(master: pd.DataFrame) -> pd.DataFrame:
    prioridad = [
        "ANIO",
        "ORIGEN_ARCHIVO",
        "ORIGEN_HOJA",
        "ORIGEN_2025",
        "PAGINA_PDF",
        "CVU",
        "APELLIDO_PATERNO",
        "APELLIDO_MATERNO",
        "NOMBRES",
        "NOMBRE_COMPLETO_PDF",
        "NIVEL",
    ]

    primeras = [
        columna
        for columna in prioridad
        if columna in master.columns
    ]

    restantes = [
        columna
        for columna in master.columns
        if columna not in primeras
    ]

    return master[primeras + restantes]


def guardar_salidas(
    master: pd.DataFrame,
    control: pd.DataFrame,
    incidencias: pd.DataFrame,
) -> None:
    asegurar_directorio_salida()

    ruta_parquet = DATA_DIR / "SNII_MASTER.parquet"
    ruta_csv = DATA_DIR / "SNII_MASTER.csv"
    ruta_excel = DATA_DIR / "SNII_MASTER.xlsx"
    ruta_control = DATA_DIR / "CONTROL_ARCHIVOS.csv"
    ruta_incidencias = DATA_DIR / "INCIDENCIAS.csv"

    master.to_parquet(
        ruta_parquet,
        index=False,
        compression="snappy",
    )

    master.to_csv(
        ruta_csv,
        index=False,
        encoding="utf-8-sig",
    )

    with pd.ExcelWriter(
        ruta_excel,
        engine="openpyxl",
    ) as writer:
        master.to_excel(
            writer,
            sheet_name="MASTER",
            index=False,
        )

        control.to_excel(
            writer,
            sheet_name="CONTROL_ARCHIVOS",
            index=False,
        )

        incidencias.to_excel(
            writer,
            sheet_name="INCIDENCIAS",
            index=False,
        )

    control.to_csv(
        ruta_control,
        index=False,
        encoding="utf-8-sig",
    )

    incidencias.to_csv(
        ruta_incidencias,
        index=False,
        encoding="utf-8-sig",
    )

    print("\nArchivos generados:")
    print(f"- {ruta_parquet}")
    print(f"- {ruta_csv}")
    print(f"- {ruta_excel}")
    print(f"- {ruta_control}")
    print(f"- {ruta_incidencias}")


def construir_master() -> None:
    asegurar_directorio_salida()

    carpeta_temporal = Path(
        tempfile.mkdtemp(prefix="snii_builder_")
    )

    try:
        print("1. Descargando archivos de Google Drive...")

        descargados = descargar_carpeta_drive(
            carpeta_temporal
        )

        if not descargados:
            raise RuntimeError(
                "Google Drive no devolvió archivos. "
                "Verifica que la carpeta sea pública."
            )

        print("2. Inventariando archivos...")

        archivos = listar_archivos(carpeta_temporal)

        if archivos.empty:
            raise RuntimeError(
                "No se encontraron archivos compatibles."
            )

        tablas: list[pd.DataFrame] = []
        incidencias: list[dict[str, Any]] = []

        total = len(archivos)

        for posicion, fila in archivos.iterrows():
            ruta = Path(fila["ruta_completa"])
            extension = ruta.suffix.lower()

            print(
                f"3. Procesando {posicion + 1}/{total}: "
                f"{ruta.name}"
            )

            if extension in {".xlsx", ".xls"}:
                nuevas_tablas, errores = leer_excel(ruta)
                tablas.extend(nuevas_tablas)
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

        print("4. Uniendo tablas...")

        master = pd.concat(
            tablas,
            ignore_index=True,
            sort=False,
        )

        master = ordenar_columnas(master)

        control = archivos[COLUMNAS_CONTROL].copy()

        incidencias_df = pd.DataFrame(
            incidencias,
            columns=[
                "archivo",
                "tipo",
                "estado",
                "detalle",
            ],
        )

        print("5. Guardando archivos...")

        guardar_salidas(
            master=master,
            control=control,
            incidencias=incidencias_df,
        )

        print("\nResumen")
        print(f"- Registros: {len(master):,}")
        print(
            "- Archivos fuente: "
            f"{master['ORIGEN_ARCHIVO'].nunique():,}"
        )

        registros_pdf = (
            master["ORIGEN_HOJA"].eq("PDF").sum()
            if "ORIGEN_HOJA" in master.columns
            else 0
        )

        print(f"- Registros PDF: {registros_pdf:,}")
        print(f"- Incidencias: {len(incidencias_df):,}")

    finally:
        shutil.rmtree(
            carpeta_temporal,
            ignore_errors=True,
        )


if __name__ == "__main__":
    construir_master()
