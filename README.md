# SNII Insight

Arquitectura separada en dos etapas:

## 1. Constructor

`builder.py` descarga y procesa los archivos históricos.

```bash
pip install -r requirements-builder.txt
python builder.py
```

Genera:

- `data/SNII_MASTER.parquet`
- `data/SNII_MASTER.csv`
- `data/SNII_MASTER.xlsx`
- `data/CONTROL_ARCHIVOS.csv`
- `data/INCIDENCIAS.csv`

## 2. Dashboard

`app.py` únicamente lee:

```text
data/SNII_MASTER.parquet
```

Para ejecutarlo:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso en GitHub Actions

El repositorio incluye:

```text
.github/workflows/build_master.yml
```

En GitHub:

1. Abre la pestaña **Actions**.
2. Selecciona **Construir MASTER SNII**.
3. Pulsa **Run workflow**.
4. El proceso genera y guarda los archivos dentro de `data/`.
5. Streamlit detectará el nuevo `SNII_MASTER.parquet`.

## Ventaja principal

Streamlit ya no instala `pypdf`, no descarga Google Drive y no procesa los PDF al iniciar. Sólo carga el archivo Parquet previamente construido.
