from __future__ import annotations

import hmac
import os
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from boletin_core import (
    APP_VERSION,
    ai_filters,
    ai_summary,
    apply_filters,
    dataframe_for_display,
    load_database,
    refresh_database_from_folder,
    refresh_database_from_uploaded_files,
    save_database,
    to_excel_bytes,
)
from onedrive_graph import OneDriveConfig, load_onedrive_folder

load_dotenv()

APP_DIR = Path(__file__).resolve().parent
DEFAULT_FOLDER = APP_DIR / "ejemplo_boletines"
DEFAULT_DB = APP_DIR / "data" / "boletines.sqlite"

st.set_page_config(
    page_title="Chatbot Boletín Oficial CABA",
    page_icon="📄",
    layout="wide",
)


def secret(name: str, default: str = "") -> str:
    """Lee primero st.secrets y después variables de entorno/.env."""
    try:
        value: Any = st.secrets.get(name, None)
        if value is not None:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def secret_bool(name: str, default: bool = False) -> bool:
    value = secret(name, "")
    if value == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "si", "sí"}


def check_password() -> bool:
    app_password = secret("APP_PASSWORD", "")
    if not app_password:
        st.warning("Modo sin contraseña. Para publicarlo online, configurá APP_PASSWORD en Secrets.")
        return True

    if st.session_state.get("authenticated"):
        return True

    st.title(f"📄 Chatbot Boletín Oficial CABA {APP_VERSION}")
    st.caption("Ingreso privado")
    password = st.text_input("Contraseña", type="password")
    if st.button("Ingresar", type="primary"):
        if hmac.compare_digest(password, app_password):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False


if not check_password():
    st.stop()

st.title(f"📄 Chatbot Boletín Oficial CABA {APP_VERSION}")
st.caption("Consultá diariamente los Excel del Boletín Oficial cargados en OneDrive/SharePoint, en una carpeta local o subidos manualmente.")

openai_key = secret("OPENAI_API_KEY", "")
openai_model = secret("OPENAI_MODEL", "gpt-4o-mini")

with st.sidebar:
    st.header("Configuración")

    source_default = secret("DATA_SOURCE", "local").lower()
    source_options = {
        "local": "Carpeta local / OneDrive sincronizado",
        "upload": "Subir Excel manualmente",
        "onedrive": "OneDrive / SharePoint directo",
    }
    default_index = list(source_options.keys()).index(source_default) if source_default in source_options else 0
    source_label = st.selectbox("Fuente de datos", list(source_options.values()), index=default_index)
    source = {v: k for k, v in source_options.items()}[source_label]

    db_path = st.text_input(
        "Base SQLite",
        value=secret("BOLETIN_DB", str(DEFAULT_DB)),
        help="Archivo donde se consolida la información de todos los Excel. En nube puede ser /tmp/boletines.sqlite.",
    )

    use_ai = st.toggle(
        "Usar IA para interpretar preguntas",
        value=bool(openai_key),
        help="Requiere OPENAI_API_KEY. Si está apagado, funciona igual con filtros y búsqueda por palabras clave.",
    )

    st.divider()

    uploaded_files = []
    folder_path = ""
    onedrive_config = None

    if source == "local":
        folder_path = st.text_input(
            "Carpeta con Excel",
            value=secret("BOLETIN_FOLDER", str(DEFAULT_FOLDER)),
            help="Pegá la ruta de la carpeta sincronizada con OneDrive. Ejemplo: C:\\Users\\TuUsuario\\OneDrive - Empresa\\Boletines",
        )
        refresh_label = "🔄 Actualizar base desde la carpeta"

    elif source == "upload":
        uploaded_files = st.file_uploader(
            "Subí uno o más Excel del Boletín",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
        )
        refresh_label = "🔄 Actualizar base con archivos subidos"

    else:
        onedrive_config = OneDriveConfig(
            tenant_id=secret("MS_TENANT_ID", ""),
            client_id=secret("MS_CLIENT_ID", ""),
            client_secret=secret("MS_CLIENT_SECRET", ""),
            drive_id=secret("MS_DRIVE_ID", "") or None,
            site_id=secret("MS_SITE_ID", "") or None,
            site_hostname=secret("MS_SITE_HOSTNAME", "") or None,
            site_path=secret("MS_SITE_PATH", "") or None,
            user_id=secret("MS_USER_ID", "") or None,
            folder_path=secret("MS_FOLDER_PATH", "/"),
            recursive=secret_bool("MS_RECURSIVE", False),
        )
        if onedrive_config.is_configured:
            st.success("Microsoft Graph configurado.")
        else:
            st.info("Faltan Secrets de Microsoft Graph. Revisá README_DEPLOY.md.")
        refresh_label = "🔄 Actualizar base desde OneDrive/SharePoint"

    refresh = st.button(refresh_label, use_container_width=True)

    if st.session_state.get("authenticated") and secret("APP_PASSWORD", ""):
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()

if refresh:
    with st.spinner("Leyendo Excel y actualizando base..."):
        if source == "local":
            df_refreshed, errors = refresh_database_from_folder(folder_path, db_path)
        elif source == "upload":
            df_refreshed, errors = refresh_database_from_uploaded_files(uploaded_files, db_path)
        else:
            df_refreshed, errors = load_onedrive_folder(onedrive_config)  # type: ignore[arg-type]
            if not df_refreshed.empty:
                save_database(df_refreshed, db_path)

    st.success(f"Base actualizada: {len(df_refreshed):,} normas cargadas.".replace(",", "."))
    if errors:
        with st.expander("Archivos o procesos con errores"):
            for err in errors:
                st.warning(err)

# Primera carga automática solo para la versión local de ejemplo.
df = load_database(db_path)
if df.empty and source == "local" and Path(str(folder_path).strip().strip('"')).expanduser().exists():
    with st.spinner("Primera carga: leyendo Excel de la carpeta..."):
        df, errors = refresh_database_from_folder(folder_path, db_path)
    if errors:
        with st.expander("Archivos con errores"):
            for err in errors:
                st.warning(err)

if df.empty:
    st.info("Todavía no hay datos cargados. Elegí la fuente de datos y tocá el botón de actualización.")
    st.stop()

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Normas cargadas", f"{len(df):,}".replace(",", "."))
with c2:
    boletines = df["nro_boletin"].nunique() if "nro_boletin" in df.columns else 0
    st.metric("Boletines", boletines)
with c3:
    fechas = pd.to_datetime(df["fecha"], errors="coerce") if "fecha" in df.columns else pd.Series(dtype="datetime64[ns]")
    st.metric("Última fecha", fechas.max().strftime("%d/%m/%Y") if fechas.notna().any() else "-")
with c4:
    anexos = df["tiene_anexos"].fillna("").astype(str).str.lower().str.startswith(("sí", "si")).sum() if "tiene_anexos" in df.columns else 0
    st.metric("Con anexos", int(anexos))

st.divider()

st.subheader("Consulta en lenguaje natural")

examples = [
    "Licitaciones de los últimos 15 días",
    "Mostrame resoluciones de APRA del 28/04/2026",
    "Buscá normas vinculadas a espacio público",
    "¿Qué documentos tienen anexos?",
    "Listame normas del Ministerio de Hacienda y Finanzas",
    "Normas de salud del último mes",
]
question = st.text_input("Escribí tu pregunta", placeholder=examples[0])

with st.expander("Ejemplos de preguntas"):
    for ex in examples:
        st.write(f"• {ex}")

with st.expander("Filtros manuales opcionales"):
    f1, f2, f3, f4 = st.columns(4)

    fechas = pd.to_datetime(df["fecha"], errors="coerce") if "fecha" in df.columns else pd.Series(dtype="datetime64[ns]")
    min_date = fechas.min().date() if fechas.notna().any() else None
    max_date = fechas.max().date() if fechas.notna().any() else None

    usar_fechas = st.checkbox("Filtrar por rango de fechas", value=False)
    with f1:
        fecha_desde = st.date_input("Desde", value=min_date, min_value=min_date, max_value=max_date, disabled=not usar_fechas)
    with f2:
        fecha_hasta = st.date_input("Hasta", value=max_date, min_value=min_date, max_value=max_date, disabled=not usar_fechas)
    with f3:
        tipos = sorted([x for x in df.get("tipo_norma", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        tipo = st.selectbox("Tipo de norma", [""] + tipos)
    with f4:
        anexos_opt = st.selectbox("Anexos", ["", "Con anexos", "Sin anexos"])

    f5, f6, f7 = st.columns([2, 2, 1])
    with f5:
        areas = sorted([x for x in df.get("area", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        area = st.selectbox("Área", [""] + areas)
    with f6:
        poderes = sorted([x for x in df.get("poder_organismo", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        poder = st.selectbox("Poder / Organismo", [""] + poderes)
    with f7:
        limitar_resultados = st.checkbox("Limitar resultados", value=False)
        limit = None
        if limitar_resultados:
            limit = st.number_input("Cantidad máxima", min_value=1, max_value=10000, value=500, step=100)

run_query = st.button("🔎 Consultar", type="primary")

if run_query or question:
    api_key = openai_key if use_ai else ""
    filters = ai_filters(question or "", df, api_key=api_key, model=openai_model)

    if usar_fechas and fecha_desde:
        filters["fecha_desde"] = fecha_desde.strftime("%Y-%m-%d")
    if usar_fechas and fecha_hasta:
        filters["fecha_hasta"] = fecha_hasta.strftime("%Y-%m-%d")
    if tipo:
        filters["tipo_norma"] = tipo
    if area:
        filters["area"] = area
    if poder:
        filters["poder_organismo"] = poder
    if anexos_opt == "Con anexos":
        filters["tiene_anexos"] = True
    elif anexos_opt == "Sin anexos":
        filters["tiene_anexos"] = False
    filters["limit"] = int(limit) if limit else None

    results = apply_filters(df, filters)

    st.markdown("### Respuesta")
    st.write(ai_summary(question or "Consulta filtrada", results, api_key=api_key, model=openai_model))

    with st.expander("Filtros interpretados"):
        st.caption(f"Versión del parser: {APP_VERSION}. Para fechas relativas, se usa como referencia la última fecha cargada en la base.")
        st.json(filters)

    st.markdown("### Resultados")
    if filters.get("limit"):
        st.caption(f"Mostrando hasta {filters['limit']} resultados. Para ver todos, desactivá 'Limitar resultados'.")
    else:
        st.caption(f"Mostrando todos los resultados encontrados: {len(results)}.")
    display_df = dataframe_for_display(results)
    column_config = {}
    if "URL Documento" in display_df.columns:
        column_config["URL Documento"] = st.column_config.LinkColumn("URL Documento", display_text="Abrir documento")

    st.dataframe(display_df, use_container_width=True, hide_index=True, column_config=column_config)

    cdl1, cdl2 = st.columns(2)
    with cdl1:
        csv = display_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Descargar CSV", data=csv, file_name="resultados_boletin_caba.csv", mime="text/csv")
    with cdl2:
        st.download_button(
            "⬇️ Descargar Excel",
            data=to_excel_bytes(display_df),
            file_name="resultados_boletin_caba.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.markdown("### Últimas normas cargadas")
    latest = df.sort_values("fecha", ascending=False).head(50) if "fecha" in df.columns else df.head(50)
    st.dataframe(dataframe_for_display(latest), use_container_width=True, hide_index=True)
