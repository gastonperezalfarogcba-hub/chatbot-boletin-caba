from __future__ import annotations

import json
import re
import sqlite3
import unicodedata
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd


EXPECTED_COLUMNS = {
    "fecha": "fecha",
    "n boletin": "nro_boletin",
    "nro boletin": "nro_boletin",
    "numero boletin": "nro_boletin",
    "poder organismo": "poder_organismo",
    "poder / organismo": "poder_organismo",
    "tipo de norma": "tipo_norma",
    "area": "area",
    "nombre": "nombre",
    "sumario": "sumario",
    "url documento": "url_documento",
    "tiene anexos": "tiene_anexos",
}

APP_VERSION = "v2.3"

DISPLAY_COLUMNS = [
    "fecha",
    "nro_boletin",
    "poder_organismo",
    "tipo_norma",
    "area",
    "nombre",
    "sumario",
    "url_documento",
    "tiene_anexos",
    "archivo_origen",
]

STOPWORDS = {
    "que", "qué", "cual", "cuál", "cuales", "cuáles", "como", "cómo", "donde", "dónde",
    "cuando", "cuándo", "quien", "quién", "quienes", "quiénes", "por", "para", "con",
    "sin", "los", "las", "una", "uno", "unos", "unas", "del", "de", "la", "el", "en",
    "y", "o", "u", "a", "al", "se", "me", "te", "lo", "le", "les", "sus", "su",
    "mostrame", "mostrar", "listame", "listar", "busca", "buscar", "buscame", "dame",
    "traeme", "traer", "consulta", "consultar", "norma", "normas", "salio", "salió",
    "salieron", "boletin", "boletín", "oficial", "caba", "dia", "día", "dias", "días",
    "semana", "semanas", "mes", "meses", "ano", "año", "ver", "hay", "hubo", "quiero",
    "necesito", "informacion", "información", "sobre", "ultimos", "últimos", "ultimas",
    "últimas", "ultimo", "último", "ultima", "última", "desde", "hasta",
    "resolucion", "resoluciones", "decreto", "decretos", "disposicion", "disposiciones",
    "ley", "leyes", "licitacion", "licitaciones", "publica", "publicas", "pública", "públicas",
    "comunicado", "comunicados", "edicion", "edición",
}


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_colname(col: Any) -> str:
    text = normalize_text(col)
    text = text.replace("°", "").replace("º", "")
    text = re.sub(r"[^a-z0-9/ ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return EXPECTED_COLUMNS.get(text, text.replace(" ", "_").replace("/", ""))


def _read_normas_excel(excel_source: Any, source_name: str) -> pd.DataFrame:
    """Lee la hoja Normas de un Excel del Boletín Oficial CABA.

    Estructura esperada:
    - Hoja: Normas
    - Encabezados: fila 3 del Excel, es decir header=2 en pandas.
    """
    try:
        df = pd.read_excel(
            excel_source,
            sheet_name="Normas",
            header=2,
            dtype=str,
            engine="openpyxl",
        )
    except ValueError as exc:
        raise ValueError(f"El archivo {source_name} no tiene una hoja llamada 'Normas'.") from exc
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"No pude leer {source_name}: {exc}") from exc

    df = df.dropna(axis=1, how="all")
    df.columns = [normalize_colname(c) for c in df.columns]

    for col in DISPLAY_COLUMNS:
        if col not in df.columns and col != "archivo_origen":
            df[col] = ""

    if "sumario" in df.columns:
        df = df[df["sumario"].notna()]
    df = df.dropna(how="all").copy()

    df["archivo_origen"] = source_name

    if "fecha" in df.columns:
        fecha = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")
        df["fecha"] = fecha.dt.strftime("%Y-%m-%d").fillna(df["fecha"].astype(str))

    for col in DISPLAY_COLUMNS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    key_cols = [c for c in ["fecha", "nro_boletin", "tipo_norma", "area", "nombre", "url_documento", "sumario"] if c in df.columns]
    if key_cols:
        df = df.drop_duplicates(subset=key_cols, keep="first")

    return df[[c for c in DISPLAY_COLUMNS if c in df.columns]]


def read_normas_from_excel(file_path: str | Path) -> pd.DataFrame:
    file_path = Path(file_path)
    return _read_normas_excel(file_path, file_path.name)


def read_normas_from_bytes(file_name: str, content: bytes) -> pd.DataFrame:
    return _read_normas_excel(BytesIO(content), file_name)


def list_excel_files(folder_path: str | Path, recursive: bool = True) -> List[Path]:
    folder = Path(str(folder_path).strip().strip('"')).expanduser()
    if not folder.exists():
        return []
    pattern = "**/*.xls*" if recursive else "*.xls*"
    files = []
    for p in folder.glob(pattern):
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() in {".xlsx", ".xlsm"}:
            files.append(p)
    return sorted(files)


def consolidate_frames(frames: List[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame(columns=DISPLAY_COLUMNS)
    df = pd.concat(frames, ignore_index=True)
    key_cols = [c for c in ["fecha", "nro_boletin", "tipo_norma", "area", "nombre", "url_documento", "sumario"] if c in df.columns]
    if key_cols:
        df = df.drop_duplicates(subset=key_cols, keep="first")
    return df[[c for c in DISPLAY_COLUMNS if c in df.columns]]


def load_folder(folder_path: str | Path) -> Tuple[pd.DataFrame, List[str]]:
    rows = []
    errors = []
    for file_path in list_excel_files(folder_path):
        try:
            rows.append(read_normas_from_excel(file_path))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{file_path.name}: {exc}")
    return consolidate_frames(rows), errors


def load_uploaded_files(uploaded_files: List[Any]) -> Tuple[pd.DataFrame, List[str]]:
    rows = []
    errors = []
    for uploaded in uploaded_files or []:
        try:
            name = getattr(uploaded, "name", "archivo.xlsx")
            rows.append(read_normas_from_bytes(name, uploaded.getvalue()))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{getattr(uploaded, 'name', 'archivo')}: {exc}")
    return consolidate_frames(rows), errors


def save_database(df: pd.DataFrame, db_path: str | Path) -> None:
    db_path = Path(str(db_path).strip().strip('"'))
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        df.to_sql("normas", conn, if_exists="replace", index=False)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_normas_fecha ON normas(fecha)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_normas_tipo ON normas(tipo_norma)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_normas_area ON normas(area)")
        conn.commit()


def refresh_database_from_folder(folder_path: str | Path, db_path: str | Path) -> Tuple[pd.DataFrame, List[str]]:
    df, errors = load_folder(folder_path)
    save_database(df, db_path)
    return df, errors


def refresh_database_from_uploaded_files(uploaded_files: List[Any], db_path: str | Path) -> Tuple[pd.DataFrame, List[str]]:
    df, errors = load_uploaded_files(uploaded_files)
    save_database(df, db_path)
    return df, errors


def load_database(db_path: str | Path) -> pd.DataFrame:
    db_path = Path(str(db_path).strip().strip('"'))
    if not db_path.exists():
        return pd.DataFrame(columns=DISPLAY_COLUMNS)
    with sqlite3.connect(db_path) as conn:
        try:
            return pd.read_sql_query("SELECT * FROM normas", conn)
        except Exception:
            return pd.DataFrame(columns=DISPLAY_COLUMNS)


def extract_keywords(question: str) -> List[str]:
    text = normalize_text(question)
    words = re.findall(r"[a-z0-9]{3,}", text)
    keywords: List[str] = []
    for w in words:
        if re.fullmatch(r"20\d{2}|\d{1,3}", w):
            continue
        if w in STOPWORDS:
            continue
        if w not in keywords:
            keywords.append(w)
    return keywords[:15]


def reference_date(df: pd.DataFrame) -> datetime:
    if not df.empty and "fecha" in df.columns:
        fechas = pd.to_datetime(df["fecha"], errors="coerce")
        if fechas.notna().any():
            return fechas.max().to_pydatetime()
    return datetime.now()


def parse_date_filters(question: str, df: pd.DataFrame) -> Dict[str, str | None]:
    """Interpreta fechas escritas en lenguaje natural.

    Importante: usa como referencia la fecha máxima cargada en la base, no la fecha real del día.
    Ejemplo: si la última fecha cargada es 2026-04-28, "últimos 15 días" = 2026-04-13 a 2026-04-28.
    """
    text = normalize_text(question)
    out: Dict[str, str | None] = {"fecha_desde": None, "fecha_hasta": None}
    ref = reference_date(df)

    def set_range(days: int) -> Dict[str, str | None]:
        return {
            "fecha_desde": (ref - timedelta(days=days)).strftime("%Y-%m-%d"),
            "fecha_hasta": ref.strftime("%Y-%m-%d"),
        }

    # Fecha exacta: 28/04/2026 o 28-04-2026.
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b", text)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            date = datetime(y, mo, d).strftime("%Y-%m-%d")
            return {"fecha_desde": date, "fecha_hasta": date}
        except ValueError:
            pass

    # Mes mencionado: abril 2026, abril.
    months = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
        "noviembre": 11, "diciembre": 12,
    }
    for name, month in months.items():
        if re.search(rf"\b{name}\b", text):
            year_match = re.search(r"\b(20\d{2})\b", text)
            year = int(year_match.group(1)) if year_match else ref.year
            start = datetime(year, month, 1)
            end = datetime(year, 12, 31) if month == 12 else datetime(year, month + 1, 1) - timedelta(days=1)
            return {"fecha_desde": start.strftime("%Y-%m-%d"), "fecha_hasta": end.strftime("%Y-%m-%d")}

    # Expresiones relativas.
    if re.search(r"\bhoy\b", text):
        date = ref.strftime("%Y-%m-%d")
        return {"fecha_desde": date, "fecha_hasta": date}
    if re.search(r"\bayer\b", text):
        date = (ref - timedelta(days=1)).strftime("%Y-%m-%d")
        return {"fecha_desde": date, "fecha_hasta": date}

    # Casos robustos: ultimos 15 dias / últimos quince días / ultimas 2 semanas / ultimo mes.
    # normalize_text ya convierte "últimos días" en "ultimos dias".
    number_words = {
        "uno": 1, "una": 1, "dos": 2, "tres": 3, "cuatro": 4, "cinco": 5,
        "seis": 6, "siete": 7, "ocho": 8, "nueve": 9, "diez": 10,
        "once": 11, "doce": 12, "trece": 13, "catorce": 14, "quince": 15,
        "veinte": 20, "treinta": 30,
    }
    num_pattern = r"(\d{1,3}|uno|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce|trece|catorce|quince|veinte|treinta)"

    m = re.search(rf"\bultim(?:o|a|os|as)?\s+{num_pattern}\s+d(?:ia|ias)\b", text)
    if m:
        raw = m.group(1)
        days = int(raw) if raw.isdigit() else number_words.get(raw, 15)
        return set_range(days)

    m = re.search(rf"\bultim(?:o|a|os|as)?\s+{num_pattern}\s+semanas?\b", text)
    if m:
        raw = m.group(1)
        weeks = int(raw) if raw.isdigit() else number_words.get(raw, 1)
        return set_range(weeks * 7)

    m = re.search(rf"\bultim(?:o|a|os|as)?\s+{num_pattern}\s+mes(?:es)?\b", text)
    if m:
        raw = m.group(1)
        months_n = int(raw) if raw.isdigit() else number_words.get(raw, 1)
        return set_range(months_n * 30)

    if re.search(r"\b(ultima|ultimo)\s+semana\b", text):
        return set_range(7)
    if re.search(r"\b(ultima|ultimo)\s+quincena\b", text) or re.search(r"\bultim(?:o|a|os|as)?\s+15\b", text):
        return set_range(15)
    if re.search(r"\b(ultimo|ultima)\s+mes\b", text):
        return set_range(30)

    return out

def heuristic_filters(question: str, df: pd.DataFrame) -> Dict[str, Any]:
    text = normalize_text(question)
    filters: Dict[str, Any] = {
        "fecha_desde": None,
        "fecha_hasta": None,
        "tipo_norma": None,
        "area": None,
        "poder_organismo": None,
        "keywords": extract_keywords(question),
        "tiene_anexos": None,
        "limit": None,
    }
    filters.update(parse_date_filters(question, df))

    tipos_posibles = [
        "ley", "decreto", "resolucion", "resolución", "disposicion", "disposición",
        "licitacion", "licitación", "comunicado", "edicto", "convenio", "acta",
        "resolucion de firma conjunta", "separata",
    ]
    for tipo in tipos_posibles:
        if normalize_text(tipo) in text:
            filters["tipo_norma"] = tipo
            break

    if "con anex" in text:
        filters["tiene_anexos"] = True
    elif "sin anex" in text:
        filters["tiene_anexos"] = False

    if not df.empty and "area" in df.columns:
        areas = sorted({str(x) for x in df["area"].dropna().unique() if str(x).strip()}, key=len, reverse=True)
        for area in areas[:5000]:
            if len(area) >= 3 and normalize_text(area) in text:
                filters["area"] = area
                break

    if not df.empty and "poder_organismo" in df.columns:
        orgs = sorted({str(x) for x in df["poder_organismo"].dropna().unique() if str(x).strip()}, key=len, reverse=True)
        for org in orgs[:1000]:
            if len(org) >= 3 and normalize_text(org) in text:
                filters["poder_organismo"] = org
                break

    return filters


def ai_filters(question: str, df: pd.DataFrame, api_key: str | None, model: str = "gpt-4o-mini") -> Dict[str, Any]:
    if not api_key:
        return heuristic_filters(question, df)

    try:
        from openai import OpenAI
    except Exception:
        return heuristic_filters(question, df)

    sample_areas: List[str] = []
    sample_tipos: List[str] = []
    sample_orgs: List[str] = []
    if not df.empty:
        if "area" in df.columns:
            sample_areas = sorted(df["area"].dropna().astype(str).unique().tolist())[:250]
        if "tipo_norma" in df.columns:
            sample_tipos = sorted(df["tipo_norma"].dropna().astype(str).unique().tolist())[:80]
        if "poder_organismo" in df.columns:
            sample_orgs = sorted(df["poder_organismo"].dropna().astype(str).unique().tolist())[:80]

    system = (
        "Sos un asistente que transforma preguntas sobre el Boletín Oficial CABA en filtros JSON. "
        "No inventes datos. No escribas SQL. Devolvé SOLO JSON válido con estas claves: "
        "fecha_desde, fecha_hasta, tipo_norma, area, poder_organismo, keywords, tiene_anexos, limit. "
        "Las fechas deben estar en formato YYYY-MM-DD o null. keywords debe ser lista de palabras o frases relevantes. "
        "tiene_anexos debe ser true, false o null. limit debe ser null por defecto; "
        "solo usá un entero si el usuario pide explícitamente primeros N, top N, solo N o límite N. "
        "Si el usuario pide 'últimos N días', usá como fecha_hasta la fecha máxima de la base: "
        f"{reference_date(df).strftime('%Y-%m-%d')}."
    )
    user = {
        "pregunta": question,
        "tipos_disponibles_ejemplo": sample_tipos,
        "areas_disponibles_ejemplo": sample_areas,
        "poderes_u_organismos_ejemplo": sample_orgs,
    }

    try:
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
            ],
            response_format={"type": "json_object"},
        )
        raw = resp.choices[0].message.content or "{}"
        data = json.loads(raw)
    except Exception:
        return heuristic_filters(question, df)

    fallback = heuristic_filters(question, df)
    for key in fallback:
        if key not in data:
            data[key] = fallback[key]

    # Si la IA no interpreta fechas relativas como "últimos 15 días",
    # conservamos las fechas calculadas por el parser local.
    for key in ["fecha_desde", "fecha_hasta"]:
        if not data.get(key) and fallback.get(key):
            data[key] = fallback[key]

    if not isinstance(data.get("keywords"), list):
        data["keywords"] = fallback["keywords"]
    else:
        # Quitamos palabras de tiempo que no aportan al filtrado textual.
        data["keywords"] = [kw for kw in data["keywords"] if normalize_text(kw) not in STOPWORDS]
        if not data["keywords"] and fallback.get("keywords"):
            data["keywords"] = fallback["keywords"]

    raw_limit = data.get("limit")
    if raw_limit in (None, "", "null", "todos", "all"):
        data["limit"] = None
    else:
        try:
            data["limit"] = max(1, min(int(raw_limit), 10000))
        except Exception:
            data["limit"] = None
    return data


def contains_norm(series: pd.Series, value: str) -> pd.Series:
    needle = normalize_text(value)
    if not needle:
        return pd.Series([True] * len(series), index=series.index)
    return series.fillna("").astype(str).map(lambda x: needle in normalize_text(x))


def apply_filters(df: pd.DataFrame, filters: Dict[str, Any]) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    out = df.copy()

    if filters.get("fecha_desde") and "fecha" in out.columns:
        out = out[pd.to_datetime(out["fecha"], errors="coerce") >= pd.to_datetime(filters["fecha_desde"], errors="coerce")]
    if filters.get("fecha_hasta") and "fecha" in out.columns:
        out = out[pd.to_datetime(out["fecha"], errors="coerce") <= pd.to_datetime(filters["fecha_hasta"], errors="coerce")]

    for field in ["tipo_norma", "area", "poder_organismo"]:
        value = filters.get(field)
        if value and field in out.columns:
            out = out[contains_norm(out[field], str(value))]

    if filters.get("tiene_anexos") is not None and "tiene_anexos" in out.columns:
        want = bool(filters["tiene_anexos"])
        out = out[out["tiene_anexos"].fillna("").astype(str).map(lambda x: normalize_text(x).startswith("si")) == want]

    keywords = [str(k).strip() for k in filters.get("keywords", []) if str(k).strip()]
    if keywords:
        searchable_cols = [c for c in ["area", "nombre", "sumario", "tipo_norma", "poder_organismo", "nro_boletin"] if c in out.columns]
        if searchable_cols:
            joined = out[searchable_cols].fillna("").astype(str).agg(" | ".join, axis=1).map(normalize_text)
            scores = pd.Series(0, index=out.index, dtype=int)
            for kw in keywords:
                nkw = normalize_text(kw)
                if len(nkw) < 3:
                    continue
                scores += joined.str.contains(re.escape(nkw), na=False).astype(int)
            out = out[scores > 0].assign(_score=scores[scores > 0]).sort_values(["_score", "fecha"], ascending=[False, False])
    else:
        if "fecha" in out.columns:
            out = out.sort_values("fecha", ascending=False)

    limit = filters.get("limit")

    if "_score" in out.columns:
        out = out.drop(columns=["_score"])

    # Por defecto no se corta la consulta: se devuelven todos los resultados.
    # Esto evita errores de lectura cuando hay muchas licitaciones u otras normas
    # dentro de un rango de fechas. Solo se limita si el usuario lo pide o si
    # se activa el límite manual desde la interfaz.
    if limit in (None, "", "null", "todos", "all", 0, "0"):
        return out

    try:
        limit_int = max(1, min(int(limit), 10000))
    except Exception:
        return out

    return out.head(limit_int)


def dataframe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "fecha" in out.columns:
        parsed = pd.to_datetime(out["fecha"], errors="coerce")
        out["fecha"] = parsed.dt.strftime("%d/%m/%Y").fillna(out["fecha"])
    rename = {
        "fecha": "Fecha",
        "nro_boletin": "N° Boletín",
        "poder_organismo": "Poder / Organismo",
        "tipo_norma": "Tipo de Norma",
        "area": "Área",
        "nombre": "Nombre",
        "sumario": "Sumario",
        "url_documento": "URL Documento",
        "tiene_anexos": "Tiene Anexos",
        "archivo_origen": "Archivo origen",
    }
    return out.rename(columns=rename)


def make_deterministic_summary(results: pd.DataFrame, question: str = "") -> str:
    if results.empty:
        return "No encontré resultados para esa consulta con los datos cargados."

    total = len(results)
    tipos = results["tipo_norma"].value_counts().head(5).to_dict() if "tipo_norma" in results else {}
    areas = results["area"].value_counts().head(5).to_dict() if "area" in results else {}

    parts = [f"Encontré {total} resultado{'s' if total != 1 else ''}."]
    if tipos:
        parts.append("Tipos principales: " + ", ".join([f"{k} ({v})" for k, v in tipos.items() if k]) + ".")
    if areas:
        parts.append("Áreas más frecuentes: " + ", ".join([f"{k} ({v})" for k, v in areas.items() if k]) + ".")

    first = results.head(5)
    if not first.empty:
        ejemplos = []
        for _, row in first.iterrows():
            nombre = row.get("nombre", "")
            area = row.get("area", "")
            tipo = row.get("tipo_norma", "")
            fecha = row.get("fecha", "")
            ejemplos.append(f"{fecha} — {tipo} {nombre} — {area}".strip(" —"))
        parts.append("Primeros resultados: " + " | ".join(ejemplos) + ".")
    return " ".join(parts)


def ai_summary(question: str, results: pd.DataFrame, api_key: str | None, model: str = "gpt-4o-mini") -> str:
    if not api_key or results.empty:
        return make_deterministic_summary(results, question)

    try:
        from openai import OpenAI
    except Exception:
        return make_deterministic_summary(results, question)

    cols = [c for c in ["fecha", "nro_boletin", "tipo_norma", "area", "nombre", "sumario", "url_documento", "tiene_anexos"] if c in results.columns]
    sample = results[cols].head(30).to_dict(orient="records")

    prompt = (
        "Respondé en español rioplatense, de forma clara y ejecutiva. "
        "Resumí los resultados encontrados en el Boletín Oficial CABA. "
        "No inventes normas ni conclusiones que no estén en la tabla. "
        "Mencioná cantidad de resultados, patrones relevantes y 3 a 5 hallazgos concretos. "
        "Si hay URLs, no las inventes ni las modifiques."
    )
    try:
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            temperature=0.2,
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": json.dumps({"pregunta": question, "resultados": sample}, ensure_ascii=False)},
            ],
        )
        return resp.choices[0].message.content or make_deterministic_summary(results, question)
    except Exception:
        return make_deterministic_summary(results, question)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")
    bio.seek(0)
    return bio.read()
