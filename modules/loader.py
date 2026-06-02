# ═══════════════════════════════════════════════════════════
# BISOPI Automator v1.0 — Mayo 2026
# Autor: Edson Leon
# Cargo: Data Insights Lead — Operaciones
# Empresa: Bision Consulting
# Contacto: edson.leon@bisionconsulting.com
# ═══════════════════════════════════════════════════════════

"""
Carga registros desde un archivo .xlsx o desde texto copy-paste.
Devuelve un DataFrame con las columnas internas estándar.
"""
import io
import pandas as pd

# Columnas internas del modelo (snake-style sin espacios)
REGISTRO_COLUMNS = [
    "Proyecto",
    "GrupoTarea",
    "Tarea",
    "TipoHora",
    "FechaRegistro",
    "Horas",
    "Minutos",
    "HoraInicio",
    "Comentario",
    "Estado",
    "RespuestaAPI",
]

# Mapeo nombre Excel → nombre interno
_COL_MAP = {
    "Proyecto":       "Proyecto",
    "Grupo Tarea":    "GrupoTarea",
    "Tarea":          "Tarea",
    "Tipo Hora":      "TipoHora",
    "Fecha Registro": "FechaRegistro",
    "Horas":          "Horas",
    "Minutos":        "Minutos",
    "Hora Inicio":    "HoraInicio",
    "Comentario":     "Comentario",
    "Estado":         "Estado",
    "Respuesta API":  "RespuestaAPI",
}

_REQUIRED = {"Proyecto", "GrupoTarea", "Tarea", "TipoHora", "FechaRegistro", "Horas", "Minutos"}
_OPTIONAL = {"HoraInicio", "Comentario", "Estado", "RespuestaAPI"}


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Limpieza y normalización de tipos. Asume que las columnas ya están renombradas."""

    # Añadir columnas opcionales ausentes
    for col in _OPTIONAL:
        if col not in df.columns:
            df[col] = ""

    # Descartar filas sin los campos mínimos de un registro real (incluye la fila de NOTAS al pie)
    df = df.dropna(subset=["Tarea", "FechaRegistro"], how="any")
    df = df[df["Proyecto"].astype(str).str.strip() != ""]
    df = df[df["Tarea"].astype(str).str.strip() != ""]

    # Ignorar silenciosamente las filas ya cargadas en sesiones previas
    if "Estado" in df.columns:
        df = df[df["Estado"].astype(str).str.strip() != "✅ Cargado"]

    if df.empty:
        return pd.DataFrame(columns=REGISTRO_COLUMNS)

    # FechaRegistro → string YYYY-MM-DD
    df["FechaRegistro"] = (
        pd.to_datetime(df["FechaRegistro"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )

    # Horas y Minutos → int
    df["Horas"]   = pd.to_numeric(df["Horas"],   errors="coerce").fillna(0).astype(int)
    df["Minutos"] = pd.to_numeric(df["Minutos"], errors="coerce").fillna(0).astype(int)

    # Resto → string limpio
    str_cols = [c for c in REGISTRO_COLUMNS if c not in ("Horas", "Minutos", "FechaRegistro")]
    for col in str_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Estado vacío → "Pendiente"
    df["Estado"] = df["Estado"].where(df["Estado"] != "", "Pendiente")

    return df[REGISTRO_COLUMNS].reset_index(drop=True)


def load_from_file(file) -> pd.DataFrame:
    """
    Recibe el objeto de archivo de st.file_uploader.
    Lee la hoja 'Registro' (fila 1 = título, fila 2 = encabezados).
    """
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"No se pudo abrir el archivo Excel: {exc}") from exc

    if "Registro" not in xls.sheet_names:
        hojas = ", ".join(xls.sheet_names)
        raise ValueError(
            f"El archivo no contiene la hoja 'Registro'. "
            f"Hojas encontradas: {hojas}"
        )

    # header=1 → pandas usa la fila de índice 1 (segunda fila) como encabezado,
    # saltando el título de la fila 0.
    df = pd.read_excel(xls, sheet_name="Registro", header=1, engine="openpyxl")
    df = df.rename(columns=_COL_MAP)

    missing = _REQUIRED - set(df.columns)
    if missing:
        raise ValueError(
            f"La hoja 'Registro' no tiene las columnas requeridas: "
            f"{', '.join(sorted(missing))}"
        )

    return _normalize(df)


# ── Días de la semana en español ──────────────────────────────────────────────
_DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


def _fmt_time(minutes: int) -> str:
    h, m = divmod(abs(int(minutes)), 60)
    return f"{h}h {m:02d}m"


def build_summary(df: pd.DataFrame) -> dict:
    """
    Construye métricas globales y tabla de resumen por día.
    Solo incluye filas con Estado == "Pendiente" (las que se enviarán a BISOPI).
    Las filas con error se excluyen porque no se subirán.

    Retorna un dict con:
      total_registros       int
      horas_laborales_fmt   str  "Xh YYm"
      horas_adicionales_fmt str
      dias_distintos        int
      resumen_df            pd.DataFrame  (filas por día + fila TOTAL al pie)
      laboral_mins_por_dia  list[int]     paralela a resumen_df para highlighting
    """
    df_summary = df[df["Estado"] != "❌ Error"].copy() if "Estado" in df.columns else df.copy()
    df_summary["_min"] = df_summary["Horas"] * 60 + df_summary["Minutos"]
    df_summary["_dt"]  = pd.to_datetime(df_summary["FechaRegistro"], errors="coerce")

    es_laboral   = df_summary["TipoHora"].str.strip().str.lower() == "laboral"
    es_adicional = df_summary["TipoHora"].str.strip().str.lower() == "adicional"

    total_lab_min = int(df_summary.loc[es_laboral,   "_min"].sum())
    total_adi_min = int(df_summary.loc[es_adicional, "_min"].sum())
    dias_distintos = int(df_summary["_dt"].dropna().dt.date.nunique())

    # ── Tabla por día ─────────────────────────────────────────────────────────
    filas: list[dict] = []
    laboral_mins: list[int] = []

    for dt, grupo in df_summary.groupby("_dt", sort=True):
        lab = int(grupo.loc[grupo["TipoHora"].str.strip().str.lower() == "laboral",   "_min"].sum())
        adi = int(grupo.loc[grupo["TipoHora"].str.strip().str.lower() == "adicional", "_min"].sum())
        dia = _DIAS_ES[dt.weekday()]
        filas.append({
            "Fecha":             f"{dt.strftime('%d/%m/%Y')} ({dia})",
            "Registros":         len(grupo),
            "Horas Laborales":   _fmt_time(lab),
            "Horas Adicionales": _fmt_time(adi),
            "Total del día":     _fmt_time(lab + adi),
        })
        laboral_mins.append(lab)

    # Fila de totales
    filas.append({
        "Fecha":             "TOTAL",
        "Registros":         len(df_summary),
        "Horas Laborales":   _fmt_time(total_lab_min),
        "Horas Adicionales": _fmt_time(total_adi_min),
        "Total del día":     _fmt_time(total_lab_min + total_adi_min),
    })
    laboral_mins.append(-1)  # -1 = fila de totales, nunca resaltar

    return {
        "total_registros":       len(df_summary),
        "horas_laborales_fmt":   _fmt_time(total_lab_min),
        "horas_adicionales_fmt": _fmt_time(total_adi_min),
        "dias_distintos":        dias_distintos,
        "resumen_df":            pd.DataFrame(filas),
        "laboral_mins_por_dia":  laboral_mins,
    }


def load_from_text(text: str) -> pd.DataFrame:
    """
    Recibe texto TSV copiado de Excel (primera fila = encabezados de la plantilla).
    """
    text = text.strip()
    if not text:
        raise ValueError("El área de texto está vacía.")

    try:
        df = pd.read_csv(io.StringIO(text), sep="\t", dtype=str)
    except Exception as exc:
        raise ValueError(f"No se pudo parsear el texto pegado: {exc}") from exc

    df = df.rename(columns=_COL_MAP)

    missing = _REQUIRED - set(df.columns)
    if missing:
        raise ValueError(
            f"El texto no contiene las columnas mínimas requeridas: "
            f"{', '.join(sorted(missing))}.\n"
            f"Asegúrate de copiar desde la fila de encabezados de la hoja 'Registro'."
        )

    return _normalize(df)
