# ═══════════════════════════════════════════════════════════
# BISOPI Automator v1.0 — Mayo 2026
# Autor: Edson Leon
# Cargo: Data Insights Lead — Operaciones
# Empresa: Bision Consulting
# Contacto: edson.leon@bisionconsulting.com
# ═══════════════════════════════════════════════════════════

"""
Valida reglas de negocio sobre el DataFrame cargado.
Actualiza Estado y RespuestaAPI sin llamar a la API.
"""
from datetime import date, datetime, timedelta
import re

import pandas as pd

# ── Festivos Colombia 2025-2026 (Ley Emiliani aplicada) ───────────────────────
_FESTIVOS_CO: set[date] = {
    # 2025
    date(2025,  1,  1),  # Año Nuevo
    date(2025,  1,  6),  # Reyes Magos
    date(2025,  3, 24),  # San José
    date(2025,  4, 17),  # Jueves Santo
    date(2025,  4, 18),  # Viernes Santo
    date(2025,  5,  1),  # Día del Trabajo
    date(2025,  6,  2),  # Ascensión
    date(2025,  6, 23),  # Corpus Christi
    date(2025,  6, 30),  # Sagrado Corazón
    date(2025,  7,  7),  # San Pedro y San Pablo
    date(2025,  7, 20),  # Independencia
    date(2025,  8,  7),  # Batalla de Boyacá
    date(2025,  8, 18),  # Asunción de la Virgen
    date(2025, 10, 13),  # Día de la Raza
    date(2025, 11,  3),  # Todos los Santos
    date(2025, 11, 17),  # Independencia de Cartagena
    date(2025, 12,  8),  # Inmaculada Concepción
    date(2025, 12, 25),  # Navidad
    # 2026
    date(2026,  1,  1),  # Año Nuevo
    date(2026,  1, 12),  # Reyes Magos
    date(2026,  3, 23),  # San José
    date(2026,  4,  2),  # Jueves Santo
    date(2026,  4,  3),  # Viernes Santo
    date(2026,  5,  1),  # Día del Trabajo
    date(2026,  5, 18),  # Ascensión
    date(2026,  6,  8),  # Corpus Christi
    date(2026,  6, 15),  # Sagrado Corazón
    date(2026,  6, 29),  # San Pedro y San Pablo
    date(2026,  7, 20),  # Independencia
    date(2026,  8,  7),  # Batalla de Boyacá
    date(2026,  8, 17),  # Asunción de la Virgen
    date(2026, 10, 12),  # Día de la Raza
    date(2026, 11,  2),  # Todos los Santos
    date(2026, 11, 16),  # Independencia de Cartagena
    date(2026, 12,  8),  # Inmaculada Concepción
    date(2026, 12, 25),  # Navidad
}

_RE_FECHA        = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_RE_HORA         = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
_MINUTOS_VALIDOS = {0, 15, 30, 45}
_LIMITE_LAB_MIN  = 40 * 60  # 2400 minutos

# Estados que el validador nunca debe modificar (ya tienen su estado final)
_PRESERVED_STATES: frozenset = frozenset({"✅ Cargado", "⚠ Sin clasificar", "⚠ Incompleto"})


# ── Helpers ───────────────────────────────────────────────────────────────────

def _hhmm_to_min(s: str) -> int | None:
    m = _RE_HORA.match(str(s).strip())
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def _is_week_potentially_closed(fecha: date) -> bool:
    """
    True si la semana de 'fecha' ya superó su fecha de cierre.
    Cierre = el lunes siguiente a esa semana a las 12:30.
    Si ese lunes es festivo colombiano, el cierre se corre al martes.
    """
    hoy = date.today()
    now = datetime.now()

    fecha_monday = fecha - timedelta(days=fecha.weekday())
    hoy_monday   = hoy  - timedelta(days=hoy.weekday())

    if fecha_monday >= hoy_monday:
        return False  # misma semana o futura

    closure_monday = fecha_monday + timedelta(days=7)
    closure_day = (
        closure_monday + timedelta(days=1)
        if closure_monday in _FESTIVOS_CO
        else closure_monday
    )
    closure_dt = datetime(closure_day.year, closure_day.month, closure_day.day, 12, 30)
    return now >= closure_dt


def _append_msg(df: pd.DataFrame, idx, msg: str, is_error: bool) -> None:
    current = str(df.at[idx, "RespuestaAPI"]).strip()
    df.at[idx, "RespuestaAPI"] = f"{current} | {msg}".strip(" |") if current else msg
    if is_error:
        df.at[idx, "Estado"] = "❌ Error"


# ── Validaciones por fila ─────────────────────────────────────────────────────

def _row_errors(row: pd.Series) -> list[str]:
    errors: list[str] = []

    # Campos obligatorios
    for campo in ("Proyecto", "GrupoTarea", "Tarea", "TipoHora"):
        if not str(row.get(campo, "")).strip():
            errors.append(f"{campo} es obligatorio")

    # TipoHora válido
    tipo = str(row.get("TipoHora", "")).strip()
    if tipo and tipo not in ("Laboral", "Adicional"):
        errors.append("TipoHora debe ser 'Laboral' o 'Adicional'")

    # FechaRegistro
    fecha_str = str(row.get("FechaRegistro", "")).strip()
    fecha_valida = False
    if not _RE_FECHA.match(fecha_str):
        errors.append("FechaRegistro debe tener formato YYYY-MM-DD")
    else:
        try:
            datetime.strptime(fecha_str, "%Y-%m-%d")
            fecha_valida = True
        except ValueError:
            errors.append("FechaRegistro no es una fecha válida")

    # Horas
    try:
        horas = int(row.get("Horas", 0))
        if horas < 0:
            errors.append("Horas debe ser >= 0")
    except (ValueError, TypeError):
        errors.append("Horas debe ser un número entero")
        horas = 0

    # Minutos
    try:
        minutos = int(row.get("Minutos", 0))
        if minutos not in _MINUTOS_VALIDOS:
            errors.append("Minutos debe ser 0, 15, 30 o 45")
    except (ValueError, TypeError):
        errors.append("Minutos debe ser un número entero (0, 15, 30 o 45)")
        minutos = 0

    # No ambos cero
    if horas == 0 and minutos == 0:
        errors.append("Horas y Minutos no pueden ser ambos 0")

    # Reglas específicas de Adicional
    if tipo == "Adicional":
        inicio_str = str(row.get("HoraInicio", "")).strip()
        comentario = str(row.get("Comentario", "")).strip()

        inicio_min = _hhmm_to_min(inicio_str) if inicio_str else None

        if not inicio_str:
            errors.append("HoraInicio es obligatorio para horas Adicionales")
        elif inicio_min is None:
            errors.append("HoraInicio debe tener formato HH:MM (ej. 18:00)")

        if not comentario:
            errors.append("Comentario es obligatorio para horas Adicionales")

        # Verificar que inicio + duración no supere la medianoche
        if inicio_min is not None:
            fin_min = inicio_min + horas * 60 + minutos
            if fin_min > 23 * 60 + 59:
                errors.append(
                    "La actividad adicional supera la medianoche — "
                    "divide el registro en dos días"
                )

    return errors


# ── Validación global: límite 40h laborales por semana ───────────────────────

def _check_40h_limit(df: pd.DataFrame) -> None:
    dates    = pd.to_datetime(df["FechaRegistro"], errors="coerce")
    iso      = dates.dt.isocalendar()
    week_key = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
    week_key[dates.isna()] = ""
    df["_wk"] = week_key

    for wk, group in df.groupby("_wk"):
        if not wk:
            continue
        lab_rows = group[
            (group["TipoHora"].str.strip() == "Laboral") &
            (~group["Estado"].isin(_PRESERVED_STATES))
        ].index
        running  = 0
        for idx in lab_rows:
            row_min  = int(df.at[idx, "Horas"]) * 60 + int(df.at[idx, "Minutos"])
            running += row_min
            if running > _LIMITE_LAB_MIN:
                h, m = divmod(running, 60)
                msg  = f"Límite semanal superado: acumulas {h}h {m:02d}m laborales esta semana (máx. 40h)"
                _append_msg(df, idx, msg, is_error=True)

    df.drop(columns=["_wk"], inplace=True)


# ── Validación global: semana potencialmente cerrada ─────────────────────────

def _check_closed_week(df: pd.DataFrame) -> None:
    warn = "⚠ Semana potencialmente cerrada — verifica con la persona responsable"
    for idx, row in df.iterrows():
        if df.at[idx, "Estado"] == "❌ Error" or df.at[idx, "Estado"] in _PRESERVED_STATES:
            continue
        fecha_str = str(row.get("FechaRegistro", "")).strip()
        if not _RE_FECHA.match(fecha_str):
            continue
        try:
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d").date()
        except ValueError:
            continue
        if _is_week_potentially_closed(fecha):
            _append_msg(df, idx, warn, is_error=False)


# ── Punto de entrada ──────────────────────────────────────────────────────────

def validate(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Estado" not in df.columns:
        df["Estado"] = "Pendiente"
    else:
        df["Estado"] = df["Estado"].where(df["Estado"].isin(_PRESERVED_STATES), "Pendiente")
    # Preservar RespuestaAPI de estados especiales; resetear el resto
    _mask_preserve = df["Estado"].isin(_PRESERVED_STATES)
    df.loc[~_mask_preserve, "RespuestaAPI"] = ""

    # 1. Errores por fila (se omiten filas con estado preservado)
    for idx, row in df.iterrows():
        if df.at[idx, "Estado"] in _PRESERVED_STATES:
            continue
        errs = _row_errors(row)
        if errs:
            df.at[idx, "Estado"]       = "❌ Error"
            df.at[idx, "RespuestaAPI"] = " | ".join(errs)

    # 2. Límite 40h (puede agregar errores a filas ya marcadas)
    _check_40h_limit(df)

    # 3. Advertencia semana cerrada (solo filas sin error)
    _check_closed_week(df)

    return df
