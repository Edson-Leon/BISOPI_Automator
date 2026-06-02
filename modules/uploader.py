# ═══════════════════════════════════════════════════════════
# BISOPI Automator v1.0 — Mayo 2026
# Autor: Edson Leon
# Cargo: Data Insights Lead — Operaciones
# Empresa: Bision Consulting
# Contacto: edson.leon@bisionconsulting.com
# ═══════════════════════════════════════════════════════════

"""
Envía cada fila con Estado == "Pendiente" al endpoint POST /api/ImputarHoras de BISOPI.
Actualiza Estado y RespuestaAPI en el DataFrame según la respuesta.
"""
from __future__ import annotations

from datetime import datetime, timedelta

import requests
import pandas as pd

from config import API_ENDPOINT


class TokenError(Exception):
    """Lanzada cuando la API devuelve 401 Unauthorized."""


# ── Construcción del payload ───────────────────────────────────────────────────

def _build_payload(row: pd.Series, email: str) -> dict:
    tipo = str(row.get("TipoHora", "")).strip()

    payload: dict = {
        "proyecto":         str(row["Proyecto"]).strip(),
        "grupoTarea":       str(row["GrupoTarea"]).strip(),
        "tarea":            str(row["Tarea"]).strip(),
        "tipoHora":         tipo,
        "fechaRegistro":    str(row["FechaRegistro"]).strip(),
        "tiempoEjecutado": {
            "horas":   int(row["Horas"]),
            "minutos": int(row["Minutos"]),
        },
        "emailColaborador": email,
    }

    comentario = str(row.get("Comentario", "")).strip()
    if comentario:
        payload["comentario"] = comentario

    if tipo == "Adicional":
        hora_inicio_str = str(row.get("HoraInicio", "")).strip()
        inicio = datetime.strptime(hora_inicio_str, "%H:%M")
        delta  = timedelta(hours=int(row["Horas"]), minutes=int(row["Minutos"]))
        hora_fin = (inicio + delta).strftime("%H:%M")
        payload["horaInicio"] = hora_inicio_str
        payload["horaFin"]    = hora_fin
        # comentario es obligatorio para Adicional; si no venía arriba, incluir vacío
        if "comentario" not in payload:
            payload["comentario"] = ""

    return payload


# ── Extracción de mensaje desde la respuesta ──────────────────────────────────

def _extract_msg(resp: requests.Response) -> str:
    """Intenta obtener un texto legible del cuerpo JSON o lo devuelve plano."""
    try:
        body = resp.json()
        for key in ("message", "msg", "error", "detail", "data"):
            val = body.get(key)
            if val:
                return str(val)
    except Exception:
        pass
    return resp.text.strip() or f"HTTP {resp.status_code}"


# ── Función principal ─────────────────────────────────────────────────────────

def upload(
    df: pd.DataFrame,
    email: str,
    token: str,
    on_progress=None,
) -> pd.DataFrame:
    """
    Itera sobre las filas con Estado == "Pendiente" y las envía a BISOPI.

    Parámetros
    ----------
    df           : DataFrame completo con REGISTRO_COLUMNS.
    email        : EMAIL_COLABORADOR del archivo .env.
    token        : BISOPI_API_TOKEN del archivo .env.
    on_progress  : callable(current: int, total: int, label: str) — actualiza UI.

    Retorna el DataFrame con Estado y RespuestaAPI actualizados.
    Lanza TokenError si el servidor responde 401.
    """
    df = df.copy()
    pending_idx: list = df.index[df["Estado"] == "Pendiente"].tolist()
    total = len(pending_idx)

    if total == 0:
        return df

    headers = {
        "X-API-Token":  token,
        "Content-Type": "application/json",
    }

    # Priorizar laborales sobre adicionales en el envío
    df_pendiente   = df[df["Estado"] == "Pendiente"].copy()
    df_laborales   = df_pendiente[df_pendiente["TipoHora"] == "Laboral"]
    df_adicionales = df_pendiente[df_pendiente["TipoHora"] == "Adicional"]
    df_ordered     = pd.concat([df_laborales, df_adicionales])
    ordered_idx    = df_ordered.index.tolist()
    total          = len(ordered_idx)

    for i, idx in enumerate(ordered_idx, start=1):
        row = df.loc[idx]
        label = f"{str(row['Tarea'])[:45]} · {row['FechaRegistro']}"

        # Notificar inicio de esta fila
        if on_progress:
            on_progress(i - 1, total, label)

        # ── Envío HTTP ────────────────────────────────────────────────────────
        try:
            payload = _build_payload(row, email)
            resp = requests.post(API_ENDPOINT, json=payload, headers=headers, timeout=30)

        except requests.exceptions.Timeout:
            df.at[idx, "Estado"]       = "❌ Error"
            df.at[idx, "RespuestaAPI"] = "Tiempo de espera agotado — revisa tu conexión a internet"
            if on_progress:
                on_progress(i, total, label)
            continue

        except requests.exceptions.RequestException as exc:
            df.at[idx, "Estado"]       = "❌ Error"
            df.at[idx, "RespuestaAPI"] = f"Error de conexión: {exc}"
            if on_progress:
                on_progress(i, total, label)
            continue

        # ── Manejo de códigos de respuesta ────────────────────────────────────
        if resp.status_code == 401:
            raise TokenError("Token inválido o expirado (401 Unauthorized)")

        if resp.status_code == 200:
            df.at[idx, "Estado"]       = "✅ Cargado"
            df.at[idx, "RespuestaAPI"] = ""

        elif resp.status_code == 400:
            df.at[idx, "Estado"]       = "❌ Error"
            df.at[idx, "RespuestaAPI"] = _extract_msg(resp) or "Error 400 — datos inválidos"

        else:
            df.at[idx, "Estado"]       = "❌ Error"
            df.at[idx, "RespuestaAPI"] = f"Error inesperado del servidor ({resp.status_code})"

        # Notificar fin de esta fila
        if on_progress:
            on_progress(i, total, label)

    return df
