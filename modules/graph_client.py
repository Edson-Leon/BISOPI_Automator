# ═══════════════════════════════════════════════════════════
# BISOPI Automator v1.0 — Mayo 2026
# Autor: Edson Leon
# Cargo: Data Insights Lead — Operaciones
# Empresa: Bision Consulting
# Contacto: edson.leon@bisionconsulting.com
# ═══════════════════════════════════════════════════════════

"""
Autenticación y obtención de eventos del calendario vía Microsoft Graph API.

Responsabilidades de este módulo
──────────────────────────────────
  - Autenticar al usuario con MSAL (flujo interactive + silent con caché en disco).
  - Obtener la lista de eventos desde /me/calendarView con paginación automática.
  - Convertir los eventos del JSON de Graph al DataFrame BISOPI delegando en
    outlook_parser._process_event_list para no duplicar lógica de procesamiento.
  - Extraer el email del usuario desde los claims del token JWT.

El procesamiento posterior (validación, upload, histórico) lo realizan los
módulos existentes sin modificación.  No tocar ningún archivo de Archivo plano.
"""
from __future__ import annotations

import base64
import json
import os
import re as _re
import tempfile
from datetime import date, datetime, timedelta

import pandas as pd

try:
    import msal as _msal
    _MSAL_AVAILABLE = True
except ImportError:
    _MSAL_AVAILABLE = False

try:
    import requests as _requests
    _REQUESTS_AVAILABLE = True
except ImportError:
    _REQUESTS_AVAILABLE = False

from modules.outlook_parser import (
    _process_event_list,
    _round_down_15,
    _BISOPI_TAG,
    _ALLDAY_MINUTES,
)
from modules.loader import REGISTRO_COLUMNS

# Credenciales Azure AD — se leen en tiempo de importación desde el .env ya cargado
try:
    from config import AZURE_CLIENT_ID as _CLIENT_ID, AZURE_TENANT_ID as _TENANT_ID
except ImportError:
    _CLIENT_ID = ""
    _TENANT_ID = ""


# ── Constantes ────────────────────────────────────────────────────────────────

_SCOPES     = ["Calendars.Read"]
_CACHE_FILE = os.path.join(tempfile.gettempdir(), ".bisopi_msal_cache.json")
_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_RE_HTML    = _re.compile(r"<[^>]+>")


# ── Token cache ───────────────────────────────────────────────────────────────

def _load_cache() -> "_msal.SerializableTokenCache":
    cache = _msal.SerializableTokenCache()
    if os.path.isfile(_CACHE_FILE):
        try:
            with open(_CACHE_FILE, "r", encoding="utf-8") as fh:
                cache.deserialize(fh.read())
        except Exception:
            pass  # caché corrupta — empezar de cero
    return cache


def _save_cache(cache: "_msal.SerializableTokenCache") -> None:
    if cache.has_state_changed:
        try:
            with open(_CACHE_FILE, "w", encoding="utf-8") as fh:
                fh.write(cache.serialize())
        except Exception:
            pass  # no es crítico si no se puede guardar


# ── Helper interno: construir MSAL app ───────────────────────────────────────

def _make_app() -> "tuple[_msal.PublicClientApplication, _msal.SerializableTokenCache]":
    """Crea el PublicClientApplication con la caché de tokens cargada desde disco."""
    if not _MSAL_AVAILABLE:
        raise ImportError(
            "La librería 'msal' no está instalada. "
            "Ejecuta: pip install msal"
        )
    cache     = _load_cache()
    authority = f"https://login.microsoftonline.com/{_TENANT_ID}"
    app       = _msal.PublicClientApplication(
        _CLIENT_ID,
        authority=authority,
        token_cache=cache,
    )
    return app, cache


def _try_silent(app: "_msal.PublicClientApplication") -> "dict | None":
    """Intenta obtener un token en silencio desde la caché. Retorna None si no hay cuenta."""
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result
    return None


# ── Autenticación ─────────────────────────────────────────────────────────────

# Cadenas que identifican un error de Redirect URI mal configurado en Azure AD
_REDIRECT_MARKERS = ("aadsts900971", "900971", "redirect", "redirect_uri")


def authenticate_interactive() -> str:
    """
    Autentica al usuario abriendo el navegador (flujo interactivo de MSAL).

    Intenta primero recuperar un token en silencio desde la caché.
    Si no hay token válido, abre el navegador para que el usuario inicie sesión.

    Retorna
    -------
    str — access token JWT.

    Lanza
    -----
    ImportError  — si 'msal' no está instalado.
    RuntimeError — con mensaje de orientación si el error es de configuración en Azure AD
                   (redirect URI no registrado → sugerencia de usar código de dispositivo).
    RuntimeError — para cualquier otro error de autenticación.
    """
    app, cache = _make_app()

    # Intento silencioso primero
    result = _try_silent(app)

    if result is None:
        try:
            result = app.acquire_token_interactive(scopes=_SCOPES)
        except Exception as exc:
            exc_lower = str(exc).lower()
            if any(m in exc_lower for m in _REDIRECT_MARKERS):
                raise RuntimeError(
                    "Error de configuración en Azure AD — "
                    "intenta con el método de código de dispositivo"
                ) from exc
            raise RuntimeError(f"No se pudo autenticar con Microsoft: {exc}") from exc

    _save_cache(cache)

    if "access_token" not in result:
        err_desc = (result.get("error_description") or result.get("error") or "").lower()
        if any(m in err_desc for m in _REDIRECT_MARKERS):
            raise RuntimeError(
                "Error de configuración en Azure AD — "
                "intenta con el método de código de dispositivo"
            )
        raise RuntimeError(
            result.get("error_description") or result.get("error") or "Error desconocido"
        )

    return result["access_token"]


def initiate_device_flow() -> dict:
    """
    Inicia el Device Code Flow y devuelve el objeto flow completo.

    El dict retornado contiene al menos:
      user_code        — código de 8-9 letras que el usuario teclea en el portal.
      verification_uri — URL donde el usuario introduce el código.
      message          — instrucciones listas para mostrar al usuario.
      device_code      — token interno usado por complete_device_flow().

    Lanza
    -----
    ImportError  — si 'msal' no está instalado.
    RuntimeError — con error_description si Microsoft rechaza el inicio del flujo.
    """
    app, _cache = _make_app()
    flow = app.initiate_device_flow(scopes=_SCOPES)
    if "error" in flow:
        err = flow.get("error_description") or flow.get("error") or "Error desconocido"
        raise RuntimeError(err)
    return flow


def poll_device_flow(flow: dict) -> str | None:
    """
    Hace un único intento de canjear el device_code por un access token.

    A diferencia de acquire_token_by_device_flow() (que hace polling bloqueante),
    esta función realiza UNA sola petición HTTP al endpoint de token y retorna
    inmediatamente. Esto es necesario en Streamlit para no bloquear el hilo
    del servidor.

    Parámetros
    ----------
    flow : dict devuelto por initiate_device_flow().
           Debe contener al menos 'device_code' y 'interval'.

    Retorna
    -------
    str  — access token JWT si el usuario ya completó la autenticación.
    None — si el token aún no está listo (authorization_pending o slow_down).
           El llamador debe informar al usuario y reintentar más tarde.

    Lanza
    -----
    ImportError  — si 'requests' no está instalado.
    RuntimeError — si el flujo expiró, fue denegado u ocurrió otro error.
    """
    if not _REQUESTS_AVAILABLE:
        raise ImportError(
            "La librería 'requests' no está instalada. "
            "Ejecuta: pip install requests"
        )

    # MSAL incluye el endpoint correcto en el flow dict — usarlo directamente
    # evita problemas de formato con TENANT_ID (URL larga vs GUID vs dominio).
    token_url = (
        flow.get("token_endpoint")
        or f"https://login.microsoftonline.com/{_TENANT_ID}/oauth2/v2.0/token"
    )
    data = {
        "client_id":   _CLIENT_ID,
        "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": flow["device_code"],
    }

    resp = _requests.post(token_url, data=data, timeout=15)
    result = resp.json()

    if "access_token" in result:
        # Guardar en caché MSAL para que el token silent funcione en sesiones futuras
        try:
            app, cache = _make_app()
            # Usar el refresh_token para renovar la caché si está disponible
            refresh = result.get("refresh_token")
            if refresh:
                app.acquire_token_by_refresh_token(
                    refresh, scopes=_SCOPES
                )
                _save_cache(cache)
        except Exception:
            pass  # La caché es opcional; el token ya está en session_state
        return result["access_token"]

    error = result.get("error", "")

    # Aún esperando — el usuario no terminó de autenticarse
    if error in ("authorization_pending", "slow_down"):
        return None

    # Cualquier otro error es definitivo
    err_msg = result.get("error_description") or error or "Error desconocido"
    raise RuntimeError(err_msg)


# ── Información del usuario ──────────────────────────────────────────────────

def get_user_info(token: str) -> dict:
    """
    Obtiene el email y nombre del usuario desde /me en Microsoft Graph.

    Requiere el permiso User.Read (incluido por defecto en cuentas corporativas).

    Retorna
    -------
    dict con claves:
      "email" : mail o userPrincipalName del usuario.
      "name"  : displayName del usuario.
    """
    # Intentar obtener nombre y email desde Graph API /me
    if _REQUESTS_AVAILABLE:
        try:
            response = _requests.get(
                f"{_GRAPH_BASE}/me",
                headers={"Authorization": f"Bearer {token}"},
                timeout=10,
            )
            if response.ok:
                data = response.json()
                email = data.get("mail") or data.get("userPrincipalName") or ""
                name  = data.get("displayName", "")
                if email:
                    return {"email": email, "name": name}
        except Exception:
            pass  # fallback al decode JWT

    # Fallback: extraer email desde los claims del JWT (sin llamada extra a la API)
    email = get_user_email(token)
    return {"email": email, "name": ""}


def get_user_email(token: str) -> str:
    """
    Extrae el email del usuario desde los claims del token JWT.

    No requiere el permiso User.Read ni llamada adicional a la API —
    decodifica directamente el payload base64url del JWT.

    Retorna el email encontrado o cadena vacía si los claims no contienen ninguno.
    """
    try:
        payload_part = token.split(".")[1]
        # Añadir padding base64 si es necesario
        padding = 4 - len(payload_part) % 4
        if padding != 4:
            payload_part += "=" * padding
        claims = json.loads(
            base64.urlsafe_b64decode(payload_part).decode("utf-8")
        )
        return (
            claims.get("preferred_username")
            or claims.get("upn")
            or claims.get("email")
            or ""
        )
    except Exception:
        return ""


# ── Obtener eventos del calendario ────────────────────────────────────────────

def get_calendar_events(
    token: str,
    start_date: date,
    end_date: date,
) -> list[dict]:
    """
    Obtiene todos los eventos del calendario personal del usuario desde Graph API.

    Utiliza /me/calendarView con paginación automática (sigue @odata.nextLink
    hasta obtener todos los eventos del rango).

    Parámetros
    ----------
    token       : access token devuelto por authenticate().
    start_date  : primer día del rango (inclusive).
    end_date    : último día del rango (inclusive, hasta las 23:59:59).

    Retorna
    -------
    Lista de dicts con el esquema de Graph API (campo 'value').

    Lanza
    -----
    ImportError  — si 'requests' no está instalado.
    RuntimeError — para errores HTTP (401, 403, 5xx, etc.).
    """
    if not _REQUESTS_AVAILABLE:
        raise ImportError(
            "La librería 'requests' no está instalada. "
            "Ejecuta: pip install requests"
        )

    start_iso = datetime(
        start_date.year, start_date.month, start_date.day, 0, 0, 0
    ).strftime("%Y-%m-%dT%H:%M:%S")
    end_iso = datetime(
        end_date.year, end_date.month, end_date.day, 23, 59, 59
    ).strftime("%Y-%m-%dT%H:%M:%S")

    url: str | None = f"{_GRAPH_BASE}/me/calendarView"
    params: dict | None = {
        "startDateTime": start_iso,
        "endDateTime":   end_iso,
        "$select":       "subject,start,end,isAllDay,body",
        "$top":          "100",
        "$orderby":      "start/dateTime",
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer":        'outlook.timezone="UTC"',
    }

    events: list[dict] = []

    while url:
        resp = _requests.get(url, headers=headers, params=params, timeout=30)

        if resp.status_code == 401:
            raise RuntimeError(
                "❌ Token expirado (401 Unauthorized). "
                "Haz clic en **Conectar con Microsoft** para renovar la sesión."
            )
        if resp.status_code == 403:
            raise RuntimeError(
                "❌ Sin permisos para leer el calendario (403 Forbidden). "
                "Verifica que la app de Azure tiene el permiso 'Calendars.Read' "
                "y que el administrador lo ha consentido."
            )
        if resp.status_code >= 500:
            raise RuntimeError(
                f"❌ Error en el servidor de Microsoft ({resp.status_code}). "
                "Inténtalo de nuevo en unos minutos."
            )
        if not resp.ok:
            raise RuntimeError(
                f"❌ Error al obtener eventos del calendario ({resp.status_code}): "
                f"{resp.text[:200]}"
            )

        data = resp.json()
        events.extend(data.get("value", []))
        url    = data.get("@odata.nextLink")  # None si no hay más páginas
        params = None                          # los params van codificados en nextLink

    return events


# ── Convertir eventos de Graph al DataFrame BISOPI ────────────────────────────

def events_to_dataframe(
    events: list[dict],
    week_start: date,
    week_end: date,
) -> pd.DataFrame:
    """
    Convierte la lista de eventos de Graph API al DataFrame compatible con BISOPI.

    Normaliza cada evento al esquema interno all_evts y delega en
    outlook_parser._process_event_list para aplicar la misma lógica de
    clasificación y ordenación que parse_ics (sin duplicar código).

    Parámetros
    ----------
    events      : lista de eventos devuelta por get_calendar_events().
    week_start  : primer día del rango (filtrado de seguridad).
    week_end    : último día del rango.

    Retorna
    -------
    DataFrame con REGISTRO_COLUMNS.
    """
    all_evts: list[dict] = []

    for evt in events:
        summary   = evt.get("subject") or ""
        is_allday = bool(evt.get("isAllDay", False))

        # Parsear fechas — Graph devuelve ISO 8601 en UTC cuando se usa el header Prefer.
        # Nota: Graph puede devolver hasta 7 decimales en los segundos
        # (ej. "2026-05-25T13:00:00.0000000"), pero datetime.fromisoformat solo
        # acepta hasta 6 en Python ≤ 3.10. Se trunca a 6 antes de parsear.
        start_raw = (evt.get("start") or {}).get("dateTime", "")
        end_raw   = (evt.get("end")   or {}).get("dateTime", "")

        def _parse_dt(s: str) -> datetime:
            s = s.rstrip("Z").split("+")[0]   # quitar zona horaria
            if "." in s:
                base, frac = s.split(".", 1)
                s = f"{base}.{frac[:6]}"      # truncar a máx. 6 decimales (microsegundos)
            return datetime.fromisoformat(s)

        try:
            dtstart = _parse_dt(start_raw)
        except (ValueError, AttributeError):
            continue

        try:
            dtend = _parse_dt(end_raw) if end_raw else None
        except (ValueError, AttributeError):
            dtend = None

        event_date = dtstart.date()
        if not (week_start <= event_date <= week_end):
            continue

        if is_allday:
            total_minutes = _ALLDAY_MINUTES
        elif dtend is not None and dtend > dtstart:
            total_minutes = int((dtend - dtstart).total_seconds() / 60)
        else:
            total_minutes = 60  # fallback: 1 hora

        horas, minutos = _round_down_15(total_minutes)
        if horas == 0 and minutos == 0:
            continue

        dtend_eff = dtend if (dtend is not None and dtend > dtstart) else (
            dtstart + timedelta(hours=horas, minutes=minutos)
        )

        # Extraer descripción del body (HTML o texto plano)
        body         = evt.get("body") or {}
        body_content = body.get("content", "")
        if body.get("contentType", "text").lower() == "html":
            body_content = _RE_HTML.sub("\n", body_content)
            body_content = (
                body_content
                .replace("&nbsp;", " ")
                .replace("&lt;",   "<")
                .replace("&gt;",   ">")
                .replace("&amp;",  "&")
                .strip()
            )

        all_evts.append({
            "summary":         summary,
            "is_bisopi":       bool(_BISOPI_TAG.search(summary)),
            "is_allday":       is_allday,
            "dtstart":         dtstart,
            "dtend_eff":       dtend_eff,
            "event_date":      event_date,
            "horas":           horas,
            "minutos":         minutos,
            "description_raw": body_content,
        })

    if not all_evts:
        return pd.DataFrame(columns=REGISTRO_COLUMNS)

    # Delegar en la misma función que usa parse_ics
    return _process_event_list(all_evts)
