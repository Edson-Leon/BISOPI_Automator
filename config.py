# ═══════════════════════════════════════════════════════════
# BISOPI Automator v1.0 — Mayo 2026
# Autor: Edson Leon
# Cargo: Data Insights Lead — Operaciones
# Empresa: Bision Consulting
# Contacto: edson.leon@bisionconsulting.com
# ═══════════════════════════════════════════════════════════

from __future__ import annotations

from dotenv import load_dotenv
import os

load_dotenv()

EMAIL_COLABORADOR: str = os.getenv("EMAIL_COLABORADOR", "")
BISOPI_API_TOKEN: str  = os.getenv("BISOPI_API_TOKEN",  "")

# Opcional — ruta local a la plantilla.  None si no está configurada.
_raw_path = os.getenv("PLANTILLA_PATH", "").strip()
PLANTILLA_PATH: str | None = _raw_path if _raw_path else None

API_BASE_URL = "https://bisopi-open-hkdhcucjeuadafhc.eastus-01.azurewebsites.net"
API_ENDPOINT = f"{API_BASE_URL}/api/ImputarHoras"

# Opcional — credenciales Azure AD para Submodo B (Microsoft Graph API).
# Deja vacíos si aún no tienes las credenciales o no usas el Submodo B.
AZURE_CLIENT_ID: str = os.getenv("AZURE_CLIENT_ID", "").strip()
AZURE_TENANT_ID: str = os.getenv("AZURE_TENANT_ID", "").strip()


def has_local_path() -> bool:
    """
    True si PLANTILLA_PATH está configurada en .env y el archivo existe en disco.
    False si no está configurada o si la ruta no apunta a un archivo real.
    """
    return bool(PLANTILLA_PATH and os.path.isfile(PLANTILLA_PATH))


def has_graph_access() -> bool:
    """
    True si AZURE_CLIENT_ID y AZURE_TENANT_ID están ambos configurados en .env.
    False si alguno de los dos está vacío.
    """
    return bool(AZURE_CLIENT_ID and AZURE_TENANT_ID)
