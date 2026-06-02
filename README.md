# BISOPI Automator

Herramienta interna de Bision Consulting para automatizar el registro semanal de horas en BISOPI. Elimina la carga manual tarea por tarea reemplazandola por un proceso de una sola accion: preparas los registros desde tu fuente de datos, la app los valida contra las reglas de negocio y los envia a la API de BISOPI en un solo paso.

---

## Fuentes de datos

| Fuente | Estado |
|---|---|
| Plantilla Excel (archivo `.xlsx` o copy-paste) | Disponible |
| Agenda Outlook — archivo ICS exportado | Disponible |
| Agenda Outlook — Microsoft Graph API (conexion directa) | Disponible |
| Azure DevOps | En desarrollo (V2) |

---

## Requisitos

- Python 3.10 o superior
- Las dependencias listadas en `requirements.txt` (`streamlit`, `pandas`, `openpyxl`, `requests`, `python-dotenv`, `msal`)
- Para la integracion con Outlook via Graph API: credenciales de Azure AD entregadas por IT (`AZURE_CLIENT_ID`, `AZURE_TENANT_ID`)

---

## Instalacion (version local)

```bash
# 1. Clonar el repositorio
git clone https://github.com/Edson-Leon/BISOPI_Automator.git
cd BISOPI_Automator

# 2. Instalar dependencias
python -m pip install -r requirements.txt

# 3. Configurar credenciales
cp .env.example .env
# Editar .env con los valores reales
```

En Windows tambien puedes ejecutar `install.bat` en lugar del paso 2.

---

## Configuracion del archivo `.env`

| Variable | Requerida | Descripcion |
|---|---|---|
| `EMAIL_COLABORADOR` | Si (solo local) | Email corporativo del usuario. En cloud se obtiene del login con Microsoft. |
| `BISOPI_API_TOKEN` | Si | Token de autenticacion de la API de BISOPI. Solicitar a Carlos Manrique. |
| `PLANTILLA_PATH` | No | Ruta completa al archivo Excel de plantilla. Si se configura, la app lee y escribe directamente en disco. |
| `AZURE_CLIENT_ID` | No* | Client ID de la app registrada en Azure AD. Requerido para la integracion con Outlook via Graph API. |
| `AZURE_TENANT_ID` | No* | Tenant ID de Azure AD. |
| `BISOPI_ENV` | No | `local` (por defecto) o `cloud`. Controla el modo de despliegue. |

Copia `.env.example` como punto de partida.

---

## Ejecucion

**Version local**

```bash
python -m streamlit run main.py
```

O ejecutar `run.bat` en Windows. La app abre automaticamente en `http://localhost:8501`.

**Version Cloud**

Acceso directo desde el navegador (sin instalacion):

```
https://bisopi-automator.streamlit.app
```

El acceso es privado. Los usuarios deben ser invitados por el administrador. Las credenciales se configuran en **Streamlit Cloud → Settings → Secrets** (ver `docs/DISTRIBUTION.md`).

---

## Estructura del proyecto

```
BISOPI_Automator/
├── main.py                        # Interfaz Streamlit y logica de presentacion
├── config.py                      # Variables de entorno y deteccion del modo de despliegue
├── requirements.txt               # Dependencias Python
├── install.bat                    # Instalacion de dependencias en Windows
├── run.bat                        # Arranque de la app en Windows
├── .env.example                   # Plantilla de configuracion
├── Iniciar_BISOPI_Automator.vbs   # Acceso directo silencioso para escritorio Windows
├── modules/
│   ├── loader.py                  # Carga desde Excel y copy-paste, esquema REGISTRO_COLUMNS
│   ├── validator.py               # Reglas de negocio (40h, semana cerrada, campos requeridos)
│   ├── uploader.py                # Envio a POST /api/ImputarHoras
│   ├── file_manager.py            # Guardado en disco o serializacion a bytes (cloud)
│   ├── history.py                 # Escritura al Historico en la plantilla Excel
│   ├── outlook_parser.py          # Parseo de archivos ICS y exportacion a plantilla
│   └── graph_client.py            # Autenticacion MSAL y obtencion de eventos via Graph API
├── data/
│   └── Plantilla_BISOPI_Automator.xlsx   # Plantilla oficial con hojas Registro, Historico, Catalogos
└── docs/
    ├── DISTRIBUTION.md            # Detalle de las dos modalidades de distribucion
    └── ARCHITECTURE.md            # Referencia tecnica de arquitectura para desarrolladores
```

---

## Documentacion adicional

- **`docs/DISTRIBUTION.md`** — comparacion detallada entre la version Cloud y la version local, scripts de instalacion y configuracion de Streamlit Cloud Secrets.
- **`docs/ARCHITECTURE.md`** — diagrama de flujo, referencia de modulos, decisiones tecnicas, API de BISOPI y guia para agregar nuevas fuentes de datos.

---

*BISOPI Automator v1.0 — Mayo 2026 | Edson David Leon Araujo | Bision Consulting*
