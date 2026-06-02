# BISOPI Automator — Modalidades de Distribución

## Descripción general

BISOPI Automator está disponible en dos modalidades de distribución según el perfil del usuario y el nivel de funcionalidad requerido. Ambas versiones comparten el mismo código base y se diferencian únicamente por la variable de entorno `BISOPI_ENV`.

---

## Modalidad 1 — Streamlit Community Cloud (acceso web privado)

La app corre en los servidores de Streamlit y los usuarios acceden por URL desde cualquier navegador. El acceso es privado — solo usuarios invitados por el administrador pueden entrar.

### Características

- URL pública pero privada: `https://bisopi-automator.streamlit.app`
- Sin instalación para el usuario — solo abrir el navegador
- Variables de entorno configuradas en el panel de Streamlit Cloud
- La plantilla Excel siempre se descarga tras cada sesión — no hay archivo persistente en el servidor
- Autenticación con Outlook solo por Device Code Flow (código de 8 caracteres en `microsoft.com/devicelogin`)
- Actualizaciones automáticas — cualquier push al repositorio actualiza la app

### Limitaciones

- Sin ruta local configurable (`PLANTILLA_PATH` no aplica)
- El flujo interactivo de Outlook (ventana emergente) no está disponible
- Requiere conexión a internet

### Configuración de variables de entorno

Las variables se configuran en el panel de Streamlit Cloud, en la sección **Settings → Secrets**:

```toml
EMAIL_COLABORADOR = "edson.leon@bisionconsulting.com"
BISOPI_API_TOKEN  = "tu-token-aqui"
AZURE_CLIENT_ID   = "tu-client-id-aqui"
AZURE_TENANT_ID   = "tu-tenant-id-aqui"
BISOPI_ENV        = "cloud"
```

### Ideal para

Usuarios que no quieren instalar nada, acceso desde múltiples dispositivos, equipos distribuidos.

---

## Modalidad 2 — Instalación local (scripts de instalación)

La app corre en el ordenador del usuario. Se distribuye como una carpeta comprimida con scripts de instalación y ejecución. Requiere Python 3.10 o superior instalado una vez.

### Características

- Corre en `http://localhost:8501` — sin dependencia de internet para el uso básico
- Variables de entorno en archivo `.env` propio de cada usuario
- `PLANTILLA_PATH` disponible — la app lee y escribe directamente sobre el archivo Excel del usuario sin necesidad de cargar ni descargar
- Ambos métodos de autenticación con Outlook disponibles (ventana interactiva y Device Code Flow)
- Actualizaciones manuales — el usuario reemplaza los archivos cuando hay nueva versión

### Instalación (una sola vez)

```
1. Descomprimir el paquete en cualquier carpeta
2. Copiar .env.example → .env y configurar credenciales
3. Ejecutar install.bat
```

### Uso diario

```
Doble clic en run.bat → se abre la app en el navegador automáticamente
```

### Archivo .env

```dotenv
EMAIL_COLABORADOR=edson.leon@bisionconsulting.com
BISOPI_API_TOKEN=tu-token-aqui
PLANTILLA_PATH=C:/Users/TuUsuario/bisopi/Plantilla_BISOPI_Automator.xlsx
AZURE_CLIENT_ID=tu-client-id-aqui
AZURE_TENANT_ID=tu-tenant-id-aqui
BISOPI_ENV=local
```

### Contenido del paquete distribuible

```
BISOPI_Automator/
├── main.py
├── config.py
├── modules/
│   ├── loader.py
│   ├── validator.py
│   ├── uploader.py
│   ├── history.py
│   ├── file_manager.py
│   ├── outlook_parser.py
│   └── graph_client.py
├── requirements.txt
├── .env.example
├── Plantilla_BISOPI_Automator.xlsx
├── install.bat
└── run.bat
```

### Ideal para

Usuarios técnicos, quienes quieren la funcionalidad completa de archivo local, uso sin conexión a internet.

---

## Comparación rápida

| Factor | Cloud | Local |
|---|---|---|
| Instalación | Ninguna | Python + `install.bat` (una vez) |
| Acceso | URL desde cualquier dispositivo | Solo el ordenador donde está instalado |
| Plantilla Excel | Siempre descargar/cargar | Lectura/escritura directa en disco |
| Outlook — flujo interactivo | ❌ No disponible | ✅ Disponible |
| Outlook — Device Code Flow | ✅ Disponible | ✅ Disponible |
| Actualizaciones | Automáticas con cada push a GitHub | Manuales — reemplazar archivos |
| Requiere internet | Sí | Solo para API de BISOPI y Outlook |
| Control de acceso | Invitación por email o cuenta GitHub | Quien tenga el paquete instalado |
| `PLANTILLA_PATH` | ❌ No aplica | ✅ Disponible |
| Variable `BISOPI_ENV` | `cloud` | `local` |

---

## Scripts de instalación (Modalidad Local)

### install.bat

```batch
@echo off
echo Verificando Python...
python --version
if errorlevel 1 (
    echo ERROR: Python no encontrado. Instala Python 3.10 o superior desde https://www.python.org
    pause
    exit /b 1
)
echo Instalando dependencias...
python -m pip install -r requirements.txt
echo.
echo Instalacion completa.
echo Copia .env.example a .env y configura tus credenciales.
echo Luego ejecuta run.bat para arrancar la app.
pause
```

### run.bat

```batch
@echo off
echo Iniciando BISOPI Automator...
python -m streamlit run main.py
```

---

*BISOPI Automator v1.0 — Mayo 2026 | Edson David Leon Araujo | Bision Consulting*
