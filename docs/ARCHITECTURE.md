# BISOPI Automator — Referencia de Arquitectura

## Vision general

BISOPI Automator es una aplicacion Streamlit de proceso unico sin backend propio. Toda la logica corre en el proceso Python que sirve la interfaz. La arquitectura establece una separacion estricta entre las fuentes de datos (como entran los registros al sistema), el pipeline interno de procesamiento (validacion, envio, historico) y la capa de interfaz (main.py). Cada fuente de datos es un modulo independiente que normaliza su entrada al esquema interno compartido (`REGISTRO_COLUMNS`), a partir del cual todas las fuentes comparten exactamente la misma logica de validacion, envio y persistencia. La aplicacion soporta dos modos de despliegue — local y Streamlit Cloud — controlados por la variable de entorno `BISOPI_ENV`, que activa comportamientos exclusivos del cloud (login obligatorio con Microsoft, modo descarga forzado para archivos) sin duplicar ninguna logica de negocio.

---

## Diagrama de flujo

```
Entrada del usuario
    |
    +-- Archivo Excel / copy-paste ──> loader.py ──────────+
    |                                                       |
    +-- Archivo ICS ────────────> outlook_parser.py ────────+
    |                                                       |
    +-- Microsoft Graph API ───> graph_client.py ──────────+
                                        |                   |
                                        v                   v
                                 DataFrame REGISTRO_COLUMNS
                                        |
                                   validator.py
                                 (reglas de negocio)
                                        |
                                   main.py UI
                                 (tabla editable)
                                        |
                                   uploader.py
                               POST /api/ImputarHoras
                                        |
                              +-------- + --------+
                              |                   |
                         file_manager.py      history.py
                       (guardar en disco o   (agregar al
                        retornar bytes)      Historico)
```

---

## Referencia de modulos

| Modulo | Responsabilidad | Depende de |
|---|---|---|
| `config.py` | Variables de entorno, deteccion del modo de despliegue (`is_cloud()`), resolucion del email del colaborador (`get_email_colaborador()`) | `python-dotenv`, `os` |
| `modules/loader.py` | Parsear archivos Excel y texto TSV al DataFrame interno. Define `REGISTRO_COLUMNS` — el esquema canonico usado por todos los modulos. Construye metricas de resumen semanal. | `pandas`, `openpyxl` |
| `modules/validator.py` | Aplicar reglas de negocio fila a fila y de forma global. Gestiona festivos colombianos para la deteccion de semanas cerradas. Nunca llama a la API. | `pandas` |
| `modules/uploader.py` | Enviar cada fila en estado `Pendiente` a `POST /api/ImputarHoras`. Ordena laborales antes que adicionales. Maneja timeout, errores de conexion, 400/401/5xx. | `requests`, `config.py` |
| `modules/file_manager.py` | Abrir y guardar el workbook Excel. En modo cloud fuerza serializacion a bytes en lugar de escribir en disco. | `openpyxl`, `config.py` |
| `modules/history.py` | Agregar las filas enviadas exitosamente a la hoja `Historico`. Generar bytes de descarga en modo no local. | `openpyxl`, `pandas` |
| `modules/outlook_parser.py` | Parsear archivos `.ics` al DataFrame interno. Define `_process_event_list` (compartida con `graph_client.py`). Maneja deteccion de etiqueta `[BISOPI]`, parseo de clave-valor en descripciones, redondeo a 15 minutos, eventos de dia completo y deteccion de solapamiento. Exporta a la plantilla Excel con celdas resaltadas. | `icalendar`, `openpyxl`, `pandas`, `loader.py` |
| `modules/graph_client.py` | Integracion con Microsoft Graph API. Autentica via MSAL (solo `PublicClientApplication` — sin client secret). Implementa polling no bloqueante de device code (`poll_device_flow`). Obtiene eventos del calendario con paginacion. Convierte el JSON de Graph al DataFrame interno delegando en `_process_event_list`. | `msal`, `requests`, `outlook_parser.py` |
| `main.py` | Interfaz Streamlit. Bloque de autenticacion obligatoria en cloud. Tres pestanas de fuentes. Funcion `_render_ol_results()` compartida para la seccion de resultados de Outlook. Resuelve `email_colaborador` dinamicamente segun el ambiente. | Todos los modulos anteriores |

---

## Decisiones tecnicas

**Esquema interno unico (`REGISTRO_COLUMNS`)**
Todas las fuentes normalizan al mismo DataFrame de 11 columnas antes de validar y enviar. Esto hace que `validator.py`, `uploader.py`, `history.py` y `file_manager.py` sean completamente independientes de la fuente. Agregar una nueva fuente solo requiere producir un DataFrame conforme al esquema.

**Polling no bloqueante del device code (`poll_device_flow`)**
`acquire_token_by_device_flow()` de MSAL es un loop bloqueante que consulta el endpoint de token cada 5-15 segundos hasta que el usuario se autentica o el flujo expira. En Streamlit esto congela el hilo del servidor y cierra la conexion WebSocket, perdiendo el estado de sesion. La solucion es un unico POST HTTP no bloqueante por clic de boton que retorna el token inmediatamente si esta listo, o `None` si aun esta pendiente. El usuario vuelve a hacer clic tras autenticarse. La URL del endpoint se toma de `flow["token_endpoint"]` (provisto por MSAL en `initiate_device_flow`) en lugar de construirla con `TENANT_ID`, evitando problemas de formato cuando el identificador del tenant no es un GUID.

**Funcion `_process_event_list` compartida**
La ruta ICS (`parse_ics`) y la ruta Graph API (`events_to_dataframe`) producen ambas una lista de dicts de eventos normalizados antes de convertir al DataFrame. Toda la logica de conversion — clasificacion por etiqueta BISOPI, deteccion de solapamiento, redondeo a 15 minutos, construccion y ordenacion del DataFrame — vive en `_process_event_list` dentro de `outlook_parser.py`, importada por `graph_client.py`. No se duplica logica.

**Solo `PublicClientApplication`**
La integracion con Azure AD usa unicamente el flujo de cliente publico (ventana interactiva y device code). No existe ningun client secret en el codigo. Es intencional: la app actua como cliente de permisos delegados en nombre del usuario autenticado, coherente con el modelo de permisos `Calendars.Read` y `User.Read`.

**`BISOPI_ENV` para el modo de despliegue**
En lugar de mantener dos bases de codigo o flags por funcion, una unica funcion `is_cloud()` en `config.py` controla todo el comportamiento exclusivo del cloud. Los tres puntos afectados son: persistencia de archivos (descarga forzada), selector del metodo de autenticacion en Outlook (solo device code en cloud sin URI de redireccion configurada por IT) y el bloque de login obligatorio.

---

## Como agregar una nueva fuente de datos

1. Crear `modules/<nombre_fuente>.py`. La funcion publica del modulo debe retornar `pd.DataFrame` con exactamente las columnas de `REGISTRO_COLUMNS` (definidas en `loader.py`). Las filas que no puedan clasificarse deben tener `Estado` igual a `"⚠ Sin clasificar"` o `"⚠ Incompleto"`. Las filas listas para enviar deben tener `Estado` igual a `"Pendiente"`.

2. Agregar una nueva pestana en `main.py` siguiendo el patron existente:
   - Inicializar claves de session state con un prefijo especifico de la fuente (ej. `devops__`).
   - Llamar a `validate()` sobre el DataFrame antes de mostrarlo.
   - Renderizar la tabla con `st.data_editor`.
   - Llamar a `upload(df, email_colaborador, BISOPI_API_TOKEN, on_progress)`.
   - Persistir el resultado con `save_plantilla` y `append_historico`.

3. No se requieren cambios en `validator.py`, `uploader.py`, `history.py` ni `file_manager.py`.

---

## API — BISOPI ImputarHoras

| Parametro | Valor |
|---|---|
| URL base | `https://bisopi-open-hkdhcucjeuadafhc.eastus-01.azurewebsites.net` |
| Endpoint | `POST /api/ImputarHoras` |
| Autenticacion | Header `X-API-Token: <token>` |
| Content-Type | `application/json` |
| Exito | `200 OK` |
| Error de validacion | `400 Bad Request` con mensaje en el cuerpo |
| Error de autenticacion | `401 Unauthorized` — lanza `TokenError`, detiene el batch |
| Error del servidor | `500 Internal Server Error` |

**Payload — Laboral:**
```json
{
  "proyecto": "NOMBRE-EXACTO",
  "grupoTarea": "Nombre del grupo",
  "tarea": "Nombre exacto de la tarea",
  "tipoHora": "Laboral",
  "fechaRegistro": "YYYY-MM-DD",
  "tiempoEjecutado": { "horas": 8, "minutos": 0 },
  "emailColaborador": "usuario@bisionconsulting.com"
}
```

**Payload — Adicional** (agrega `horaInicio`, `horaFin`, `comentario`):
```json
{
  "tipoHora": "Adicional",
  "horaInicio": "18:00",
  "horaFin": "19:30",
  "comentario": "Soporte urgente"
}
```

`horaFin` es calculado por la app: `horaInicio + tiempoEjecutado`. La API no acepta este campo en blanco.

---

## Hallazgos conocidos de la API

| Hallazgo | Comportamiento | Manejo en la app |
|---|---|---|
| Proyecto inexistente devuelve 500 | Cuando `proyecto` no existe en BISOPI la API responde `500` en lugar de un `400` descriptivo. En revision por el equipo de BISOPI. | `uploader.py` captura todos los codigos que no sean 200/400/401 como error y muestra el codigo HTTP. La fila queda marcada en rojo. |
| Semana cerrada no validada por la API | La API acepta registros de semanas ya cerradas (responde `200 OK` cuando deberia rechazarlos). Bug confirmado, pendiente de correccion. | `validator.py` calcula el cierre localmente (el lunes siguiente a la semana a las 12:30, desplazado al martes si es festivo colombiano) y agrega una advertencia en amarillo. La fila no se bloquea — si la API corrige esto en el futuro, la advertencia resulta inocua. |
| Limite de 40h semanales no validado por la API | La API permite registrar mas de 40 horas laborales por semana sin error. Bug confirmado, pendiente de correccion. | `validator.py` acumula los minutos laborales por semana ISO en el batch y marca como error las filas que superen 2400 minutos antes de hacer cualquier llamada a la API. |

---

## Escalabilidad futura

| Item | Notas |
|---|---|
| Integracion con Azure DevOps | Planeada como V2. La pestana ya esta reservada en la interfaz (`tab_devops`). Requeriria un token de acceso personal de DevOps y una definicion del mapeo de campos de work items a los campos de BISOPI. La implementacion sigue el mismo patron de modulo de fuente descrito arriba. |
| Catalogo automatico desde BISOPI | La hoja `Catalogos` de la plantilla Excel se mantiene manualmente. Una version futura podria obtener los nombres validos de proyectos y tareas directamente desde un endpoint de catalogo de BISOPI (si se expone) y poblar desplegables en la interfaz. |
| Registro de auditoria multiusuario | En el despliegue cloud actual cada sesion es independiente y sin estado. Un log compartido (por ejemplo en una base de datos cloud o en un Excel compartido via SharePoint) requeriria una capa de almacenamiento persistente que Streamlit Cloud no provee de forma nativa. |
| URI de redireccion para login interactivo | El metodo de login con ventana de navegador esta implementado pero requiere que IT registre la URI de redireccion de la app en Azure AD. No se necesitan cambios en el codigo una vez configurado — el selector en la pantalla de login ya expone ambos metodos. |

---

*BISOPI Automator v1.0 — Mayo 2026 | Edson David Leon Araujo | Bision Consulting*
