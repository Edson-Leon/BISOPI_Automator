# BISOPI Automator — Guía para Claude Code

Herramienta interna de Bision Consulting que automatiza el registro semanal de horas
en la plataforma BISOPI. Reemplaza la carga manual tarea por tarea por un proceso de
una sola acción: se preparan los registros desde una fuente de datos, se validan contra
las reglas de negocio y se envían a la API de BISOPI en un solo paso.

Stack: Streamlit + pandas + openpyxl + requests + MSAL. Sin backend propio — toda la
lógica corre en el proceso Python que sirve la interfaz.

---

## Restricciones — no violar

Estas reglas se han mantenido a lo largo de todo el desarrollo. Respetarlas siempre:

1. **No tocar el módulo de Archivo plano** ni ninguna de sus claves de `st.session_state`.
   Es el módulo original y estable; los módulos nuevos se agregan al lado, no encima.

2. **`graph_client.py` solo hace autenticación y obtención de eventos.** Todo el
   procesamiento posterior lo hacen los módulos existentes. No duplicar lógica ahí.

3. **Prefijos obligatorios en `session_state`:** `outlook__` para el módulo de Agenda
   Outlook, `cloud__` para el gate de autenticación cloud. Una fuente de datos nueva
   usa su propio prefijo (ej. `devops__`).

4. **Solo `PublicClientApplication`.** Nunca `ConfidentialClientApplication`, nunca
   `client_secret`. La app usa permisos delegados en nombre del usuario autenticado.

5. **`.env` nunca se versiona.** Está en `.gitignore` y ahí se queda.

---

## Arquitectura

Cada fuente de datos normaliza su entrada al esquema interno `REGISTRO_COLUMNS`
(11 columnas, definido en `modules/loader.py`). A partir de ahí, todas las fuentes
comparten exactamente la misma lógica de validación, envío y persistencia.

```
Excel / copy-paste ──> loader.py ─────────┐
Archivo ICS ────────> outlook_parser.py ──┤
Microsoft Graph ────> graph_client.py ────┤
                                          v
                            DataFrame REGISTRO_COLUMNS
                                          |
                                    validator.py  (reglas de negocio)
                                          |
                                     main.py UI   (tabla editable)
                                          |
                                    uploader.py   (POST /api/ImputarHoras)
                                          |
                            file_manager.py  +  history.py
```

| Módulo | Responsabilidad |
|---|---|
| `config.py` | Variables de entorno, `is_cloud()`, `get_email_colaborador()`, `has_graph_access()` |
| `modules/loader.py` | Excel y TSV → DataFrame. **Define `REGISTRO_COLUMNS`** — esquema canónico |
| `modules/validator.py` | Reglas de negocio fila a fila y globales. Festivos colombianos. Nunca llama a la API |
| `modules/uploader.py` | Envía filas `Pendiente` a la API. Laborales antes que adicionales |
| `modules/file_manager.py` | Guarda el workbook. En cloud fuerza serialización a bytes |
| `modules/history.py` | Agrega filas exitosas a la hoja `Historico` |
| `modules/outlook_parser.py` | Parseo ICS, `_process_event_list` (compartida), `recalculate_estados`, exportación a plantilla |
| `modules/graph_client.py` | MSAL + Graph API. Delega el procesamiento en `outlook_parser` |
| `main.py` | UI Streamlit. Gate de auth cloud, tres pestañas, `_render_ol_results()` |

---

## Ambientes

`BISOPI_ENV` controla el modo. Una sola función `is_cloud()` en `config.py` gobierna
todo el comportamiento exclusivo del cloud — no hay dos bases de código.

| | Local | Cloud |
|---|---|---|
| Email del colaborador | `EMAIL_COLABORADOR` del `.env` | Del login con Microsoft (`cloud__user_email`) |
| Plantilla Excel | Lectura/escritura directa vía `PLANTILLA_PATH` | Siempre descarga; fallback a `data/Plantilla_BISOPI_Automator.xlsx` |
| Login Microsoft | Ventana interactiva + device code | Ambos, pero el interactivo requiere redirect URI registrada por IT |
| Gate de autenticación | No aplica | Obligatorio antes de renderizar cualquier pestaña |

Variables del `.env`: `EMAIL_COLABORADOR`, `BISOPI_API_TOKEN`, `PLANTILLA_PATH` (opcional),
`AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `BISOPI_ENV`. En cloud van en Streamlit Cloud →
Settings → Secrets, en formato TOML.

---

## Trampas conocidas — ya resueltas, no reintroducir

**Device code flow debe ser no bloqueante.** `acquire_token_by_device_flow()` de MSAL
hace polling bloqueante que congela el hilo del servidor de Streamlit y cierra el
WebSocket, perdiendo el estado de sesión. La solución es `poll_device_flow()`: un único
POST HTTP por clic de botón que retorna el token o `None` si sigue pendiente.

**La URL del token sale de `flow["token_endpoint"]`**, no se construye con `TENANT_ID`.
Si el tenant no es un GUID (puede ser una URL STS completa o un dominio), construirla a
mano produce una URL malformada y Azure responde `authorization_pending` para siempre.

**`data/` está en `.gitignore` con excepción.** La regla es `data/*` más
`!data/Plantilla_BISOPI_Automator.xlsx`. La plantilla base debe estar versionada porque
en cloud es el fallback de `export_to_template`. Las plantillas con datos reales del
usuario no se suben.

**`datetime.fromisoformat()` en Python ≤ 3.10** solo acepta hasta 6 decimales; Graph API
devuelve 7. Hay truncamiento en el parseo.

**`CONCAT_WS` / claves compuestas:** un componente nulo anula la clave completa.
Aplica al armar identificadores compuestos.

---

## Estado de las fuentes de datos

| Fuente | Estado |
|---|---|
| Plantilla Excel (archivo o copy-paste) | Cerrada |
| Agenda Outlook — archivo ICS | Cerrada |
| Agenda Outlook — Microsoft Graph API | Cerrada |
| Azure DevOps | En desarrollo (V2). Pestaña `tab_devops` ya reservada en la UI |

Despliegue cloud activo en `https://bisopi-automator.streamlit.app` (privado, por invitación).
Repositorio: `https://github.com/Edson-Leon/BISOPI_Automator` (privado).

---

## Cómo agregar una fuente de datos nueva

1. Crear `modules/<fuente>.py`. La función pública retorna un `pd.DataFrame` con
   exactamente las columnas de `REGISTRO_COLUMNS`. Filas sin clasificar →
   `Estado = "⚠ Sin clasificar"` o `"⚠ Incompleto"`; listas → `"Pendiente"`.
2. Agregar la pestaña en `main.py` siguiendo el patrón existente: prefijo propio en
   `session_state`, `validate()` antes de mostrar, `st.data_editor`, luego
   `upload(df, email_colaborador, BISOPI_API_TOKEN, on_progress)` y persistencia con
   `save_plantilla` + `append_historico`.
3. No se requieren cambios en `validator.py`, `uploader.py`, `history.py` ni `file_manager.py`.

---

## Comandos

```bash
# Ejecutar (local)
python -m streamlit run main.py

# Instalar dependencias
python -m pip install -r requirements.txt
```

En Windows: `install.bat` y `run.bat`. `run.bat` no debe abrir el navegador manualmente
— Streamlit ya lo hace, y hacerlo dos veces abre dos pestañas.

---

## Documentación

- `README.md` — portada del repositorio, instalación y configuración
- `docs/ARCHITECTURE.md` — referencia técnica: decisiones, API, escalabilidad
- `docs/DISTRIBUTION.md` — modalidades de distribución cloud vs local

---

*BISOPI Automator v1.0 — Edson David Leon Araujo | Bision Consulting*
