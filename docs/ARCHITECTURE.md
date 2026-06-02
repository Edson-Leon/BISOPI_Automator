# BISOPI Automator — Architecture Reference

## Overview

BISOPI Automator is a single-process Streamlit application with no custom backend. All logic runs in the Python process that serves the UI. The architecture follows a strict separation between data sources (how records enter the system), the internal processing pipeline (validation, upload, history), and the UI layer (main.py). Each data source is an independent module that normalizes its input into a shared internal DataFrame schema (`REGISTRO_COLUMNS`), after which all sources share the exact same validation, upload, and persistence logic. The application supports two deployment modes — local and Streamlit Cloud — controlled by the `BISOPI_ENV` environment variable, which gates cloud-only behaviors (mandatory Microsoft login, forced download mode for files) without duplicating any business logic.

---

## Flow Diagram

```
User input
    |
    +-- Excel file / copy-paste ──> loader.py ──────────+
    |                                                    |
    +-- ICS file ──────────────> outlook_parser.py ─────+
    |                                                    |
    +-- Microsoft Graph API ───> graph_client.py ────── +
                                        |                |
                                        v                v
                                 REGISTRO_COLUMNS DataFrame
                                        |
                                   validator.py
                                  (business rules)
                                        |
                                   main.py UI
                                  (editable table)
                                        |
                                   uploader.py
                               POST /api/ImputarHoras
                                        |
                              +-------- + --------+
                              |                   |
                         file_manager.py      history.py
                       (save to disk or     (append to
                        return bytes)       Historico sheet)
```

---

## Module Reference

| Module | Responsibility | Depends on |
|---|---|---|
| `config.py` | Environment variables, deployment mode detection (`is_cloud()`), email resolution (`get_email_colaborador()`) | `python-dotenv`, `os` |
| `modules/loader.py` | Parse Excel files and TSV copy-paste into the internal DataFrame. Defines `REGISTRO_COLUMNS` — the canonical schema used by all modules. Builds weekly summary metrics. | `pandas`, `openpyxl` |
| `modules/validator.py` | Apply business rules row-by-row and globally. Manages Colombia public holidays for closed-week detection. Never calls the API. | `pandas` |
| `modules/uploader.py` | Send each `Pendiente` row to `POST /api/ImputarHoras`. Orders laborales before adicionales. Handles timeout, connection errors, 400/401/5xx. | `requests`, `config.py` |
| `modules/file_manager.py` | Open and save the Excel workbook. In cloud mode, forces serialization to bytes instead of writing to disk. | `openpyxl`, `config.py` |
| `modules/history.py` | Append successfully uploaded rows to the `Historico` sheet. Generate download bytes for non-local mode. | `openpyxl`, `pandas` |
| `modules/outlook_parser.py` | Parse `.ics` files into the internal DataFrame. Defines `_process_event_list` (shared with `graph_client.py`). Handles `[BISOPI]` tag detection, description key-value parsing, 15-minute rounding, all-day events, and overlap detection. Exports to the Excel template with highlighted cells. | `icalendar`, `openpyxl`, `pandas`, `loader.py` |
| `modules/graph_client.py` | Microsoft Graph API integration. Authenticates via MSAL (`PublicClientApplication` only — no client secret). Implements non-blocking device code polling (`poll_device_flow`). Fetches calendar events with pagination. Converts Graph JSON to the internal DataFrame by delegating to `_process_event_list`. | `msal`, `requests`, `outlook_parser.py` |
| `main.py` | Streamlit UI. Cloud authentication gate. Three source tabs. Shared `_render_ol_results()` for the Outlook results section. Routes `email_colaborador` dynamically based on environment. | All modules above |

---

## Technical Decisions

**Single internal schema (`REGISTRO_COLUMNS`)**
All data sources normalize to the same 11-column DataFrame before validation and upload. This means validator.py, uploader.py, history.py, and file_manager.py are completely source-agnostic. Adding a new source only requires producing a conforming DataFrame.

**Non-blocking device code polling (`poll_device_flow`)**
MSAL's `acquire_token_by_device_flow()` is a blocking loop that polls every 5-15 seconds until the user authenticates or the flow expires. In Streamlit, this freezes the server thread and drops the WebSocket connection, losing session state. The solution is a single non-blocking HTTP POST per button click that returns the token immediately if ready, or `None` if still pending. The user clicks the button again after authenticating. The token endpoint URL is taken from `flow["token_endpoint"]` (provided by MSAL during `initiate_device_flow`) rather than constructed from `TENANT_ID`, which avoids URL formatting issues with non-GUID tenant identifiers.

**`_process_event_list` shared function**
The ICS path (`parse_ics`) and the Graph API path (`events_to_dataframe`) both produce a list of normalized event dicts before converting to a DataFrame. The entire conversion logic — BISOPI tag classification, overlap detection, 15-minute rounding, DataFrame construction and sorting — lives in `_process_event_list` in `outlook_parser.py`, imported by `graph_client.py`. No logic is duplicated.

**`PublicClientApplication` only**
The Azure AD integration uses only the public client flow (interactive window and device code). There is no client secret anywhere in the codebase. This is intentional: the app acts as a delegated-permission client on behalf of the signed-in user, consistent with the `Calendars.Read` and `User.Read` permission model.

**`BISOPI_ENV` for deployment mode**
Rather than maintaining two codebases or using feature flags per function, a single `is_cloud()` function in `config.py` gates all cloud-specific behavior. The three affected points are: file persistence (forced download), Outlook auth method selector (device code only shown in cloud without IT configuring a redirect URI), and the mandatory login gate.

---

## Adding a New Data Source

1. Create `modules/<source_name>.py`. The module's public function must return `pd.DataFrame` with exactly the columns in `REGISTRO_COLUMNS` (defined in `loader.py`). Rows that cannot be classified must set `Estado` to `"⚠ Sin clasificar"` or `"⚠ Incompleto"`. Rows ready to send must set `Estado` to `"Pendiente"`.

2. Add a new tab in `main.py` following the existing pattern:
   - Initialize session state keys with a source-specific prefix (e.g. `devops__`).
   - Call `validate()` on the DataFrame before displaying.
   - Render the table with `st.data_editor`.
   - Call `upload(df, email_colaborador, BISOPI_API_TOKEN, on_progress)`.
   - Persist the result with `save_plantilla` and `append_historico`.

3. No changes are required in `validator.py`, `uploader.py`, `history.py`, or `file_manager.py`.

---

## API — BISOPI ImputarHoras

| Parameter | Value |
|---|---|
| Base URL | `https://bisopi-open-hkdhcucjeuadafhc.eastus-01.azurewebsites.net` |
| Endpoint | `POST /api/ImputarHoras` |
| Authentication | Header `X-API-Token: <token>` |
| Content-Type | `application/json` |
| Success | `200 OK` |
| Validation error | `400 Bad Request` with message in body |
| Auth error | `401 Unauthorized` — raises `TokenError`, stops the batch |
| Server error | `500 Internal Server Error` |

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

**Payload — Adicional** (adds `horaInicio`, `horaFin`, `comentario`):
```json
{
  "tipoHora": "Adicional",
  "horaInicio": "18:00",
  "horaFin": "19:30",
  "comentario": "Soporte urgente"
}
```

`horaFin` is computed by the app: `horaInicio + tiempoEjecutado`. The API does not accept it as blank.

---

## Known API Issues

| Issue | Behavior | App workaround |
|---|---|---|
| Non-existent project returns 500 | When `proyecto` does not exist in BISOPI the API responds `500` instead of a descriptive `400`. Under review by the BISOPI team. | `uploader.py` catches all non-200/400/401 codes as errors and shows the HTTP status code. The row is marked red. |
| Closed-week not validated by API | The API accepts records for weeks already closed (returns `200 OK` when it should reject them). Confirmed bug, pending correction. | `validator.py` computes the closure deadline locally (Monday after the week at 12:30, shifted to Tuesday on Colombia public holidays) and adds a yellow warning. The row is not blocked — if the API eventually fixes this, the warning remains harmless. |
| 40-hour weekly limit not validated by API | The API allows more than 40 labor hours per week without error. Confirmed bug, pending correction. | `validator.py` accumulates labor minutes per ISO week across the batch and marks rows that exceed 2400 minutes as errors before any API call is made. |

---

## Future Scalability

| Item | Notes |
|---|---|
| Azure DevOps integration | Planned as V2. Tab already reserved in the UI (`tab_devops`). Would require a DevOps personal access token and a mapping definition from work item fields to BISOPI fields. Implementation follows the same source module pattern described above. |
| Automatic catalog from BISOPI | The `Catalogos` sheet in the Excel template is currently maintained manually. A future version could fetch valid project/task names directly from a BISOPI catalog endpoint (if exposed) and populate dropdowns in the UI. |
| Multi-user audit trail | In the current cloud deployment, each session is independent and stateless. A shared audit log (e.g. stored in a cloud database or appended to a shared Excel via SharePoint) would require a persistent storage layer, which Streamlit Cloud does not provide natively. |
| Redirect URI for interactive login | The interactive browser login method is implemented but requires IT to register the app's redirect URI in Azure AD. No code changes are needed once this is configured — the radio button in the login screen already exposes both methods. |

---

*BISOPI Automator v1.0 — Mayo 2026 | Edson David Leon Araujo | Bision Consulting*
