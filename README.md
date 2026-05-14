# Excel Agent

A terminal assistant that manages Google Sheets and Excel (.xlsx) files
on Google Drive through a Turkish- or English-speaking chat. It is
built on the Agno framework, powered by OpenAI `gpt-4o-mini`, and
talks to Google's official Sheets/Drive APIs. It can read and edit
both spreadsheets that live on Drive and .xlsx files on the local
disk.

---

## Key Features

- Google Drive integration: file search, listing, opening.
- Reads and writes native Google Sheets files directly through the
  Sheets API.
- Edits Excel (.xlsx) files **while preserving the format**.
  .xlsx files on Drive are never converted to Google Sheets; they are
  downloaded, edited locally with openpyxl, and uploaded back to the
  same Drive ID. The file remains .xlsx; sharing settings and links
  are preserved.
- Bulk operations across a Drive folder
  (`bulk_find_replace_in_folder`, with recursive support).
- Opt-in local folder mode: the agent cannot touch any file on the
  local disk unless the user explicitly opens a folder. Once
  `open_local_folder` is called, the agent can list, read and edit
  files inside that folder. The project ships with a suggested
  `workbooks/` drop folder; `open_local_folder()` with no argument
  opens it. Any other path can be passed explicitly, including `.`
  for the whole project root.
- Bulk operations across the active local folder
  (`bulk_find_replace_in_local_folder`).
- High-level Excel tools: find/replace, conditional row delete,
  filtering, column statistics, formula application, row append,
  new sheet creation, row range insert/delete.
- Multilingual UI: a TR / EN picker at startup; all strings are pulled
  from a single `i18n` module so adding a new language is a one-file
  change.
- Persistent chat memory: conversation history and long-term user
  memories are stored in a local SQLite database. The agent
  remembers previous turns and durable facts (preferred language,
  frequently mentioned files/folders, workflow habits) across
  terminal restarts. A small local model running on **Ollama**
  (`qwen2.5:3b-instruct` by default) performs memory extraction in the
  background — no extra API cost, no Excel/Drive context leaves the
  machine. The main agent talks to OpenAI `gpt-4o-mini` for the
  tool-using work.
- Rolling sessions with manual rotation: one active session is kept
  by default; type `:new` in the terminal to start a fresh session.
  Old sessions stay in the DB and can be surfaced via session search.
- Path-traversal protection: in local folder mode the agent cannot
  escape the opened folder.
- Secrets stay out of git: `.env`, `credentials.json`, `token.json`,
  service account files and the local DB (`agent_data/`) are excluded
  via `.gitignore`.

---

## File Layout

```
F:\excel agent\
├── main.py                # Terminal chat entry point
├── pyproject.toml         # uv project definition and dependencies
├── uv.lock                # uv lockfile
├── .env.example           # Environment variable template
├── .gitignore
├── README.md
├── credentials.json       # (you provide — Drive OAuth client)
├── token.json             # (auto-generated after the first auth)
├── workbooks/             # suggested drop folder for local .xlsx files
├── excel_workdir/         # (runtime) downloaded .xlsx files
├── agent_data/            # (runtime) SQLite DB + active session id
└── src/
    ├── __init__.py
    ├── agent.py           # Agno agent setup + OpenAI gpt-4o-mini
    ├── excel_tools.py     # Drive .xlsx tools + folder tools
    ├── local_tools.py     # Local folder tools (opt-in)
    ├── memory.py          # SQLite DB + session/user memory wiring
    └── i18n.py            # All TR/EN strings in one place
```

---

## Architecture and Toolkits

The agent operates with four toolkits. It picks the correct one for
each task by itself.

### 1) `GoogleDriveTools` (official Agno)

Finds and lists files on Drive.

- `list_files` — returns recent files.
- `search_files` — Drive query-syntax search.
- `read_file` — text-based content read.

### 2) `GoogleSheetsTools` (official Agno)

For native Google Sheets files.

- `read_sheet` — reads the given range.
- `update_sheet` — writes data.
- `create_sheet` — creates a new Sheets file.
- `create_duplicate_sheet` — copies an existing Sheets file.

### 3) `ExcelTools` (custom — `src/excel_tools.py`)

For .xlsx files on Drive and Drive folders.

Drive transfer:

- `download_excel(file_id)` — downloads the `.xlsx` into
  `excel_workdir/`.
- `upload_excel(file_id)` — uploads back to the same Drive ID;
  format is preserved.

Local inspection (on the downloaded file):

- `list_sheet_names`
- `read_excel_range`
- `find_cells_excel`
- `filter_rows_excel`
- `column_summary_excel`

Single-cell / block editing:

- `update_excel_cell`
- `update_excel_range`
- `append_excel_row`
- `create_excel_sheet_tab`
- `set_formula_excel` — applies a formula with a `{row}` placeholder.

Single-file bulk operations:

- `find_and_replace_excel` — case_sensitive, whole_cell, use_regex,
  sheet_name and column filters.
- `delete_rows_excel`
- `delete_rows_where` — conditional delete
  (==, !=, >, <, >=, <=, contains, not_contains, empty, not_empty).
- `insert_rows_excel`

Column operations and sort:

- `delete_excel_columns`
- `insert_excel_columns`
- `sort_excel_by_column`

Sheet tab management:

- `rename_excel_sheet_tab`
- `delete_excel_sheet_tab`

Analytical, export and creation:

- `describe_excel` — pandas-style per-column summary.
- `export_drive_excel_to_csv(file_id, sheet, output_path)` — saves a
  sheet to a local CSV (does not push back to Drive).
- `create_drive_xlsx_file(name, sheet_name, target_folder_id)` —
  uploads a brand-new empty .xlsx to Drive and caches it locally for
  immediate editing.

Drive-wide file management:

- `move_drive_file_to_trash(file_id)` — recoverable from Drive Trash.
- `rename_drive_file(file_id, new_name)`
- `copy_drive_file(file_id, new_name, target_folder_id)`

Folder level (Drive):

- `find_folder_by_name(name)`
- `list_excels_in_folder(folder_id, include_sheets, recursive, max_results)`
- `bulk_find_replace_in_folder(folder_id, find, replace, ..., recursive)`
  — runs the download → find/replace → upload loop on behalf of the
  agent. Files with zero replacements are not re-uploaded.

### 4) `LocalExcelTools` (custom — `src/local_tools.py`)

For .xlsx files on the local disk. **Stateful**: no tool runs until a
folder is opened.

Workspace lifecycle:

- `open_local_folder(path)` — activates the folder and returns the
  initial listing.
- `close_local_folder()` — deactivates the workspace.
- `show_local_folder()` — shows the active folder.

File system:

- `list_local_folder(subpath, recursive)` — the `ls`. With
  `recursive=True` it walks subdirectories and returns entries as
  relative paths.

Excel inspection and editing — local counterparts of every
`ExcelTools` operation:

- `list_local_sheet_names`
- `read_local_excel`
- `find_cells_local_excel`, `filter_local_excel_rows`,
  `column_summary_local_excel`
- `update_local_excel_cell`, `update_local_excel_range`,
  `append_local_excel_row`, `create_local_excel_sheet_tab`,
  `set_formula_local_excel`
- `find_and_replace_local_excel`, `delete_local_excel_rows`,
  `delete_local_excel_rows_where`, `insert_local_excel_rows`

Column operations and sort:

- `delete_local_excel_columns`, `insert_local_excel_columns`
- `sort_local_excel_by_column`

Sheet tab management:

- `rename_local_excel_sheet_tab`, `delete_local_excel_sheet_tab`

File management:

- `move_local_excel_file_to_trash` — sends to the OS Recycle Bin.
- `rename_local_excel_file`, `copy_local_excel_file`,
  `move_local_excel_file` (subdirs allowed inside the workspace).

Analytical and export:

- `describe_local_excel` — pandas-style per-column summary.
- `search_in_all_local_files` — pattern hit-list across all workspace
  .xlsx files (optionally recursive).
- `export_local_excel_to_csv` — write a sheet as CSV inside the workspace.

Folder-wide bulk operation:

- `bulk_find_replace_in_local_folder(find, replace, ..., recursive)`

Local files are saved **in place**. There is no download/upload step.

---

## Memory and Sessions

State lives under `agent_data/` in the project root and is wired up in
`src/memory.py`:

```
agent_data/
├── agent.db                # SQLite database (sessions + memories)
└── current_session.txt     # ID of the active rolling session
```

### What is persisted

- **Conversation history** — every user/agent turn is stored in
  `agent.db` along with the session ID. When the agent rebuilds it
  resumes the same session, so closing and reopening the terminal does
  not wipe context.
- **User memories** — after each run a `MemoryManager` extracts durable
  facts about the user from the conversation: language preference,
  frequently used files and folders (Drive and local), workflow habits,
  naming conventions. Memory extraction runs on a local **Ollama** model
  (`qwen2.5:3b-instruct` by default), so it costs nothing and your data does
  not leave the machine. The main agent stays on `gpt-4o-mini`.

  Prerequisite: install [Ollama](https://ollama.com), then pull the
  model once:

  ```powershell
  ollama pull qwen2.5:3b-instruct
  ```

  The Ollama daemon must be running (it normally starts automatically
  after install).

- **Past sessions** — older sessions are kept; the agent can search the
  three most recent sessions (`search_session_history=True`,
  `num_history_sessions=3`) when context is needed.

### Ollama setup (memory engine)

Memory extraction runs against a small local model served by
[Ollama](https://ollama.com). Once it is set up the agent uses it
transparently in the background; you do not interact with Ollama
directly.

#### 1. Install Ollama

- **Windows**: download the installer from <https://ollama.com/download>
  and run it. After install, an Ollama icon (a small llama) appears in
  the system tray; the daemon runs as long as that icon is there.
- **macOS**: download the `.dmg` from the same page and drag to
  Applications. Launch it once to start the menu-bar agent.
- **Linux**: `curl -fsSL https://ollama.com/install.sh | sh` starts
  the daemon as a systemd unit.

#### 2. Confirm the daemon is running

The daemon listens on `localhost:11434`. The simplest check:

```cmd
ollama list
```

It should print a (possibly empty) model table without errors. If the
command isn't found, see the PATH section below.

#### 3. Pull the memory model

```cmd
ollama pull qwen2.5:3b-instruct
```

This downloads about **2 GB** the first time. After that the model is
cached locally and `ollama list` will show it.

#### 4. Windows PATH troubleshooting

On a fresh install you may see:

```
'ollama' is not recognized as an internal or external command
```

This means the installer added `ollama.exe` to the user PATH, but the
shell window you have open was launched **before** the PATH change.
Fixes, in order of preference:

1. **Close every cmd / PowerShell window**, then open a new one and
   retry. The new shell picks up the updated PATH.
2. **Verify Ollama is installed**:

   ```cmd
   dir "%LOCALAPPDATA%\Programs\Ollama"
   ```

   You should see `ollama.exe` and `ollama app.exe`.

3. **Use the full path** (works without PATH):

   ```cmd
   "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" pull qwen2.5:3b-instruct
   ```

4. **Manually add to PATH** if it really is missing:
   - Win + R → `sysdm.cpl` → `Advanced` → `Environment Variables`
   - Under _User variables_, select `Path` → `Edit` → `New`
   - Add: `C:\Users\<your-user>\AppData\Local\Programs\Ollama`
   - OK out, close all shells, open a fresh one, retry.

#### 5. Verify the full chain

In a new shell:

```cmd
ollama --version
ollama list
```

`qwen2.5:3b-instruct` should appear in the list. The agent will start
using it on its next run.

#### 6. Opting out of Ollama

If you do not want to install Ollama, edit `src/agent.py` and set:

```python
enable_user_memories=False,
```

The agent will still keep conversation history within a session, but
the long-term `MemoryManager` (which is the only thing that talks to
Ollama) will be disabled.

### Identity

The CLI is single-user. The user ID defaults to `default_user` and can
be overridden with the `EXCEL_AGENT_USER_ID` environment variable.

### Session lifecycle

There is **one active session at a time**. Its ID lives in
`current_session.txt` and is reloaded automatically on every start.

Rotate to a fresh session by typing one of these in the chat:

```
:new
/new
yeni oturum
new session
```

A new session ID is generated and the agent is rebuilt with it. The
previous session stays in the database and remains searchable.

### Disabling memory

If you want a stateless setup, edit `src/agent.py` and set:

```python
enable_user_memories=False,
search_session_history=False,
```

The conversation history within a single session still works because
the agent reuses its own DB rows for the current session; remove `db`
entirely to drop all persistence.

### Wiping data

To start with a clean slate, delete the `agent_data/` directory. It
is in `.gitignore`, so it will not have been committed.

---

## Installation

### Prerequisites

- Python 3.11 or later
- The `uv` package manager (https://docs.astral.sh/uv/)
- A Google Cloud project with Sheets API and Drive API enabled
- An OpenAI API key (https://platform.openai.com/api-keys)

### Step 1 — Dependencies

The project was already set up with `uv init` and the packages were
added. After a fresh clone, sync the dependencies with:

```powershell
uv sync
```

Manual package add (only if needed):

```powershell
uv add agno google-genai google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dotenv openpyxl sqlalchemy
```

### Step 2 — OpenAI API key

1. Get a key at https://platform.openai.com/api-keys.
2. Create a `.env` file in the project root (copy `.env.example` as a
   template).
3. Put this line in it:

```
OPENAI_API_KEY=sk-...your_key
```

### Step 3 — Google Cloud OAuth

1. Open your project at https://console.cloud.google.com.
2. From `APIs & Services > Library` enable both:
   - Google Sheets API
   - Google Drive API
3. Open `OAuth consent screen`. Choose External. Add your Google
   account as a test user.
4. Go to `Credentials > Create credentials > OAuth client ID > Desktop
app` and create a new OAuth client.
5. Download the JSON file and save it as
   **`F:\excel agent\credentials.json`**.

On the first run a browser window opens automatically: pick your
Google account and grant the permissions. After that `token.json` is
created and you are not asked again.

---

## Running

In the project root:

```powershell
uv run python main.py
```

It asks for a language at startup:

```
Language / Dil [tr (Türkçe) / en (English)] (default: tr):
```

Accepted inputs: `tr`, `en`, or display names like `Türkçe`,
`English`. An empty Enter falls back to Turkish. The selection drives
both the banner and the agent's reply language.

To exit: `çık`, `cik`, `exit`, `quit`, or `:q`.

To start a new session (rotate the rolling session id): `:new`, `/new`,
`yeni oturum`, or `new session`. The previous session stays in the
local DB.

---

## Sample Dialogs

### Finding and reading a file on Drive

```
You > open the 2024 sales file
Agent > Searching Drive for "2024 sales"...
        3 matches found. Which one should I open?
        1) 2024 Sales (January).xlsx
        2) 2024 Sales (February).xlsx
        3) 2024 Sales Summary (Sheets)

You > 3
Agent > Opened 2024 Sales Summary. Which range should I read?

You > the first 10 rows of Sheet1
Agent > [shows the table]
```

### Bulk operation on a Drive .xlsx

```
You > in every excel inside the Customers folder,
      replace "old address" with "new address"
Agent > Found the "Customers" folder (ID: 1AbC...).
        It contains 12 .xlsx files. I'll run the
        download/edit/upload loop for each. Confirm?

You > yes
Agent > Starting...
        Result: 47 replacements across 9 files;
        3 files had no matches (not re-uploaded).
```

### Local folder mode

```
You > open the folder ./data
Agent > Folder opened: F:\...\data
        Contents: report.xlsx, old.xlsx, notes.txt, subdir/

You > list everything including subfolders
Agent > [recursive ls output]

You > in report.xlsx, replace "in progress" with "done"
      in the Status column
Agent > Sheet1: 12 replacements, saved in place.

You > close the local folder
Agent > Closed: F:\...\data
```

### Applying a formula

```
You > write =C{row}*D{row} from E2 to E100 in report.xlsx
Agent > Confirm? E2:E100 will receive a formula.

You > yes
Agent > Written, saved in place.
```

### Memory recall across restarts

```
# First session
You > my main reports are under the "2024 Reports" Drive folder
Agent > Got it.

# (close the terminal, reopen, run again)

You > what is my main reports folder again?
Agent > You said your main reports are under the
        "2024 Reports" Drive folder.

You > :new
Agent > [OK] Started a new session (a1b2c3d4…).
        The previous conversation stays in the DB.
```

---

## Important Behavioural Rules

The agent is instructed to honour these rules:

1. **Excel is never converted to Sheets**: a .xlsx file on Drive
   always follows the download → edit → upload-back-to-same-ID flow.
   No format conversion happens.

2. **Local folder is opt-in**: the agent does not touch the local
   disk unless the user explicitly asks for `open_local_folder`.

3. **Confirmation before writes**: explicit user confirmation is
   required before any write/delete. Risky ranges trigger a warning.

4. **No LLM-side loops for bulk work**: for requests like
   "replace X with Y" it calls the `find_and_replace_*` tool directly
   instead of pulling the whole file into the model.

5. **Asks on multiple matches**: when a file/folder search returns
   several candidates the agent asks the user which one to use.

---

## Adding a New Language

Every user-visible string lives in `src/i18n.py`. To add a language
you only need to add a new entry to the `LOCALES` dict; no other file
needs to change.

Example (German):

```python
LOCALES["de"] = {
    "name": "Deutsch",
    "banner": _DE_BANNER,
    "instructions": _DE_INSTRUCTIONS,
    "ready": "Agent bereit. Stelle deine Frage:\n",
    "user": "Du > ",
    "bye": "Tschüss!",
    "err_init": "[FEHLER] Agent-Start fehlgeschlagen: {e}",
    "err_run": "[FEHLER] {e}",
    "exit_words": {"exit", "quit", "ende", ":q"},
}
```

The language picker lists the new code automatically.

---

## Security Notes

- `.env`, `credentials.json`, `token.json` and service-account JSON
  files are in `.gitignore` — they never enter version control.
- In local folder mode the agent **cannot do path traversal**.
  Attempts to escape via `../` or absolute paths are rejected.
- `local_*` tools refuse to run until a folder is opened; the default
  posture is safe-by-default.
- When uploading a .xlsx back to Drive the same file ID is reused;
  links and sharing stay intact.
- Excel lock files (`~$filename.xlsx`) are skipped automatically
  during bulk operations.

---

## Developer Notes

### Runtime

All commands run through `uv`:

```powershell
uv run python main.py                # Start the agent
uv run python -c "from src.agent import build_agent"   # Quick smoke test
```

### Dependencies

Managed in `pyproject.toml`:

- `agno` — agent framework
- `openai` — OpenAI Python SDK (main agent)
- `google-genai` — Gemini SDK (kept for future use, currently unused)
- `google-api-python-client` — Sheets + Drive REST API
- `google-auth-httplib2`, `google-auth-oauthlib` — OAuth
- `python-dotenv` — `.env` support
- `openpyxl` — .xlsx read/write
- `sqlalchemy` — required by agno's SQLite backend (sessions + memories)
- `ollama` — Python client for the local memory-extraction model
- `openai` — pulled in transitively by agno's Ollama integration

### Configuration points

- To change the main model: `OpenAIChat(id=...)` in `src/agent.py`.
- To change the memory model: `MEMORY_MODEL_ID` constant in
  `src/memory.py` (defaults to `qwen2.5:3b-instruct` on Ollama). To use a
  different Ollama model, pull it first (`ollama pull <name>`) and
  set `MEMORY_MODEL_ID` to its tag.
- History window: `Agent(..., num_history_runs=10)`.
- Past sessions surfaced via search: `num_history_sessions=3`.
- DB location: `DB_PATH` in `src/memory.py` (defaults to
  `agent_data/agent.db`).
- User identity: `EXCEL_AGENT_USER_ID` environment variable.
- OAuth scopes: `SCOPES` list in `src/excel_tools.py`. Currently set
  to `spreadsheets` + full `drive`; the full scope is required for
  the .xlsx copy/conversion paths.

### Design decisions worth knowing

- Drive .xlsx files are updated in place via `drive.files().update`,
  not by copying. Drive ID, link and sharing all stay the same.
- Folder-wide find/replace skips upload for files with zero
  replacements (bandwidth saving).
- `find_and_replace_excel` defaults to case-insensitive even in regex
  mode; pass `case_sensitive=True` to flip it.
- Operators accept user-friendly aliases: `==`, `=`, `equals`; `!=`,
  `<>`, `not_equals`; `>`, `gt`; etc.

---

## Common Issues

**"OPENAI_API_KEY is not set in .env"**
The `.env` file must be in the project root. Copy from `.env.example`
and put the real key inside.

**The browser OAuth screen says "access blocked"**
Add your Gmail address as a test user in Google Cloud Console > OAuth
consent screen.

**`File is not .xlsx (mimeType=...)`**
`download_excel` works only on Excel binary files. For native Google
Sheets, use `GoogleSheetsTools` (`read_sheet`, `update_sheet`).

**`Path escapes the active folder`**
In local folder mode the agent cannot leave the opened folder. Open a
higher-level folder that contains the path you need.

**`No folder is open. Call open_local_folder(path) first.`**
Local tools are opt-in. Call `open_local_folder` first.

**Agent doesn't remember what we talked about last time**
Memory persistence requires `agent_data/agent.db` to be present and
writable. Verify the directory exists after the first run and that
nothing is wiping it between sessions. Use `:new` only when you
actually want to start a fresh session.

**`ModuleNotFoundError: No module named 'sqlalchemy'`**
The SQLite backend needs SQLAlchemy. Run `uv add sqlalchemy` (or
`uv sync` after pulling a recent version of `pyproject.toml`).

**Ollama connection error / "model not found"**
Memory extraction runs on a local Ollama model. Make sure Ollama is
installed and the daemon is running (`ollama list` should respond),
and that the model is pulled: `ollama pull qwen2.5:3b-instruct`.
See the "Ollama setup" subsection under "Memory and Sessions" for the
full walkthrough including Windows PATH fixes.

**`'ollama' is not recognized as an internal or external command`**
PATH was updated by the installer but your current shell still has the
old PATH. Close every cmd/PowerShell window, open a fresh one, and try
again. If that does not fix it, see the PATH troubleshooting steps in
the "Ollama setup" subsection.

---

## Roadmap (possible future additions)

- Folder-wide bulk operations for native Google Sheets (via the
  Sheets API `batchUpdate.findReplace`).
- More bulk operations on the local folder side
  (`bulk_delete_rows_where_in_local_folder`, etc.).
- Async processing (parallel execution for large folders).
- A web interface (Streamlit or FastAPI).
- A pandas / DataFrame-based advanced analysis tool.
- A `describe_sheet` tool that summarises columns/types without
  feeding the data to the LLM.

---

## License

This project is distributed under the MIT License. See the `LICENSE`
file in the project root for the full text.
