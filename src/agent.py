"""Excel/Sheets agent built on Agno with OpenAI gpt-4o-mini."""

from __future__ import annotations

import os

from agno.agent import Agent
from agno.models.openai import OpenAIChat
from agno.tools.google.drive import GoogleDriveTools
from agno.tools.google.sheets import GoogleSheetsTools
from dotenv import load_dotenv

from src.excel_tools import SCOPES, ExcelTools, load_credentials
from src.i18n import DEFAULT_LANG, get_locale
from src.local_tools import LocalExcelTools
from src.memory import (
    get_db,
    get_memory_manager,
    get_user_id,
    load_or_create_session_id,
)

load_dotenv()


def build_agent(lang: str = DEFAULT_LANG) -> Agent:
    """Construct the agent.

    Args:
        lang: A language code defined in src.i18n.LOCALES. Falls back to
            DEFAULT_LANG when unknown.

    The agent is wired with:
      - Persistent SQLite storage (sessions + user memories).
      - User memory extraction after each run (local Ollama model).
      - Conversation history fed back into the context window.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY is not set in .env. "
            "Get a key at https://platform.openai.com/api-keys"
        )

    instructions = get_locale(lang)["instructions"]

    # One-shot OAuth; the resulting credentials are shared by all toolkits.
    creds = load_credentials()

    drive_tools = GoogleDriveTools(
        creds=creds,
        scopes=SCOPES,
        list_files=True,
        search_files=True,
        read_file=True,
        upload_file=False,   # .xlsx upload is owned by ExcelTools.upload_excel
        download_file=False, # .xlsx download is owned by ExcelTools.download_excel
    )

    sheets_tools = GoogleSheetsTools(
        creds=creds,
        scopes=SCOPES,
        read_sheet=True,
        create_sheet=True,
        update_sheet=True,
        create_duplicate_sheet=True,
    )

    excel_tools = ExcelTools(creds=creds)

    # Local-disk workspace; inert until the user calls open_local_folder().
    local_tools = LocalExcelTools()

    # Persistent memory wiring. The memory manager runs against a local
    # Ollama model — no api_key is needed for that.
    db = get_db()
    memory_manager = get_memory_manager()
    user_id = get_user_id()
    session_id = load_or_create_session_id()

    return Agent(
        name="ExcelAgent",
        model=OpenAIChat(id="gpt-4o-mini", api_key=api_key),
        tools=[drive_tools, sheets_tools, excel_tools, local_tools],
        instructions=instructions,
        markdown=True,
        # Hard cap on tool calls per run — prevents runaway loops where the
        # model retries unsupported operations by combining random tools.
        tool_call_limit=8,
        # Persistence
        db=db,
        user_id=user_id,
        session_id=session_id,
        # Long-term user memory extraction
        memory_manager=memory_manager,
        enable_user_memories=True,
        # Short-term: feed the recent run history back into the context
        add_history_to_context=True,
        num_history_runs=10,
        # Optional: agent can read past sessions if it needs deeper recall
        search_session_history=True,
        num_history_sessions=3,
    )
