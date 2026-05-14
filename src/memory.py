"""Persistent storage helpers — SQLite-backed sessions and user memories.

What this gives the agent:
1. Conversation history that survives terminal restarts (per session).
2. Long-term user memories (preferences, recurring files/folders) extracted
   by a `MemoryManager` after each run.
3. A simple "rolling" session model: there is one current session per user;
   `start_new_session()` rotates it (the old session stays in the DB and can
   still be referenced).

All state lives under `agent_data/` in the project root:
  - `agent.db`           : SQLite database (sessions, memories, metrics, etc.)
  - `current_session.txt`: ID of the active session
"""

from __future__ import annotations

import os
import uuid
from pathlib import Path
from typing import Optional

from agno.db.sqlite import SqliteDb
from agno.memory.manager import MemoryManager
from agno.models.ollama import Ollama

# Single-user local CLI — every conversation belongs to this user. Override
# with the EXCEL_AGENT_USER_ID environment variable if desired.
DEFAULT_USER_ID = "default_user"

# Local Ollama model used only for memory extraction. Memory extraction is a
# simple "distill durable facts from this conversation" task that small local
# models handle well — no API quota, no cloud cost, no leaked Excel data.
# Requirement: `ollama pull qwen2.5:3b-instruct` (~2 GB) and Ollama running.
MEMORY_MODEL_ID = "qwen2.5:3b-instruct"

DATA_DIR = Path("agent_data")
DB_PATH = DATA_DIR / "agent.db"
SESSION_FILE = DATA_DIR / "current_session.txt"


def _ensure_data_dir() -> None:
    DATA_DIR.mkdir(exist_ok=True)


def get_db() -> SqliteDb:
    """Return a SqliteDb pointing at the project-local DB file."""
    _ensure_data_dir()
    return SqliteDb(db_file=str(DB_PATH))


def get_user_id() -> str:
    """Resolve the user id. Override via EXCEL_AGENT_USER_ID env var."""
    return os.getenv("EXCEL_AGENT_USER_ID") or DEFAULT_USER_ID


def load_or_create_session_id() -> str:
    """Return the current rolling session id, creating one if needed."""
    _ensure_data_dir()
    if SESSION_FILE.exists():
        sid = SESSION_FILE.read_text(encoding="utf-8").strip()
        if sid:
            return sid
    return start_new_session()


def start_new_session() -> str:
    """Rotate the rolling session id and persist it. Returns the new id."""
    _ensure_data_dir()
    new_id = uuid.uuid4().hex
    SESSION_FILE.write_text(new_id, encoding="utf-8")
    return new_id


def get_memory_manager() -> MemoryManager:
    """Build a MemoryManager backed by a local Ollama model.

    The MemoryManager decides what to remember from each conversation turn
    and can update/delete prior memories. Storage is handled by the agent's
    db when wired together. Running this locally keeps memory extraction
    free and prevents sensitive Excel/Drive context from leaving the
    machine.
    """
    return MemoryManager(
        model=Ollama(id=MEMORY_MODEL_ID),
        add_memories=True,
        update_memories=True,
        delete_memories=False,
        memory_capture_instructions=(
            "Capture durable facts about the user that will help future "
            "Excel/Sheets work: preferred language, frequently mentioned "
            "files or folders (Drive and local), workflow habits, naming "
            "conventions, and project context. Skip transient details that "
            "won't matter next session."
        ),
    )
