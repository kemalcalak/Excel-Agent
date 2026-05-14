"""Terminal chat entry point for the Excel Agent."""

from __future__ import annotations

import sys

from src.agent import build_agent
from src.i18n import DEFAULT_LANG, get_locale, resolve_lang, supported_languages
from src.memory import start_new_session


def pick_language() -> str:
    """Show a dynamic picker built from supported_languages().

    The user can type a language code (e.g. "tr", "en") or a display name
    (e.g. "Türkçe", "English"). Empty input falls back to DEFAULT_LANG.
    """
    codes = supported_languages()
    options = " / ".join(f"{c} ({get_locale(c)['name']})" for c in codes)
    print(f"\nLanguage / Dil [{options}] (default: {DEFAULT_LANG}): ", end="", flush=True)
    try:
        choice = input()
    except (EOFError, KeyboardInterrupt):
        print()
        return DEFAULT_LANG
    return resolve_lang(choice)


def main() -> None:
    lang = pick_language()
    msgs = get_locale(lang)
    print(msgs["banner"])

    try:
        agent = build_agent(lang=lang)
    except Exception as e:
        print(msgs["err_init"].format(e=e))
        sys.exit(1)

    print(msgs["ready"])

    while True:
        try:
            user_input = input(msgs["user"]).strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{msgs['bye']}")
            break

        if not user_input:
            continue

        lowered = user_input.lower()
        if lowered in msgs["exit_words"]:
            print(msgs["bye"])
            break

        # `:new` rotates the rolling session id. The old session stays in the
        # DB so search_session_history can still surface it; the next turn
        # just starts a fresh thread.
        if lowered in {":new", "/new", "yeni oturum", "new session"}:
            new_id = start_new_session()
            agent = build_agent(lang=lang)
            print(msgs["new_session"].format(session_id=new_id[:8]))
            continue

        print()
        try:
            agent.print_response(user_input, stream=True)
        except Exception as e:
            print(msgs["err_run"].format(e=e))
        print()


if __name__ == "__main__":
    main()
