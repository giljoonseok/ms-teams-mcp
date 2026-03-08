# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Quick Reference

```bash
pip install -e .              # Install in dev mode with dependencies
ms-teams-mcp                  # Run MCP server
ms-teams-mcp auth             # Authenticate via Device Code Flow (headless-friendly)
ms-teams-mcp --version        # Check version
```

No test framework. No linter/formatter configured.

## Architecture

Single-file MCP server in `ms_teams_mcp/server.py`:

1. **Config & MSAL Init** — Lazy initialization: `_get_config()` reads env vars on first use (not at import). `_get_app()`/`_get_pub_app()` create MSAL apps lazily. Token cache: `~/.ms_mcp_token.json`
2. **Auth Layer** — `get_token()` attempts silent renewal → raises error on failure. CLI `cmd_auth()` and MCP tool `authenticate()` both use Device Code Flow via `_get_pub_app()`
3. **Graph API Helpers** — All Microsoft Graph calls MUST go through these functions:
   - `graph_get(path, params, url)` — GET requests (`url` allows direct nextLink calls)
   - `graph_post(path, body)` — POST requests returning JSON
   - `graph_post_action(path, body)` — POST for 202 No Content (sendMail, reply, forward)
   - `graph_patch(path, body)` — PATCH requests (update resources)
   - `graph_delete(path)` — DELETE requests (remove resources)
4. **Shared Helpers** — `_parse_recipients()` for email addresses, `_parse_attendees()` for calendar attendees, `_pagination_footer()` for next-page guidance, `strip_html()` for HTML stripping
5. **MCP Tools (32)** — Registered via `@mcp.tool()` decorator. All tools return formatted English strings (not JSON)
6. **CLI Entrypoint** — `main()` branches on `sys.argv`: no args → `mcp.run()`, `auth` → `_parse_auth_args()` + `cmd_auth()`, `--version` → print version

## Conventions

- **English only in source code**: All comments, docstrings, and user-facing output strings must be written in English. Do not use Korean or other non-English languages in source files.
- **User confirmation before sending**: All send/reply/forward/create/update/delete tools (`send_channel_message`, `reply_to_channel_message`, `send_chat_message`, `reply_to_chat_message`, `create_chat`, `send_email`, `reply_email`, `forward_email`, `create_calendar_event`, `create_recurring_event`, `create_reminder`, `update_calendar_event`, `delete_calendar_event`) MUST show the full content to the user and receive explicit confirmation before being called. Never send automatically.
- Adding a new tool: add `@mcp.tool()` function in `ms_teams_mcp/server.py`, always use `graph_get()`/`graph_post()`/`graph_patch()`/`graph_delete()` for Graph API calls
- `top` parameter limits: chats/channels/messages max 50, emails max 1000, email search max 1000, files max 200
- Pagination: all list tools support `skip` parameter. `_pagination_footer()` appends next-page guidance when `@odata.nextLink` exists
- HTML bodies are stripped via `strip_html()` before returning
- File reading: tries UTF-8 → CP949 decoding; xlsx parsed with `openpyxl`
- Windows compatibility: guard `sys.stdout`/`sys.stderr` with `is not None` check before `reconfigure()`
- Lazy init: config and MSAL apps are created on first use, not at import — allows `auth` CLI to accept `--client-id`/`--client-secret`/`--tenant-id` args
- Version: single source in `pyproject.toml`, read via `importlib.metadata` in `server.py`
- Scopes: `Mail.Read`, `Mail.Send`, `User.Read`, `Chat.Read`, `Chat.ReadWrite`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Team.ReadBasic.All`, `Files.Read.All`, `People.Read`, `Calendars.ReadWrite`

