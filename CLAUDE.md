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

Package structure — source lives in `ms_teams_mcp/`:

1. **Config & MSAL Init** — Lazy initialization: `_get_config()` reads env vars on first use (not at import). `_get_app()`/`_get_pub_app()` create MSAL apps lazily. Token cache: `~/.ms_mcp_token.json`
2. **Auth Layer** — `get_token()` attempts silent renewal → raises error on failure. CLI `cmd_auth()` and MCP tool `authenticate()` both use Device Code Flow via `_get_pub_app()`
3. **Graph API Helpers** — `graph_get(path, params, url)` / `graph_post(path, body)`: all Microsoft Graph calls MUST go through these two functions. `url` parameter allows direct nextLink calls.
4. **MCP Tools (21)** — Registered via `@mcp.tool()` decorator. All tools return formatted English strings (not JSON)
5. **CLI Entrypoint** — `main()` branches on `sys.argv`: no args → `mcp.run()`, `auth` → `_parse_auth_args()` + `cmd_auth()`, `--version` → print version

## Conventions

- **English only in source code**: All comments, docstrings, and user-facing output strings must be written in English. Do not use Korean or other non-English languages in source files.
- **User confirmation before sending**: All send/reply/forward/create tools (`send_channel_message`, `reply_to_channel_message`, `send_chat_message`, `reply_to_chat_message`, `create_chat`, `send_email`, `reply_email`, `forward_email`, `create_calendar_event`) MUST show the full content (recipients, subject, body) to the user and receive explicit confirmation before being called. Never send automatically.
- Adding a new tool: add `@mcp.tool()` function in `ms_teams_mcp/server.py`, always use `graph_get()`/`graph_post()` for Graph API calls
- `top` parameter limits: chats/channels/messages max 50, emails max 1000, email search max 1000, files max 200
- Pagination: all list tools support `skip` parameter. `_pagination_footer()` appends next-page guidance when `@odata.nextLink` exists
- HTML bodies are stripped via `strip_html()` before returning
- File reading: tries UTF-8 → CP949 decoding; xlsx parsed with `openpyxl`
- Windows compatibility: guard `sys.stdout`/`sys.stderr` with `is not None` check before `reconfigure()`
- Lazy init: config and MSAL apps are created on first use, not at import — allows `auth` CLI to accept `--client-id`/`--client-secret`/`--tenant-id` args
- Version: single source in `pyproject.toml`, read via `importlib.metadata` in `server.py`
- Scopes: `Mail.Read`, `Mail.Send`, `User.Read`, `Chat.Read`, `Chat.ReadWrite`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Team.ReadBasic.All`, `Files.Read.All`, `People.Read`, `Calendars.ReadWrite`

## MCP Registration

### Claude Code (CLI)

```bash
ms-teams-mcp auth --client-id <ID> --client-secret <SECRET> --tenant-id <TENANT>
claude mcp add microsoft-teams \
  -s user \
  -e MS_CLIENT_ID=<ID> -e MS_CLIENT_SECRET=<SECRET> -e MS_TENANT_ID=<TENANT> \
  -- ms-teams-mcp
```

Verify: `claude mcp list` | Remove: `claude mcp remove microsoft-teams`

### VS Code (Copilot / Claude Extension)

Add to `.vscode/mcp.json` in workspace root:

```json
{
  "servers": {
    "microsoft-teams": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/giljoonseok/ms-teams-mcp.git", "ms-teams-mcp"],
      "env": {
        "MS_CLIENT_ID": "<your-client-id>",
        "MS_CLIENT_SECRET": "<your-client-secret>",
        "MS_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
}
```

On Windows, use `uvx.cmd` instead of `uvx` if `uvx` is not found.
