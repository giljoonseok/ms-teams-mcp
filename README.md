# Microsoft Teams & Outlook MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that provides access to Microsoft Teams chats/channels, Outlook emails, calendar events, and SharePoint files via the Microsoft Graph API.

Use natural language in Claude Code, VS Code, Claude Desktop, or any MCP client to read Teams messages, search emails, manage calendar events, browse files, and more.

## Features

- **Teams** — List teams/channels, read/send/reply channel messages, create/read/send 1:1 & group chats
- **Outlook** — List/read/search/send/reply/forward emails, list mail folders
- **Calendar** — List/create/update/delete events, recurring events, reminders
- **SharePoint** — List channel files, read file contents (text/xlsx), build & search file index
- **People** — Search users across the organization
- **Utilities** — Unread summary, auth status, version update check
- **Auth** — Device Code Flow authentication directly from MCP or CLI

## Prerequisites

### 1. Register an Azure AD App

1. Go to [Azure Portal](https://portal.azure.com/) > **Azure Active Directory** > **App registrations** > **New registration**
2. Set **Redirect URI** to `https://login.microsoftonline.com/common/oauth2/nativeclient` (Mobile and desktop applications)
3. Under **Certificates & secrets**, create a client secret
4. Under **API permissions**, add these Microsoft Graph **delegated permissions**:

| Permission | Purpose |
|------------|---------|
| `User.Read` | User profile |
| `Mail.Read` | Read emails |
| `Mail.Send` | Send, reply, forward emails |
| `Chat.Read` | Read chats |
| `Chat.ReadWrite` | Send chat messages, create chats |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Send channel messages |
| `Team.ReadBasic.All` | List teams |
| `Files.Read.All` | Read SharePoint files |
| `People.Read` | Search users |
| `Calendars.ReadWrite` | Read/create/update/delete calendar events |

5. Click **Grant admin consent**

### 2. Note Your Credentials

You will need these three values:

- `MS_CLIENT_ID` — Application (client) ID
- `MS_CLIENT_SECRET` — Client secret value
- `MS_TENANT_ID` — Directory (tenant) ID

### 3. Install uv

This project uses [`uvx`](https://docs.astral.sh/uv/) to run the MCP server:

**Linux / macOS:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**Windows (PowerShell):**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Restart your terminal, then verify: `uvx --version`

## Installation & Usage

### Claude Code (Recommended)

No pre-installation needed — `uvx` automatically downloads and runs the package.

**Linux / macOS:**
```bash
# 1. Authenticate (one-time)
uvx --from git+https://github.com/giljoonseok/ms-teams-mcp.git \
  ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>

# 2. Register with Claude Code
claude mcp add microsoft-teams \
  -s user \
  -e MS_CLIENT_ID=<your-client-id> \
  -e MS_CLIENT_SECRET=<your-client-secret> \
  -e MS_TENANT_ID=<your-tenant-id> \
  -- uvx --from git+https://github.com/giljoonseok/ms-teams-mcp.git ms-teams-mcp
```

**Windows (PowerShell):**
```powershell
# 1. Authenticate (one-time)
uvx --from "git+https://github.com/giljoonseok/ms-teams-mcp.git" `
  ms-teams-mcp auth `
  --client-id <your-client-id> `
  --client-secret <your-client-secret> `
  --tenant-id <your-tenant-id>

# 2. Register with Claude Code
claude mcp add microsoft-teams `
  -s user `
  -e MS_CLIENT_ID=<your-client-id> `
  -e MS_CLIENT_SECRET=<your-client-secret> `
  -e MS_TENANT_ID=<your-tenant-id> `
  -- uvx --from "git+https://github.com/giljoonseok/ms-teams-mcp.git" ms-teams-mcp
```

> **Windows Note**: If `uvx` is not found, use `uvx.cmd` or the full path `%USERPROFILE%\.local\bin\uvx.cmd`.

### VS Code

Add to `.vscode/mcp.json` in your project root:

**Linux / macOS:**
```json
{
  "servers": {
    "microsoft-teams": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/giljoonseok/ms-teams-mcp.git",
        "ms-teams-mcp"
      ],
      "env": {
        "MS_CLIENT_ID": "<your-client-id>",
        "MS_CLIENT_SECRET": "<your-client-secret>",
        "MS_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
}
```

**Windows:**
```json
{
  "servers": {
    "microsoft-teams": {
      "command": "uvx.cmd",
      "args": [
        "--from",
        "git+https://github.com/giljoonseok/ms-teams-mcp.git",
        "ms-teams-mcp"
      ],
      "env": {
        "MS_CLIENT_ID": "<your-client-id>",
        "MS_CLIENT_SECRET": "<your-client-secret>",
        "MS_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
}
```

Or add the same config under `"mcp"` key in your user `settings.json` for global access.

### Claude Desktop

Add to your config file:

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoft-teams": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/giljoonseok/ms-teams-mcp.git",
        "ms-teams-mcp"
      ],
      "env": {
        "MS_CLIENT_ID": "<your-client-id>",
        "MS_CLIENT_SECRET": "<your-client-secret>",
        "MS_TENANT_ID": "<your-tenant-id>"
      }
    }
  }
}
```

> **Windows**: Replace `"command": "uvx"` with `"command": "uvx.cmd"` if `uvx` is not found.

### pip Install

**Linux / macOS:**
```bash
pip install git+https://github.com/giljoonseok/ms-teams-mcp.git

ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>

claude mcp add microsoft-teams ms-teams-mcp \
  -s user \
  -e MS_CLIENT_ID=<your-client-id> \
  -e MS_CLIENT_SECRET=<your-client-secret> \
  -e MS_TENANT_ID=<your-tenant-id>
```

**Windows (PowerShell):**
```powershell
pip install "git+https://github.com/giljoonseok/ms-teams-mcp.git"

ms-teams-mcp auth `
  --client-id <your-client-id> `
  --client-secret <your-client-secret> `
  --tenant-id <your-tenant-id>

claude mcp add microsoft-teams ms-teams-mcp `
  -s user `
  -e MS_CLIENT_ID=<your-client-id> `
  -e MS_CLIENT_SECRET=<your-client-secret> `
  -e MS_TENANT_ID=<your-tenant-id>
```

### Authentication

All installation methods require a one-time Device Code Flow authentication:

```bash
# Using uvx
uvx --from git+https://github.com/giljoonseok/ms-teams-mcp.git \
  ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>

# Using pip install
ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>
```

You can also authenticate from within any MCP client by calling the `authenticate` tool.

Token is cached at `~/.ms_mcp_token.json` and silently renewed on subsequent runs.

### Upgrade

```bash
# uvx — clear cache, auto-downloads latest on next run
uv cache clean

# pip — force reinstall
pip install --upgrade --force-reinstall git+https://github.com/giljoonseok/ms-teams-mcp.git
```

The server automatically checks for updates on startup (once per 24h) and notifies via stderr. You can also call the `check_update` tool from any MCP client.

### Verify & Remove

```bash
claude mcp list                       # List registered MCP servers
claude mcp remove microsoft-teams     # Remove server
```

## Available MCP Tools (32)

| Category | Tool | Description |
|----------|------|-------------|
| **Auth** | `auth_status` | Check authentication status and token validity |
| | `authenticate` | Authenticate via Device Code Flow |
| **Teams** | `list_teams` | List joined teams |
| | `list_channels` | List channels in a team |
| | `list_channel_messages` | Read channel messages (max 50) |
| | `send_channel_message` | Send a message to a channel |
| | `reply_to_channel_message` | Reply to a channel message |
| **Chats** | `list_chats` | List 1:1 and group chats (max 50) |
| | `list_chat_messages` | Read chat messages (max 50) |
| | `send_chat_message` | Send a chat message |
| | `reply_to_chat_message` | Reply to a chat message |
| | `create_chat` | Create a new 1:1 or group chat |
| **Outlook** | `list_emails` | List emails (max 1000) |
| | `read_email` | Read full email body |
| | `search_emails` | Search emails (max 1000) |
| | `send_email` | Send an email |
| | `reply_email` | Reply or reply-all to an email |
| | `forward_email` | Forward an email |
| | `list_mail_folders` | List mail folders |
| **Calendar** | `list_calendar_events` | List calendar events in a date range |
| | `create_calendar_event` | Create a new calendar event |
| | `update_calendar_event` | Update an existing calendar event |
| | `delete_calendar_event` | Delete a calendar event |
| | `create_recurring_event` | Create a recurring event (daily/weekly/monthly/yearly) |
| | `create_reminder` | Create a reminder with alert |
| **Files** | `list_channel_files` | List files in a channel (max 200) |
| | `read_channel_file` | Read file contents (text/xlsx, 5MB limit) |
| | `build_file_index` | Build searchable index of all accessible files |
| | `search_file_index` | Search the file index by keyword |
| **People** | `search_users` | Search users in the organization |
| **Utilities** | `get_unread_summary` | Summarize unread emails and chats |
| | `check_update` | Check for newer version on GitHub |

All list tools support `top`, `skip`, and `next_link` parameters for pagination.

### Example Prompts

```
> Show my Teams list
> Show the last 10 messages in the "General" channel of "Project A" team
> Check my inbox for today's emails
> Search emails for "meeting notes"
> Send an email to john@example.com about the project update
> Show my calendar for this week
> Create a weekly recurring meeting every Monday at 10am
> Remind me about the report at 3pm tomorrow
> Search for a user named "Kim" in the organization
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `uvx` not found (Windows) | Use `uvx.cmd` or full path `%USERPROFILE%\.local\bin\uvx.cmd` |
| `uvx` not found (Linux) | Run `source ~/.bashrc` or restart terminal after uv install |
| Token expired | Re-run `ms-teams-mcp auth ...` or call `authenticate` tool |
| Permission denied (403) | Check Azure AD app has required permissions with admin consent |
| Rate limit (429) | Wait and retry — Graph API throttles excessive requests |
| `cp949` decode error | File encoding not supported; only UTF-8 and CP949 are handled |
| Server not starting | Verify env vars `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_TENANT_ID` are set |

## Project Structure

```
microsoft-teams-mcp/
├── ms_teams_mcp/
│   ├── __init__.py    # Package init
│   └── server.py      # MCP server (single-file architecture)
├── pyproject.toml     # Package config & dependencies
├── CLAUDE.md          # Claude Code instructions
├── TODO.md            # Development roadmap
└── README.md
```

## License

MIT
