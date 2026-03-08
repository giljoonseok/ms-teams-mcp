# Microsoft Teams & Outlook MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that provides access to Microsoft Teams chats/channels, Outlook emails, and SharePoint files via the Microsoft Graph API.

Use natural language in Claude Code, VS Code, Claude Desktop, or any MCP client to read Teams messages, search emails, browse files, and more.

## Features

- **Teams** — List teams/channels, read/send channel messages, create/read/send 1:1 & group chats
- **Outlook** — List/read/search/send/reply/forward emails, list mail folders
- **SharePoint** — List channel files, read file contents (text/xlsx)
- **Auth** — Check auth status, authenticate via Device Code Flow directly from MCP

## Prerequisites

### 1. Register an Azure AD App

1. Go to [Azure Portal](https://portal.azure.com/) → **Azure Active Directory** → **App registrations** → **New registration**
2. Under **Certificates & secrets**, create a client secret
3. Under **API permissions**, add these Microsoft Graph **delegated permissions**:

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

4. Click **Grant admin consent**

### 2. Note Your Credentials

You will need these three values:

- `MS_CLIENT_ID` — Application (client) ID
- `MS_CLIENT_SECRET` — Client secret value
- `MS_TENANT_ID` — Directory (tenant) ID

### 3. Install uv

This project uses [`uvx`](https://docs.astral.sh/uv/) to run the MCP server:

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Restart your terminal, then verify: `uvx --version`

## Installation & Usage

### Claude Code (Recommended)

No pre-installation needed — `uvx` automatically downloads and runs the package.

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

### VS Code

Add to `.vscode/mcp.json` in your project root:

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

Or add the same config under `"mcp"` key in your user `settings.json` for global access.

> **Windows**: If `uvx` is not found, use `uvx.cmd` or the full path (e.g., `C:\Users\<you>\AppData\Roaming\uv\uvx.cmd`).

### Claude Desktop

Add to your config file:

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

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

### pip Install

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

### Authentication

All installation methods require a one-time Device Code Flow authentication:

```bash
uvx --from git+https://github.com/giljoonseok/ms-teams-mcp.git \
  ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>
```

You can also authenticate from within any MCP client by calling the `authenticate` tool.

### Upgrade

```bash
# uvx — clear cache, auto-downloads latest on next run
uv cache clean

# pip — force reinstall
pip install --upgrade --force-reinstall git+https://github.com/giljoonseok/ms-teams-mcp.git
```

### Verify & Remove

```bash
claude mcp list                       # List registered MCP servers
claude mcp remove microsoft-teams     # Remove server
```

## Available MCP Tools

| Category | Tool | Description |
|----------|------|-------------|
| **Auth** | `auth_status` | Check authentication status and token validity |
| | `authenticate` | Authenticate via Device Code Flow |
| **Teams** | `list_teams` | List joined teams |
| | `list_channels` | List channels in a team |
| | `list_channel_messages` | Read channel messages (max 50) |
| | `send_channel_message` | Send a message to a channel |
| **Chats** | `list_chats` | List 1:1 and group chats (max 50) |
| | `list_chat_messages` | Read chat messages (max 50) |
| | `send_chat_message` | Send a chat message |
| | `create_chat` | Create a new 1:1 or group chat |
| **Outlook** | `list_emails` | List emails (max 1000) |
| | `read_email` | Read full email body |
| | `search_emails` | Search emails (max 1000) |
| | `send_email` | Send an email |
| | `reply_email` | Reply or reply-all to an email |
| | `forward_email` | Forward an email |
| | `list_mail_folders` | List mail folders |
| **Files** | `list_channel_files` | List files in a channel (max 200) |
| | `read_channel_file` | Read file contents (text/xlsx, 5MB limit) |

All list tools support `top`, `skip`, and `next_link` parameters for pagination.

### Example Prompts

```
> Show my Teams list
> Show the last 10 messages in the "General" channel of "Project A" team
> Check my inbox for today's emails
> Search emails for "meeting notes"
> Send an email to john@example.com about the project update
```

## Project Structure

```
microsoft-teams-mcp/
├── ms_teams_mcp/
│   ├── __init__.py    # Package init
│   └── server.py      # MCP server (single-file architecture)
├── pyproject.toml     # Package config & dependencies
├── CLAUDE.md          # Claude Code instructions
└── README.md
```

## License

MIT
