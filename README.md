# Microsoft Teams & Outlook MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that provides access to Microsoft Teams chats/channels, Outlook emails, and SharePoint files via the Microsoft Graph API.

Use natural language in Claude Code, VS Code, Claude Desktop, or any MCP client to read Teams messages, search emails, browse files, and more.

## Features

- **Teams** — List teams/channels, read/send channel messages, read/send 1:1 & group chat messages
- **Outlook** — List emails, read email body, search emails, list mail folders
- **SharePoint** — List channel files, read file contents (text/xlsx)
- **Auth Management** — Check auth status and authenticate via Device Code Flow directly from MCP

## Prerequisites

### 1. Register an Azure AD App

1. Go to [Azure Portal](https://portal.azure.com/) → **Azure Active Directory** → **App registrations** → **New registration**
2. Under **Certificates & secrets**, create a client secret
3. Under **API permissions**, add these Microsoft Graph **delegated permissions**:

| Permission | Purpose |
|------------|---------|
| `User.Read` | User profile |
| `Mail.Read` | Read emails |
| `Chat.Read` | Read chats |
| `Chat.ReadWrite` | Send chat messages |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Send channel messages |
| `Team.ReadBasic.All` | List teams |
| `Files.Read.All` | Read SharePoint files |

4. Click **Grant admin consent**

### 2. Note Your Credentials

You will need these three values when registering the MCP server:

- `MS_CLIENT_ID` — Application (client) ID
- `MS_CLIENT_SECRET` — Client secret value
- `MS_TENANT_ID` — Directory (tenant) ID

### 3. Install uv (if not installed)

This project uses [`uvx`](https://docs.astral.sh/uv/) to run the MCP server. Install `uv` first:

```powershell
# Windows (PowerShell) — 공식 설치 스크립트
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

설치 후 터미널을 재시작하고, 아래 명령어로 확인:

```bash
uvx --version
```

## Installation & Usage

### Register with Claude Code (uvx — Recommended)

No pre-installation needed — `uvx` automatically downloads and runs the package.

```bash
# 1. Authenticate (one-time — credentials are passed as CLI args)
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

### VS Code (Windows / macOS / Linux)

VS Code supports MCP servers via the built-in Copilot or Claude extension.

#### Option A: Workspace config (`.vscode/mcp.json`)

Create `.vscode/mcp.json` in your project root:

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

#### Option B: User settings (`settings.json`)

Open VS Code Settings (JSON) and add:

```json
{
  "mcp": {
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
}
```

> **Windows note**: If `uvx` is not found, try `uvx.cmd` as the command, or use the full path (e.g., `C:\\Users\\<you>\\AppData\\Roaming\\uv\\uvx.cmd`).

#### Authenticate before first use

```bash
# Windows (PowerShell)
uvx --from "git+https://github.com/giljoonseok/ms-teams-mcp.git" `
  ms-teams-mcp auth `
  --client-id <your-client-id> `
  --client-secret <your-client-secret> `
  --tenant-id <your-tenant-id>

# macOS / Linux
uvx --from git+https://github.com/giljoonseok/ms-teams-mcp.git \
  ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>
```

### Claude Desktop

Add to your Claude Desktop config file:

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

### Install via pip

```bash
# Install from Git
pip install git+https://github.com/giljoonseok/ms-teams-mcp.git

# Authenticate
ms-teams-mcp auth \
  --client-id <your-client-id> \
  --client-secret <your-client-secret> \
  --tenant-id <your-tenant-id>

# Register with Claude Code
claude mcp add microsoft-teams ms-teams-mcp \
  -s user \
  -e MS_CLIENT_ID=<your-client-id> \
  -e MS_CLIENT_SECRET=<your-client-secret> \
  -e MS_TENANT_ID=<your-tenant-id>
```

### Development Setup

```bash
git clone https://github.com/giljoonseok/ms-teams-mcp.git
cd microsoft-teams-mcp
pip install -e .
```

### Upgrade

```bash
# uvx (Claude Code / VS Code / Claude Desktop) — 캐시 초기화 후 다음 실행 시 자동 최신 다운로드
uv cache clean

# pip install from git — 강제 재설치
pip install --upgrade --force-reinstall git+https://github.com/giljoonseok/ms-teams-mcp.git

# Development setup — 최신 코드 pull 후 재설치
git pull && pip install -e .
```

### Verify & Remove

```bash
claude mcp list                       # List registered MCP servers
claude mcp remove microsoft-teams     # Remove server
```

### Example Prompts

Once registered, use natural language in your MCP client:

```
> Show my Teams list
> Show the last 10 messages in the "General" channel of "Project A" team
> Check my inbox for today's emails
> Search emails for "meeting notes"
> Check my auth status
```

## Available MCP Tools

| Category | Tool | Description |
|----------|------|-------------|
| **Auth** | `auth_status` | Check authentication status and token validity |
| | `authenticate` | Authenticate via Device Code Flow within MCP |
| **Teams** | `list_teams` | List joined teams |
| | `list_channels(team_id)` | List channels in a team |
| | `list_channel_messages(team_id, channel_id, top, skip)` | Read channel messages (max 50) |
| | `send_channel_message(team_id, channel_id, message)` | Send a message to a channel |
| **Chats** | `list_chats(top, skip)` | List 1:1 and group chats (max 50) |
| | `list_chat_messages(chat_id, top, skip)` | Read chat messages (max 50) |
| | `send_chat_message(chat_id, message)` | Send a chat message |
| **Outlook** | `list_emails(folder, top, skip)` | List emails (max 1000) |
| | `read_email(message_id)` | Read full email body |
| | `search_emails(query, top, skip)` | Search emails (max 1000) |
| | `list_mail_folders` | List mail folders |
| **SharePoint** | `list_channel_files(team_id, channel_id, top, skip)` | List files in a channel (max 200) |
| | `read_channel_file(drive_id, item_id)` | Read file contents (text/xlsx, 5MB limit) |

## Project Structure

```
microsoft-teams-mcp/
├── ms_teams_mcp/
│   ├── __init__.py    # Package init
│   └── server.py      # MCP server
├── pyproject.toml     # Package config & dependencies
├── .gitignore
├── CLAUDE.md          # Claude Code instructions
├── TODO.md            # Roadmap
└── README.md
```

## License

MIT
