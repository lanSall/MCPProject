# Outlook MCP

Blank repo no more — this is a local **stdio MCP server** for Microsoft Outlook via Microsoft Graph.

## What you need on your PC

1. **Node.js 18+** (`node -v`)
2. Your Azure app **Application (client) ID** and **Directory (tenant) ID**
3. Azure app settings:
   - **Authentication** → **Allow public client flows** = **Yes** (required for device-code login)
   - **API permissions** → Microsoft Graph **Delegated**:
     - `User.Read`
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.Read`
     - `offline_access` (usually included automatically)
   - Grant admin consent if your tenant requires it

A **client secret is not required** for this server (device-code / public client). If you previously pasted a secret in chat, **rotate/delete it in Azure** anyway.

## Install (on your Windows machine)

From this folder (`MCPProject` / this repo):

```powershell
npm install
```

Optional one-time login test in a terminal:

```powershell
$env:MS_CLIENT_ID="your-client-id"
$env:MS_TENANT_ID="your-tenant-id"
npm run login
```

Open the URL printed in the terminal, enter the code, sign in with your Microsoft account.

## Cursor `mcp.json` config

Use this (Desktop). Point at the **folder** (uses `package.json` `"main"`) or the entry file:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "node",
      "args": ["C:\\Users\\lance\\OneDrive\\MCPProject\\src\\index.js"],
      "env": {
        "MS_CLIENT_ID": "your-client-id",
        "MS_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

Keep your OrCAD / MATLAB entries as they are. **Do not put client secrets in chat or commit them.**

After saving:

1. Cursor → **Settings → Tools & MCP**
2. Confirm `outlook-mcp` is connected (green)
3. Ask the agent to run `outlook_login` (or call any mail tool — first use triggers device code on stderr)
4. Complete the browser sign-in when prompted

Tokens are cached at `%USERPROFILE%\.outlook-mcp\msal-cache.json`.

## Tools

| Tool | Purpose |
| --- | --- |
| `outlook_login` | Device-code sign-in |
| `outlook_auth_status` | Check sign-in |
| `outlook_sign_out` | Clear cached tokens |
| `outlook_get_profile` | Who am I |
| `outlook_list_emails` | List folder messages |
| `outlook_get_email` | Read one message |
| `outlook_search_emails` | Search mailbox |
| `outlook_send_email` | Send mail |
| `outlook_create_draft` | Create draft |
| `outlook_list_calendar` | Upcoming events |

## Cloud Agents note

This stdio server runs on **your Windows PC**. Cursor **Cloud Agents** will not see it unless you also add a separate Cloud/HTTP MCP. For Desktop Agent chat on your machine, this config is enough.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `Missing MS_CLIENT_ID` | Add it under `env` in `mcp.json` |
| Device code never appears | Check MCP server logs in Cursor; run `npm run login` in a terminal |
| `AADSTS7000218` / public client errors | Enable **Allow public client flows** on the app |
| `AADSTS65001` / consent | Add Graph delegated permissions; admin consent if needed |
| Wrong account type | Set tenant to `common` or `consumers` for personal Microsoft accounts |
