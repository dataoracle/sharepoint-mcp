# sharepoint-mcp: SharePoint MCP server

An MCP server that lets Claude (or any MCP client) work with SharePoint as you:
find and browse sites, libraries and folders, search files by content, read
Word/Excel/PowerPoint/PDF files as text, upload and update files, and read and
write SharePoint lists and pages.

It is the sibling of [msteams-mcp](https://github.com/dataoracle/msteams-mcp)
and uses the same model: **delegated** Microsoft Graph permissions, a one-time
interactive sign-in on the server (OAuth2 authorization code + PKCE), and a
refresh token that the server keeps rotating by itself. Everything it does
appears in SharePoint as done by you, and it can only reach what your account
can reach.

It runs as a hardened systemd service. No credential is stored in plain text:

| Secret | Where it lives |
|---|---|
| Entra app client secret | TPM-sealed systemd credential |
| MCP access token | TPM-sealed systemd credential |
| Vault key (AES-256) | TPM-sealed systemd credential |
| Microsoft refresh token | `/var/lib/sharepoint-mcp/vault.json.enc`, AES-256-GCM encrypted with the vault key |
| Access tokens (1 h) | Process memory only |

**Why a vault?** Entra ID *rotates* refresh tokens: every refresh returns a new
one, and each one expires after about 90 days. The server has to save each new
token while it runs, but the sandboxed service can't use the TPM directly. So
the TPM seals a key, and that key encrypts the token file. A copy of the disk
or a backup without this machine's TPM can't read the token.

## Tools

**Addressing files and folders.** Every listing prints a **ref**
(`<drive_id>/<item_id>`) for each file and folder. Tools that take `item`,
`folder` or `destination` accept that ref, **or a SharePoint URL** (a file or
folder URL copied from the browser, a library view URL, or a sharing link).
Most also accept `site` + `library` + `path` instead. `site` itself can be a
URL, a site id, `root`, or a site name.

| Tool | Returns / does |
|---|---|
| `whoami()` | The signed-in account and the state of the sign-in |
| `list_sites(query?, limit?)` | Followed sites, or sites matching `query` (`'*'` = every site you can access) |
| `get_site(site)` | Site details and its document libraries, with each library's root ref |
| `list_folder(item? \| site, library?, path?, limit?)` | Folder contents, folders first: name, size, modified, ref |
| `get_item(item? \| site, library?, path?)` | One file/folder in full: path, type, version, check-out state, custom column values, link |
| `search_files(query?, site?, library?, folder?, file_types?, sort?, limit?, offset?)` | Full-text search (Microsoft Search, KQL) scoped to a site, library or folder. `sort="modified"` with no query = recently changed files |
| `read_file(item? \| site, library?, path?, offset?, max_chars?, version?)` | A file's text, paged. Word (headings, lists, tables), Excel (all sheets, dates converted), PowerPoint (slides + notes), PDF, legacy Office/ODF (converted to PDF by Graph), text and code formats |
| `list_versions(item)` | Version history; `read_file(version=…)` reads an older one |
| `list_lists(site)` | Lists and libraries with template and id |
| `get_list_schema(site, list)` | Columns: internal name, type, choices, required |
| `list_items(site, list, filter?, order_by?, fields?, limit?)` | List items; OData filter on `fields/…` |
| `get_list_item(site, list, item_id)` | One list item in full |
| `list_pages(site)` / `read_page(site?, page)` | Site pages and news posts, and the text of one |
| `write_file(content, name, folder \| site, library?, path? — or item, encoding?, conflict?)` ✏️ | Create a file, or replace a file's content (new version). `encoding="base64"` for binary files |
| `create_folder(name, folder \| site, library?, path?)` ✏️ | New folder |
| `move_item(item, destination?, new_name?)` ✏️ | Move within a library and/or rename |
| `copy_item(item, destination, new_name?)` ✏️ | Copy anywhere, across libraries and sites |
| `update_file_fields(item, fields)` ✏️ | Set a file's library column values (metadata) |
| `create_list_item(site, list, fields)` ✏️ | New list item |
| `update_list_item(site, list, item_id, fields)` ✏️ | Change list item fields |
| `delete_item(item)` 🗑️ | Delete a file or folder (to the recycle bin) |
| `delete_list_item(site, list, item_id)` 🗑️ | Delete a list item (to the recycle bin) |

✏️ = write tools, registered only when `ENABLE_WRITE_TOOLS=true` (the default).
🗑️ = delete tools, registered only when `ENABLE_DELETE_TOOLS=true` (**off by
default**), and never on a read-only server. Deleted items can be restored from
the site's recycle bin for 93 days.

**Output.** Read tools return compact text to keep the agent's token use low
(times in UTC). Every read tool also takes `output="json"` for structured data.
For example, `list_folder` returns:

```
Reports: 3 items (folders first; times UTC) · folder ref b!Xy…/01ABC
• Archive/ (3 items) · modified 2026-11-10 09:30 Alice Johnson · ref b!Xy…/01DEF
• Report.docx · 48.2 KB · modified 2026-11-10 09:30 Alice Johnson · ref b!Xy…/01GHI
```

and `read_file` on a long document ends with
`(chars 0-19987 of 85210: call again with offset=19987 for more)`.

## Requirements

- **Server:** Linux with systemd 250 or newer and a TPM 2.0 chip (for example
  Ubuntu 24.04 or Debian 12), plus Python 3.10+ with `venv` support.
- **Entra ID app registration** (App ID + tenant ID):
  - Delegated permissions with admin consent: `openid`, `offline_access`,
    `User.Read`, `Sites.ReadWrite.All`. The server requests
    `https://graph.microsoft.com/.default`, so it gets exactly what has been
    consented. See [Permissions](#permissions) for optional extras.
  - **Authentication → Add a platform → Web**, with the redirect URI
    `http://localhost`. You can use another localhost URI such as
    `http://localhost:8400/callback`; set `MS_REDIRECT_URI` to the exact same
    value. Nothing needs to listen on it: you copy the code from the browser's
    address bar.
  - **Certificates & secrets → New client secret.** Copy its *Value* (not the
    Secret ID) straight into `seal-credential.sh client_secret`. Note the expiry
    date (see [When the client secret expires](#when-the-client-secret-expires)).

  The same app registration as msteams-mcp can be reused if you add
  `Sites.ReadWrite.All` to it; each server has its own sign-in and vault.

## Permissions

`Sites.ReadWrite.All` (delegated) covers everything above: sites, document
libraries, files, lists, pages and Microsoft Search over SharePoint content.
Being delegated, it never goes beyond what the signed-in user can access.

Optional additions, only if you want what they unlock:

| Permission | Adds |
|---|---|
| `Files.ReadWrite.All` | Your OneDrive and "Shared with me" through the same tools (not exposed as tools yet) |
| `Sites.Manage.All` | Creating new lists and libraries (not exposed as tools yet) |

## Installation

Run these on the server from a clone of this repository.

```bash
# 1. Install code, venv, service user, config file, vault key and systemd unit
sudo ./deploy/install.sh

# 2. Set your tenant and app id
sudo nano /etc/sharepoint-mcp/sharepoint-mcp.env

# 3. Seal the secrets (input is read without echo and never written to disk;
#    vault_key was already generated by install.sh)
sudo /opt/sharepoint-mcp/bin/seal-credential.sh client_secret          # paste the app's secret Value
sudo /opt/sharepoint-mcp/bin/seal-credential.sh mcp_token --generate

# 4. Sign in to Microsoft (see below)
sudo sharepoint-mcp-authorize

# 5. Start the service and enable it at boot
sudo systemctl enable --now sharepoint-mcp

# 6. Check it
systemctl status sharepoint-mcp
journalctl -u sharepoint-mcp -n 30
curl -s localhost:8767/health                 # {"status":"ok","auth":"ok"}
```

Show the generated MCP token once, to put in your MCP client config:

```bash
sudo systemd-creds decrypt --name=mcp_token /etc/credstore.encrypted/sharepoint-mcp.mcp_token.cred -; echo
```

The service listens on port **8767**, so it can run next to msteams-mcp (8766)
on the same host.

### Signing in

`sudo sharepoint-mcp-authorize`:

1. prints a `login.microsoftonline.com` URL. Open it in a browser on any
   machine and sign in with the account the server should act as;
2. the browser is redirected to `http://localhost/?code=…&state=…` and shows
   "can't connect", which is expected;
3. you copy that full URL from the address bar and paste it at the prompt
   (input is hidden). The code is single-use, expires in about 10 minutes, and
   can only be redeemed with both the client secret and the PKCE verifier held
   by the script;
4. the script decrypts the client secret into memory, exchanges the code for
   tokens, checks `/me`, and writes the refresh token to the encrypted vault,
   owned by the service user.

The running service picks up the new vault on its next request, so you don't
need to restart it.

### When you have to sign in again

The server keeps its sign-in alive by itself: every refresh rotates the stored
refresh token, and a background refresh every `KEEPALIVE_HOURS` (12 by default)
keeps it fresh even when idle. Run `sudo sharepoint-mcp-authorize` again only
when Entra ID rejects the token: password change or reset, revoked sessions or
consent, a Conditional Access sign-in frequency or MFA requirement, or the
server being off for longer than the refresh token's lifetime (~90 days).

Every tool then returns *"The saved Microsoft sign-in is no longer valid
(AADSTS…). Run on the server: sudo sharepoint-mcp-authorize"*, and `/health`
reports `"auth":"reauth_required"`.

### When the client secret expires

No new sign-in is needed: refresh tokens belong to the App ID, not to a
particular secret. Tools return *"Entra ID rejected the app's client secret
(invalid_client: AADSTS7000222…)"*. Create a new secret in the app registration, then:

```bash
sudo /opt/sharepoint-mcp/bin/seal-credential.sh client_secret
sudo systemctl restart sharepoint-mcp
```

Do this before the old secret expires: Entra allows several secrets at once.

## Connecting a client

The service listens on `127.0.0.1:8767` only. The MCP endpoint is `/mcp`, and
every request except `/health` must send `Authorization: Bearer <token>`. To
reach it from another machine, put an HTTPS front end in front of the port (a
reverse proxy, a Tailscale service, or an SSH tunnel).

```json
{
  "mcpServers": {
    "sharepoint_mcp": {
      "command": "npx",
      "args": ["mcp-remote", "https://your-host.example.com/mcp",
               "--header", "Authorization:${AUTH}"],
      "env": { "AUTH": "Bearer <your-token>" }
    }
  }
}
```

Then ask Claude "which SharePoint sites do I follow?" to confirm it works end to end.

## Agent skill

[`skills/sharepoint/`](skills/sharepoint/SKILL.md) is an
[Agent Skill](https://agentskills.io) that teaches an agent to use these tools
well: how to find content, read long files efficiently, work with lists, and
confirm before changing anything.

- **Claude Code:** `cp -r skills/sharepoint ~/.claude/skills/`, then start a new session.
- **claude.ai / Claude Desktop:** `cd skills && zip -r sharepoint.zip sharepoint`,
  then upload it under Settings → Capabilities → Skills.

## Layout on the server

| Path | What |
|---|---|
| `/opt/sharepoint-mcp/app` | Application code (read-only to the service) |
| `/opt/sharepoint-mcp/venv` | Python virtual environment |
| `/opt/sharepoint-mcp/bin` | `authorize.sh`, `seal-credential.sh` |
| `/usr/local/sbin/sharepoint-mcp-authorize` | Symlink to the sign-in script |
| `/etc/sharepoint-mcp/sharepoint-mcp.env` | Non-secret settings |
| `/etc/credstore.encrypted/sharepoint-mcp.*.cred` | TPM-sealed `client_secret`, `mcp_token` and `vault_key` |
| `/var/lib/sharepoint-mcp/vault.json.enc` | Encrypted refresh token (0600, user `sharepoint-mcp`) |
| `/etc/systemd/system/sharepoint-mcp.service` | The unit |

## Configuration

Settings in `/etc/sharepoint-mcp/sharepoint-mcp.env` (see [`deploy/sharepoint-mcp.env.example`](deploy/sharepoint-mcp.env.example)):

| Variable | Default | Meaning |
|---|---|---|
| `MS_TENANT_ID` | required | Tenant GUID or domain |
| `MS_CLIENT_ID` | required | Application (client) ID |
| `MS_REDIRECT_URI` | `http://localhost` | Must exactly match a redirect URI registered under the *Web* platform |
| `ENABLE_WRITE_TOOLS` | `true` | `false` = read-only server |
| `ENABLE_DELETE_TOOLS` | `false` | `true` adds `delete_item` and `delete_list_item` |
| `KEEPALIVE_HOURS` | `12` | Background token refresh interval; `0` disables it |
| `MAX_READ_CHARS` | `20000` | Default page size of `read_file` / `read_page` |
| `MAX_DOWNLOAD_MB` | `50` | Largest file `read_file` downloads |
| `MAX_UPLOAD_MB` | `50` | Largest file `write_file` accepts (above 4 MB it uses an upload session) |

Restart the service after changing settings.

## Day-to-day

```bash
sudo systemctl restart sharepoint-mcp            # after changing settings or credentials
journalctl -u sharepoint-mcp -f                   # follow logs
systemd-analyze security sharepoint-mcp           # hardening score (lower is better)
sudo sharepoint-mcp-authorize                      # sign in again / switch account
sudo ./deploy/install.sh                           # update code; keeps config, credentials and sign-in

# Rotate the MCP token
sudo /opt/sharepoint-mcp/bin/seal-credential.sh mcp_token --generate
sudo systemctl restart sharepoint-mcp              # then update your MCP client

# Rotate the vault key (makes the stored sign-in unreadable: sign in again)
sudo /opt/sharepoint-mcp/bin/seal-credential.sh vault_key --generate --force
sudo systemctl restart sharepoint-mcp && sudo sharepoint-mcp-authorize
```

## Notes

- **Downloads and uploads.** File content comes from the short-lived,
  pre-authenticated SharePoint URL that Graph redirects to; large uploads go to
  a Graph upload session URL. The server fetches those with a separate HTTP
  client that never sends the Graph token.
- **File parsing** is done on the server with the standard library (Office
  Open XML through `defusedxml`) and `pypdf`, with size limits against zip
  bombs. Scanned PDFs without a text layer can't be read (no OCR). Excel
  formulas are read as their last calculated values.
- **Search** uses the Microsoft Search API: results can lag a few minutes
  behind new or changed files. `list_folder` is always current.
- **List filters** on columns that aren't indexed are sent with
  `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`; on very large lists
  (over 5,000 items) SharePoint may still refuse them: index the column, or
  use `search_files`.
- **After a BIOS/firmware update or a Secure Boot change,** the TPM may refuse
  to decrypt the credentials. Re-seal `client_secret` and `mcp_token`,
  regenerate `vault_key` with `--force`, and sign in again.
- **What this protects against:** copies of the disk, backups, sync tools or a
  git push picking up files, and other users or services on the server. It
  does not stop someone with root on the running machine. The refresh token is
  as powerful as your SharePoint access within the consented permissions.
- **Conditional Access:** if your tenant restricts sign-ins to compliant
  devices or known locations, refreshes from the server may be blocked
  (`AADSTS53003`, `AADSTS50076`...). The journal shows the code.
- If the service fails with a memory-protection error, comment out
  `MemoryDenyWriteExecute=yes` in the unit, then `daemon-reload` and restart.

## Local development

Without systemd, secrets fall back to the `MS_CLIENT_SECRET` / `MCP_TOKEN` /
`VAULT_KEY` environment variables (or to files named by `*_FILE`), and the
vault is `./vault.json.enc` (override with `VAULT_PATH`).

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
export MS_TENANT_ID=... MS_CLIENT_ID=... MCP_TOKEN=dev-token
read -rs MS_CLIENT_SECRET && export MS_CLIENT_SECRET      # paste, no echo, not in history
export VAULT_KEY=$(python -c 'import vault; print(vault.new_key_text())')
python authorize.py                             # interactive sign-in -> ./vault.json.enc
python server.py                                # serves http://127.0.0.1:8767/mcp
```

Offline tests (Graph, SharePoint downloads and Entra ID are mocked, no tenant needed):

```bash
python tests/test_sharepoint.py -v              # -v prints sample output
```

## License

MIT, see [LICENSE](LICENSE).
