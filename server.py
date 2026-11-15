"""SharePoint MCP server (delegated Microsoft Graph access).

Read tools:  whoami, list_sites, get_site, list_folder, get_item, read_file,
             search_files, list_versions, list_lists, get_list_schema,
             list_items, get_list_item, list_pages, read_page.
Write tools (ENABLE_WRITE_TOOLS=true, the default):
             write_file, create_folder, move_item, copy_item, update_file_fields,
             create_list_item, update_list_item.
Delete tools (ENABLE_DELETE_TOOLS=true, off by default):
             delete_item, delete_list_item.

Files and folders are addressed by any of:
  item="<drive_id>/<item_id>"         a ref, as printed by every listing
  item="https://…sharepoint.com/…"   a SharePoint URL of the file or folder, or a sharing link
  site=… [library=…] [path=…]         a site (URL, id or name), a document library
                                      (name, default: the site's main library) and a path
"""
from __future__ import annotations

import asyncio
import base64
import binascii
import hmac
import json
import logging
import mimetypes
import os
import re
from collections import OrderedDict
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlparse

import uvicorn
from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from starlette.requests import Request
from starlette.responses import JSONResponse

import extract
import formatting as fmt
from auth import GUID_RE, AuthError, TokenManager, check_client_id, check_tenant
from graph_client import GraphClient, GraphError
from vault import Vault, VaultError, load_key, vault_context

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("sharepoint-mcp")


# ---------------------------------------------------------------- config
def _env(name: str, default: str | None = None) -> str:
    value = os.getenv(name, default)
    if value is None or value.strip() == "":
        raise SystemExit(f"Missing required environment variable {name}")
    return value.strip()


def _secret(env_name: str, cred_name: str) -> str:
    """Load a secret, preferring systemd credentials over environment variables.

    1. $CREDENTIALS_DIRECTORY/<cred_name>  (systemd LoadCredentialEncrypted=)
    2. file named by <ENV_NAME>_FILE       (e.g. Docker secrets)
    3. <ENV_NAME> environment variable      (development only)
    """
    cred_dir = os.getenv("CREDENTIALS_DIRECTORY")
    if cred_dir:
        path = Path(cred_dir) / cred_name
        if path.is_file():
            value = path.read_text(encoding="utf-8").rstrip("\r\n")
            if value:
                return value
    file_var = os.getenv(f"{env_name}_FILE")
    if file_var:
        value = Path(file_var).read_text(encoding="utf-8").rstrip("\r\n")
        if value:
            return value
    return _env(env_name)


def _flag(name: str, default: str) -> bool:
    return os.getenv(name, default).strip().lower() in ("1", "true", "yes", "on")


TENANT = check_tenant(_env("MS_TENANT_ID"))
CLIENT_ID = check_client_id(_env("MS_CLIENT_ID"))
MCP_TOKEN = _secret("MCP_TOKEN", "mcp_token")          # shared secret for mcp-remote
CLIENT_SECRET = _secret("MS_CLIENT_SECRET", "client_secret")  # Entra app secret (Web platform)
try:
    VAULT_KEY = load_key(_secret("VAULT_KEY", "vault_key"))
except VaultError as e:
    raise SystemExit(str(e)) from None
# systemd sets STATE_DIRECTORY=/var/lib/sharepoint-mcp (StateDirectory=)
_state_dir = os.getenv("STATE_DIRECTORY", ".").split(":")[0]
VAULT_PATH = os.getenv("VAULT_PATH") or str(Path(_state_dir) / "vault.json.enc")
HOST = os.getenv("HOST", "127.0.0.1")
PORT = int(os.getenv("PORT", "8767"))
ENABLE_WRITE = _flag("ENABLE_WRITE_TOOLS", "true")
ENABLE_DELETE = ENABLE_WRITE and _flag("ENABLE_DELETE_TOOLS", "false")
KEEPALIVE_HOURS = float(os.getenv("KEEPALIVE_HOURS", "12"))
MAX_READ_CHARS = int(os.getenv("MAX_READ_CHARS", "20000"))
MAX_DOWNLOAD = int(float(os.getenv("MAX_DOWNLOAD_MB", "50")) * 2**20)
MAX_UPLOAD = int(float(os.getenv("MAX_UPLOAD_MB", "50")) * 2**20)
SIMPLE_UPLOAD_MAX = 4 * 2**20   # Graph's limit for a single PUT .../content

vault = Vault(VAULT_PATH, VAULT_KEY, vault_context(TENANT, CLIENT_ID))
tokens = TokenManager(vault, TENANT, CLIENT_ID, CLIENT_SECRET)
graph = GraphClient(tokens)

READ_ONLY = {"readOnlyHint": True, "openWorldHint": True}
WRITES = {"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True}
OVERWRITES = {**WRITES, "destructiveHint": True}
DELETES = {"readOnlyHint": False, "destructiveHint": True, "idempotentHint": True, "openWorldHint": True}

mcp = FastMCP(
    "sharepoint",
    instructions=(
        "Access to SharePoint sites, document libraries, files, lists and pages as the signed-in "
        "user (delegated permissions). Find content with search_files (full text, KQL) or browse: "
        "list_sites -> get_site -> list_folder. Every listing prints a ref "
        "('<drive_id>/<item_id>') for each file or folder; pass it back as item= or folder=. "
        "SharePoint URLs and sharing links are accepted wherever a ref is. read_file returns a "
        "file's text (Word, Excel, PowerPoint, PDF, text formats), paged by offset. Lists: "
        "list_lists -> get_list_schema (internal column names) -> list_items. "
        "Read tools return compact text (times in UTC); pass output='json' for structured data. "
        + ("Changes are made as the signed-in user and are visible to everyone with access: "
           "confirm the target and content with the user before writing, moving or overwriting." if ENABLE_WRITE
           else "This server is configured read-only.")
    ),
)

# ---------------------------------------------------------------- helpers
READ = {"annotations": READ_ONLY, "output_schema": None}
REF_RE = re.compile(r"^([A-Za-z0-9!_\-]{3,300})/([A-Za-z0-9!_\-]{3,200})$")
SITE_ID_RE = re.compile(r"^[A-Za-z0-9.\-]+,[0-9a-fA-F\-]{36},[0-9a-fA-F\-]{36}$")
HOST_RE = re.compile(r"^[A-Za-z0-9.\-]{1,253}$")
ITEM_SELECT = ("id,name,size,folder,file,root,package,parentReference,webUrl,lastModifiedDateTime,"
               "lastModifiedBy,createdDateTime,createdBy,description,publication,shared,eTag")
SITE_SELECT = "id,displayName,name,webUrl,description,lastModifiedDateTime"
BAD_NAME_CHARS = set('"*:<>?/\\|')
# URL path segments after which a SharePoint URL no longer names a site.
_NOT_SITE = {"_layouts", "forms", "sitepages", "lists", "_api", "pages", "shared documents",
             "documents", "siteassets", "_vti_bin", "_catalogs"}


def _clamp(n: int, lo: int, hi: int) -> int:
    return max(lo, min(int(n), hi))


async def _call(coro):
    """Run a data function and turn backend errors into tool errors."""
    try:
        return await coro
    except (GraphError, AuthError, extract.ExtractError) as e:
        raise ToolError(str(e)) from e


async def _run(coro, render, output: str = "text") -> str:
    """Run a data function; return compact text (default) or compact JSON."""
    if output not in ("text", "json"):
        coro.close()
        raise ToolError("output must be 'text' or 'json'.")
    data = await _call(coro)
    if output == "json":
        return json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return render(data)


def _fields_arg(fields: Any) -> dict:
    """Column values as a JSON object (some clients send it as a JSON string)."""
    if isinstance(fields, str):
        try:
            fields = json.loads(fields)
        except ValueError:
            raise ToolError('fields must be a JSON object, e.g. {"Title": "Q4", "Status": "Open"}') from None
    if not isinstance(fields, dict) or not fields:
        raise ToolError('fields must be a non-empty JSON object, e.g. {"Title": "Q4"}')
    return fields


def _check_name(name: str, what: str = "name") -> str:
    name = (name or "").strip()
    if not name or len(name) > 250 or any(c in BAD_NAME_CHARS for c in name) or name.endswith("."):
        raise ToolError(f'Invalid {what} {name[:80]!r}: 1-250 characters, none of " * : < > ? / \\ | '
                        f"and not ending with a dot.")
    return name


def _conflict(value: str) -> str:
    if value not in ("fail", "replace", "rename"):
        raise ToolError("conflict must be 'fail', 'replace' or 'rename'.")
    return value


# ---------------------------------------------------------------- resolvers
_sites: dict[str, dict] = {}
_drives: dict[str, list[dict]] = {}
_default_drive: dict[str, dict] = {}
_lists: dict[str, list[dict]] = {}


def _remember_site(key: str, s: dict) -> dict:
    if len(_sites) > 500:
        _sites.clear()
    _sites[key] = s
    return s


async def resolve_site(site: str) -> dict:
    """A site from its URL, id ('host,guid,guid'), 'root', or display name."""
    s = (site or "").strip()
    if not s:
        raise ToolError("Give a site: its URL (https://contoso.sharepoint.com/sites/Finance), "
                        "its id from list_sites, or its name.")
    key = s.lower().rstrip("/")
    if key in _sites:
        return _sites[key]
    if key == "root":
        return _remember_site(key, await graph.get("/sites/root", {"$select": SITE_SELECT}))
    if SITE_ID_RE.match(s):
        return _remember_site(key, await graph.get(f"/sites/{s}", {"$select": SITE_SELECT}))
    if s.startswith("http://") or s.startswith("https://"):
        u = urlparse(s)
        host = u.hostname or ""
        if not HOST_RE.match(host):
            raise ToolError(f"Not a SharePoint URL: {s[:100]!r}")
        segs = [unquote(x) for x in u.path.split("/") if x]
        for i, seg in enumerate(segs):
            if seg.lower() in _NOT_SITE or "." in seg or seg.startswith(":"):
                segs = segs[:i]
                break
        # Longest path first, so a subsite wins over its parent site.
        for n in range(min(len(segs), 6), -1, -1):
            if n == 1 and segs and segs[0].lower() in ("sites", "teams", "personal"):
                continue
            path = "/" + "/".join(segs[:n])
            api = f"/sites/{host}:{quote(path)}" if n else f"/sites/{host}"
            try:
                return _remember_site(key, await graph.get(api, {"$select": SITE_SELECT}))
            except GraphError as e:
                if e.status not in (400, 404):
                    raise
        raise ToolError(f"No site found at {s[:120]!r}, or you don't have access to it.")
    if ":/" in s and HOST_RE.match(s.split(":/", 1)[0]):
        host, path = s.split(":/", 1)
        return _remember_site(key, await graph.get(f"/sites/{host}:/{quote(path)}", {"$select": SITE_SELECT}))
    # A name: search, and accept a single or exact match.
    rows = (await graph.get("/sites", {"search": s, "$select": SITE_SELECT, "$top": 10})).get("value") or []
    exact = [r for r in rows if (r.get("displayName") or "").lower() == key or (r.get("name") or "").lower() == key]
    pick = exact if exact else rows
    if len(pick) == 1:
        return _remember_site(key, pick[0])
    if not pick:
        raise ToolError(f"No site matches {s!r}. Try list_sites(query=...) or pass the site URL.")
    raise ToolError(f"Several sites match {s!r}; pass one of these URLs as site:\n"
                    + fmt.render_sites([fmt.site(r) for r in pick]))


async def site_drives(site_id: str) -> list[dict]:
    if site_id not in _drives:
        rows, _ = await graph.get_all(f"/sites/{site_id}/drives", {
            "$select": "id,name,webUrl,description,driveType,lastModifiedDateTime,quota"}, 200)
        _drives[site_id] = rows
    return _drives[site_id]


async def resolve_drive(site_obj: dict, library: str = "") -> dict:
    sid = site_obj["id"]
    lib = (library or "").strip()
    if not lib:
        if sid not in _default_drive:
            _default_drive[sid] = await graph.get(f"/sites/{sid}/drive", {"$select": "id,name,webUrl"})
        return _default_drive[sid]
    drives = await site_drives(sid)
    low = lib.lower()
    for d in drives:
        tail = unquote((d.get("webUrl") or "").rstrip("/").rsplit("/", 1)[-1]).lower()
        if lib == d.get("id") or (d.get("name") or "").lower() == low or tail == low:
            return d
    names = ", ".join(d.get("name") or "?" for d in drives)
    raise ToolError(f"No document library {lib!r} in {site_obj.get('displayName') or sid}. Libraries: {names}")


def _clean_path(path: str) -> str:
    return "/".join(p for p in (path or "").replace("\\", "/").split("/") if p not in ("", "."))


async def _item_by_path(drive_id: str, path: str, expand: str = "") -> dict:
    p = _clean_path(path)
    api = f"/drives/{drive_id}/root" + (f":/{quote(p)}" if p else "")
    params = {"$select": ITEM_SELECT}
    if expand:
        params["$expand"] = expand
    return await graph.get(api, params)


async def _item_by_ref(drive_id: str, item_id: str, expand: str = "") -> dict:
    params = {"$select": ITEM_SELECT}
    if expand:
        params["$expand"] = expand
    return await graph.get(f"/drives/{drive_id}/items/{item_id}", params)


def _share_token(url: str) -> str:
    return "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")


async def _item_by_url(url: str, expand: str = "") -> dict:
    u = urlparse(url)
    if u.scheme != "https" or not HOST_RE.match(u.hostname or ""):
        raise ToolError(f"Not a SharePoint URL: {url[:100]!r}")
    q = parse_qs(u.query)
    # Library views: .../Forms/AllItems.aspx?id=/sites/x/Shared Documents/Folder
    if q.get("id") and q["id"][0].startswith("/"):
        url = f"https://{u.hostname}{quote(q['id'][0])}"
        u = urlparse(url)
    elif "/forms/" in u.path.lower():
        url = f"https://{u.hostname}{u.path[:u.path.lower().index('/forms/')]}"
        u = urlparse(url)
    is_link = u.path.startswith("/:") or "/_layouts/" in u.path.lower()
    if not is_link:
        # A plain path URL: site + the library whose URL prefixes it + the remainder.
        try:
            site_obj = await resolve_site(url)
            full = unquote(f"https://{u.hostname}{u.path}").rstrip("/").lower()
            for d in await site_drives(site_obj["id"]):
                base = unquote(d.get("webUrl") or "").rstrip("/").lower()
                if base and (full == base or full.startswith(base + "/")):
                    rest = unquote(f"https://{u.hostname}{u.path}").rstrip("/")[len(base):]
                    return await _item_by_path(d["id"], rest, expand)
        except (GraphError, ToolError) as e:
            if isinstance(e, GraphError) and e.status != 404:
                raise
    params = {"$select": ITEM_SELECT}
    if expand:
        params["$expand"] = expand
    try:
        return await graph.get(f"/shares/{_share_token(url)}/driveItem", params)
    except GraphError as e:
        if e.status == 404 or (e.status == 400 and not expand):
            raise ToolError(f"Can't find a file or folder at {url[:150]!r} (or no access). "
                            "Use a ref from list_folder/search_files instead.") from None
        raise


async def resolve_item(item: str = "", site: str = "", library: str = "", path: str = "",
                       expand: str = "", what: str = "item") -> dict:
    """A driveItem from a ref, a URL, or site + library + path."""
    it = (item or "").strip()
    if it:
        if it.startswith("http://") or it.startswith("https://"):
            return await _item_by_url(it, expand)
        m = REF_RE.match(it)
        if not m:
            raise ToolError(f"{what} {it[:80]!r} is neither a ref ('<drive_id>/<item_id>' from "
                            "list_folder or search_files) nor a SharePoint URL.")
        return await _item_by_ref(m.group(1), m.group(2), expand)
    if not (site or "").strip():
        raise ToolError(f"Give {what} (a ref or URL), or site (plus optional library and path).")
    site_obj = await resolve_site(site)
    d = await resolve_drive(site_obj, library)
    return await _item_by_path(d["id"], path, expand)


def _ids(it: dict) -> tuple[str, str]:
    drive_id = (it.get("parentReference") or {}).get("driveId")
    if not drive_id or not it.get("id"):
        raise ToolError("Graph did not return the item's drive; try again with a ref.")
    return drive_id, it["id"]


async def resolve_folder(folder: str, site: str, library: str, path: str) -> dict:
    it = await resolve_item(folder, site, library, path, what="folder")
    if it.get("folder") is None and it.get("root") is None:
        raise ToolError(f"{it.get('name')!r} is a file, not a folder.")
    return it


async def resolve_list(site: str, lst: str) -> tuple[dict, dict]:
    site_obj = await resolve_site(site)
    sid = site_obj["id"]
    name = (lst or "").strip()
    if not name:
        raise ToolError("Give list: its name or id (from list_lists).")
    if sid not in _lists:
        rows, _ = await graph.get_all(f"/sites/{sid}/lists", {
            "$select": "id,displayName,name,list,webUrl,description,lastModifiedDateTime"}, 500)
        _lists[sid] = rows
    low = name.lower()
    for row in _lists[sid]:
        if name == row.get("id") or low in ((row.get("displayName") or "").lower(), (row.get("name") or "").lower()):
            return site_obj, row
    _lists.pop(sid, None)  # maybe created since: re-read next time
    raise ToolError(f"No list {name!r} in {site_obj.get('displayName')}. Use list_lists to see them.")


# ---------------------------------------------------------------- text cache (read_file paging)
_text_cache: OrderedDict[tuple, str] = OrderedDict()


def _cache_put(key: tuple, text: str) -> None:
    _text_cache[key] = text
    _text_cache.move_to_end(key)
    while len(_text_cache) > 8 or sum(len(v) for v in _text_cache.values()) > 40_000_000:
        _text_cache.popitem(last=False)


def _page(text: str, offset: int, max_chars: int) -> tuple[str, dict]:
    total = len(text)
    offset = _clamp(offset, 0, max(total, 0))
    max_chars = _clamp(max_chars or MAX_READ_CHARS, 500, 200_000)
    end = min(offset + max_chars, total)
    if end < total:  # don't cut a line in half if a newline is near
        nl = text.rfind("\n", offset + max_chars // 2, end)
        if nl > 0:
            end = nl + 1
    meta = {"offset": offset, "end": end, "total_chars": total}
    if end < total:
        meta["next_offset"] = end
    return text[offset:end], meta


def _page_note(meta: dict) -> str | None:
    if meta.get("next_offset") is not None:
        return (f"chars {meta['offset']}-{meta['end']} of {meta['total_chars']}: call again with "
                f"offset={meta['next_offset']} for more")
    return None


# ---------------------------------------------------------------- data functions: identity & sites
async def whoami_data() -> dict:
    me = await graph.get("/me", {"$select": "id,displayName,mail,userPrincipalName,jobTitle,department"})
    return {"user": fmt.compact({"name": me.get("displayName"), "email": me.get("mail") or me.get("userPrincipalName"),
                                 "job_title": me.get("jobTitle"), "department": me.get("department")}),
            "auth": {k: v for k, v in tokens.info().items() if v}}


def whoami_text(d: dict) -> str:
    u, a = d["user"], d["auth"]
    return "\n".join([
        f"Signed in as {u.get('name')} <{u.get('email', '?')}>"
        + "".join(f" · {u[k]}" for k in ("job_title", "department") if u.get(k)),
        f"Auth: {a.get('status')}"
        + (f" · signed in {fmt.short_time(a['signed_in_at'])}" if a.get("signed_in_at") else "")
        + (f" · token rotated {fmt.short_time(a['refresh_token_rotated_at'])}" if a.get("refresh_token_rotated_at") else "")
        + (f" · problem: {a['problem']}" if a.get("problem") else ""),
    ])


async def list_sites_data(query: str, limit: int) -> dict:
    q = (query or "").strip()
    limit = _clamp(limit, 1, 200)
    if q:
        rows, more = await graph.get_all("/sites", {"search": q, "$select": SITE_SELECT, "$top": min(limit, 100)}, limit)
        source = f"sites matching {q!r}"
    else:
        rows, more = await graph.get_all("/me/followedSites", {"$select": SITE_SELECT}, limit)
        source = "followed sites"
    # Search also returns personal OneDrive sites; they aren't SharePoint team sites.
    sites = [fmt.site(s) for s in rows if "/personal/" not in (s.get("webUrl") or "")]
    out: dict = {"source": source, "count": len(sites), "sites": sites}
    if more:
        out["note"] = "more sites exist: raise limit or refine query"
    if not q and not sites:
        out["note"] = "you follow no sites: pass query to search all sites you can access (query='*' lists them)"
    return out


def list_sites_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} {d['source']}", fmt.render_sites(d["sites"]), d.get("note"))


async def get_site_data(site: str) -> dict:
    s = await resolve_site(site)
    drives = await site_drives(s["id"])
    return {**fmt.site(s), "libraries": [
        {**fmt.drive(d), "root_ref": fmt.ref(d.get("id"), "root")} for d in drives
        if d.get("driveType") in (None, "documentLibrary")]}


def get_site_text(d: dict) -> str:
    lines = [f"{d.get('name')} · {d.get('url')}", f"site {d.get('site_id')}"]
    if d.get("description"):
        lines.append(d["description"])
    lines.append(f"Document libraries ({len(d['libraries'])}):")
    for lib in d["libraries"]:
        lines.append(f"• {lib.get('name')}"
                     + (f" · {fmt.shorten(lib['description'], 60)}" if lib.get("description") else "")
                     + (f" · {fmt.size(lib['used'])} used" if lib.get("used") else "")
                     + f" · ref {lib.get('root_ref')}")
    lines.append("(browse with list_folder(item=<ref>) or list_folder(site=..., library=<name>); "
                 "lists: list_lists; pages: list_pages)")
    return "\n".join(lines)


# ---------------------------------------------------------------- data functions: files
async def list_folder_data(item: str, site: str, library: str, path: str, limit: int) -> dict:
    folder = await resolve_folder(item, site, library, path)
    drive_id, fid = _ids(folder)
    limit = _clamp(limit, 1, 1000)
    rows, more = await graph.get_all(
        f"/drives/{drive_id}/items/{fid}/children",
        {"$select": ITEM_SELECT, "$top": min(limit, 200)}, limit)
    items = [fmt.item(r) for r in rows]
    items.sort(key=lambda d: (d.get("kind") != "folder", (d.get("name") or "").lower()))
    out: dict = {"folder": fmt.item(folder, full=True), "count": len(items), "items": items}
    if more:
        out["note"] = f"more items exist beyond these {limit}: raise limit, or use search_files with folder"
    return out


def list_folder_text(d: dict) -> str:
    f = d["folder"]
    where = f.get("path") or f.get("name")
    header = f"{where}: {d['count']} items (folders first; times UTC) · folder ref {f.get('ref')}"
    return fmt.with_header(header, fmt.render_items(d["items"]) or "(empty folder)", d.get("note"))


async def get_item_data(item: str, site: str, library: str, path: str) -> dict:
    try:
        it = await resolve_item(item, site, library, path, expand="listItem($expand=fields)")
    except GraphError as e:  # e.g. library root folders have no listItem
        if e.status != 400:
            raise
        it = await resolve_item(item, site, library, path)
    return fmt.item(it, full=True)


def _search_kql(query: str, file_types: str) -> str:
    parts = []
    q = (query or "").strip()
    if q:
        parts.append(f"({q})" if " OR " in q else q)
    types = [t.strip().lstrip(".").lower() for t in (file_types or "").split(",") if t.strip()]
    for t in types:
        if not re.fullmatch(r"[a-z0-9]{1,10}", t):
            raise ToolError(f"file_types should be extensions like 'docx,pdf', got {t!r}")
    if types:
        parts.append("(" + " OR ".join(f"filetype:{t}" for t in types) + ")")
    return " AND ".join(parts)


async def search_files_data(query: str, site: str, folder: str, library: str, file_types: str,
                            sort: str, limit: int, offset: int) -> dict:
    if sort not in ("relevance", "modified"):
        raise ToolError("sort must be 'relevance' or 'modified'.")
    kql = _search_kql(query, file_types)
    scope = None
    if folder.strip():
        f = await resolve_folder(folder, "", "", "")
        scope = f.get("webUrl")
    elif site.strip():
        s = await resolve_site(site)
        scope = (await resolve_drive(s, library))["webUrl"] if library.strip() else s.get("webUrl")
    elif library.strip():
        raise ToolError("library needs site.")
    if not kql:
        if not (scope or sort == "modified"):
            raise ToolError("Give a query (words or KQL), or a site/folder with sort='modified' "
                            "to list recently changed files.")
        kql = "IsDocument:true"
    if scope:
        kql = f'{kql} AND path:"{scope}"'
    limit, offset = _clamp(limit, 1, 100), _clamp(offset, 0, 5000)
    req: dict = {"entityTypes": ["driveItem"], "query": {"queryString": kql}, "from": offset, "size": limit}
    if sort == "modified":
        req["sortProperties"] = [{"name": "lastModifiedDateTime", "isDescending": True}]
    r = await graph.post("/search/query", {"requests": [req]})
    hits, total, more = [], 0, False
    for v in r.get("value") or []:
        for c in v.get("hitsContainers") or []:
            total = max(total, c.get("total") or 0)
            more = more or bool(c.get("moreResultsAvailable"))
            hits.extend(fmt.search_hit(h) for h in c.get("hits") or [])
    out: dict = {"query": kql, "count": len(hits), "total": total, "offset": offset, "hits": hits}
    if more:
        out["note"] = f"more results: call again with offset={offset + len(hits)}"
    return out


def search_files_text(d: dict) -> str:
    rng = f"{d['offset'] + 1}-{d['offset'] + d['count']}" if d["count"] else "0"
    header = f"Results {rng} of ~{d['total']} for {d['query']} (times UTC)"
    return fmt.with_header(header, fmt.render_hits(d["hits"]), d.get("note"))


async def read_file_data(item: str, site: str, library: str, path: str, offset: int,
                         max_chars: int, version: str) -> dict:
    it = await resolve_item(item, site, library, path)
    if it.get("folder") is not None or it.get("root") is not None:
        raise ToolError(f"{it.get('name')!r} is a folder: use list_folder.")
    drive_id, iid = _ids(it)
    name = it.get("name") or ""
    ext = extract.extension(name)
    mime = (it.get("file") or {}).get("mimeType") or ""
    ver = (version or "").strip()
    if ver and not re.fullmatch(r"[0-9.]{1,20}|current", ver):
        raise ToolError("version must be a version id from list_versions, e.g. '3.0'.")
    convert = ext in extract.CONVERT_TO_PDF
    if not (convert or extract.can_extract(name, mime) or (it.get("size") or 0) < 2**20):
        raise ToolError(f"Can't read {name!r} as text (.{ext} files aren't supported). "
                        f"Open it in SharePoint: {it.get('webUrl')}")
    if convert and ver:
        raise ToolError(f"Old versions of .{ext} files can't be converted to text; read the current version.")
    if (it.get("size") or 0) > MAX_DOWNLOAD and not convert:
        raise ToolError(f"{name!r} is {fmt.size(it.get('size'))}, over the {MAX_DOWNLOAD // 2**20} MB "
                        "read limit (MAX_DOWNLOAD_MB).")
    key = (drive_id, iid, it.get("eTag") or it.get("lastModifiedDateTime"), ver)
    text = _text_cache.get(key)
    if text is None:
        if ver:
            data = await graph.download(f"/drives/{drive_id}/items/{iid}/versions/{ver}/content", MAX_DOWNLOAD)
            text = extract.to_text(data, name, mime)
        elif convert:
            data = await graph.download(f"/drives/{drive_id}/items/{iid}/content", MAX_DOWNLOAD,
                                        params={"format": "pdf"})
            text = extract.pdf_text(data)
        else:
            data = await graph.download(f"/drives/{drive_id}/items/{iid}/content", MAX_DOWNLOAD)
            text = extract.to_text(data, name, mime)
        _cache_put(key, text)
    chunk, meta = _page(text, offset, max_chars)
    return {**fmt.item(it, full=False), "url": it.get("webUrl"), "version": ver or None,
            "converted_from": ext if convert else None, **meta, "text": chunk}


def read_file_text(d: dict) -> str:
    head = f"{d.get('name')} · {fmt.size(d.get('size'))}"
    if d.get("modified"):
        head += f" · modified {fmt.short_time(d['modified'])} {d.get('modified_by') or ''}".rstrip()
    if d.get("version"):
        head += f" · version {d['version']}"
    if d.get("converted_from"):
        head += f" · converted from .{d['converted_from']} via PDF"
    head += f" · ref {d.get('ref')}"
    if not d.get("next_offset") and d.get("offset", 0) == 0:
        head += f" · {d['total_chars']} chars"
    return fmt.with_header(head, d["text"].rstrip("\n") or "(no text)", _page_note(d))


async def list_versions_data(item: str, limit: int) -> dict:
    it = await resolve_item(item)
    drive_id, iid = _ids(it)
    rows, more = await graph.get_all(f"/drives/{drive_id}/items/{iid}/versions", None, _clamp(limit, 1, 500))
    return {"name": it.get("name"), "ref": fmt.ref(drive_id, iid), "count": len(rows),
            "versions": [fmt.version(v) for v in rows], "note": "older versions exist" if more else None}


def list_versions_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} versions of {d['name']}, newest first (times UTC; read one with "
                           f"read_file(item, version=...))", fmt.render_versions(d["versions"]), d.get("note"))


# ---------------------------------------------------------------- data functions: lists
async def list_lists_data(site: str, include_hidden: bool) -> dict:
    s = await resolve_site(site)
    rows, _ = await graph.get_all(f"/sites/{s['id']}/lists", {
        "$select": "id,displayName,name,list,webUrl,description,lastModifiedDateTime"}, 500)
    _lists[s["id"]] = rows
    lists = [fmt.list_(r) for r in rows]
    if not include_hidden:
        lists = [x for x in lists if not x.get("hidden")]
    lists.sort(key=lambda x: (x.get("template") == "documentLibrary", (x.get("name") or "").lower()))
    return {"site": s.get("displayName"), "count": len(lists), "lists": lists}


def list_lists_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} lists and libraries in {d['site']}", fmt.render_lists(d["lists"]))


async def get_list_schema_data(site: str, lst: str, include_hidden: bool) -> dict:
    s, row = await resolve_list(site, lst)
    cols, _ = await graph.get_all(f"/sites/{s['id']}/lists/{row['id']}/columns", None, 500)
    if not include_hidden:
        cols = [c for c in cols if not c.get("hidden") and not (c.get("readOnly") and c.get("name") not in (
            "Created", "Modified", "Author", "Editor", "ID"))]
    return {"list": row.get("displayName"), "list_id": row.get("id"), "count": len(cols),
            "columns": [fmt.column(c) for c in cols]}


def get_list_schema_text(d: dict) -> str:
    return fmt.with_header(
        f"{d['count']} columns in {d['list']} (use the internal name before the colon in filters "
        f"and in fields when writing)", fmt.render_columns(d["columns"]))


async def list_items_data(site: str, lst: str, filter: str, order_by: str, fields: str, limit: int) -> dict:
    s, row = await resolve_list(site, lst)
    limit = _clamp(limit, 1, 500)
    keep = {f.strip() for f in (fields or "").split(",") if f.strip()} or None
    for f in keep or ():
        if not re.fullmatch(r"[A-Za-z0-9_]{1,100}", f):
            raise ToolError(f"fields must be internal column names (get_list_schema), got {f!r}")
    expand = f"fields($select={','.join(sorted(keep | {'Title'}))})" if keep else "fields"
    params: dict = {"$expand": expand, "$top": min(limit, 200),
                    "$select": "id,lastModifiedDateTime,lastModifiedBy,webUrl"}
    headers = None
    if filter.strip():
        params["$filter"] = filter.strip()
    if order_by.strip():
        params["$orderby"] = order_by.strip()
    if filter.strip() or order_by.strip():
        # Without an index on the column SharePoint refuses the query unless asked nicely.
        headers = {"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}
    try:
        rows, more = await graph.get_all(f"/sites/{s['id']}/lists/{row['id']}/items", params, limit, headers)
    except GraphError as e:
        if e.status == 400 and (filter.strip() or order_by.strip()):
            raise ToolError(f"{e}\nFilter and order_by take OData on fields, e.g. "
                            "filter=\"fields/Status eq 'Open'\", order_by='fields/Modified desc'. "
                            "Check the internal column names with get_list_schema.") from None
        raise
    items = [fmt.list_item(r, keep | {"Title"} if keep else None) for r in rows]
    out: dict = {"list": row.get("displayName"), "count": len(items), "items": items}
    if more:
        out["note"] = f"more items exist beyond these {limit}: raise limit or add a filter"
    return out


def list_items_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} items in {d['list']} (times UTC)",
                           fmt.render_list_items(d["items"]), d.get("note"))


def _list_item_id(item_id: str) -> str:
    iid = str(item_id or "").strip().lstrip("#")
    if not iid.isdigit():
        raise ToolError("item_id must be the numeric list item id (#N in list_items).")
    return iid


async def get_list_item_data(site: str, lst: str, item_id: str) -> dict:
    s, row = await resolve_list(site, lst)
    li = await graph.get(f"/sites/{s['id']}/lists/{row['id']}/items/{_list_item_id(item_id)}",
                         {"$expand": "fields"})
    d = fmt.list_item(li)
    d["created"] = li.get("createdDateTime")
    d["created_by"] = fmt.who(li.get("createdBy"))
    d["list"] = row.get("displayName")
    return fmt.compact(d)


def get_list_item_text(d: dict) -> str:
    lines = [f"{d['list']} #{d.get('item_id')}"]
    for k, v in (d.get("fields") or {}).items():
        lines.append(f"{k}: {fmt._value(v, 2000)}")
    if d.get("created"):
        lines.append(f"Created: {fmt.short_time(d['created'])} {d.get('created_by') or ''}".rstrip())
    if d.get("modified"):
        lines.append(f"Modified: {fmt.short_time(d['modified'])} {d.get('modified_by') or ''}".rstrip())
    if d.get("url"):
        lines.append(f"Open: {d['url']}")
    return "\n".join(lines)


# ---------------------------------------------------------------- data functions: pages
PAGE_SELECT = "id,name,title,webUrl,description,lastModifiedDateTime,lastModifiedBy,promotionKind"


async def list_pages_data(site: str, limit: int) -> dict:
    s = await resolve_site(site)
    rows, more = await graph.get_all(f"/sites/{s['id']}/pages/microsoft.graph.sitePage",
                                     {"$select": PAGE_SELECT}, _clamp(limit, 1, 500))
    rows.sort(key=lambda p: p.get("lastModifiedDateTime") or "", reverse=True)
    out: dict = {"site": s.get("displayName"), "count": len(rows), "pages": [fmt.page(p) for p in rows]}
    if more:
        out["note"] = "more pages exist: raise limit"
    return out


def list_pages_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} pages in {d['site']}, most recently modified first (times UTC)",
                           fmt.render_pages(d["pages"]), d.get("note"))


def _canvas_text(layout: dict) -> str:
    out: list[str] = []

    def webparts(parts: list[dict]):
        for wp in parts or []:
            if wp.get("innerHtml"):
                out.append(fmt.html_to_text(wp["innerHtml"]))
                continue
            data = wp.get("data") or {}
            title = data.get("title") or data.get("description")
            texts = [t.get("value") for t in ((data.get("serverProcessedContent") or {})
                                              .get("searchablePlainTexts") or []) if t.get("value")]
            if title or texts:
                out.append(" · ".join(x for x in [f"[{title}]" if title else None, *texts] if x))

    for sec in layout.get("horizontalSections") or []:
        for col in sec.get("columns") or []:
            webparts(col.get("webparts"))
    vs = layout.get("verticalSection") or {}
    webparts(vs.get("webparts"))
    return "\n".join(x for x in out if x)


async def read_page_data(site: str, page: str, offset: int, max_chars: int) -> dict:
    p = (page or "").strip()
    if p.startswith("https://"):
        site = site or p
        p = unquote(urlparse(p).path.rstrip("/").rsplit("/", 1)[-1])
    s = await resolve_site(site)
    if not GUID_RE.match(p):
        rows, _ = await graph.get_all(f"/sites/{s['id']}/pages/microsoft.graph.sitePage",
                                      {"$select": "id,name,title"}, 2000)
        low = p.lower()
        match = [r for r in rows if low in ((r.get("name") or "").lower(), (r.get("title") or "").lower(),
                                            (r.get("name") or "").lower().removesuffix(".aspx"))]
        if not match:
            raise ToolError(f"No page {p!r} in {s.get('displayName')}: use list_pages.")
        p = match[0]["id"]
    pg = await graph.get(f"/sites/{s['id']}/pages/{p}/microsoft.graph.sitePage", {"$expand": "canvasLayout"})
    text = _canvas_text(pg.get("canvasLayout") or {})
    if pg.get("description"):
        text = pg["description"] + "\n\n" + text
    chunk, meta = _page(text, offset, max_chars)
    return {**fmt.page(pg), **meta, "text": chunk}


def read_page_text(d: dict) -> str:
    head = f"{d.get('title')} · {fmt._mod(d)} · {d.get('url')}"
    return fmt.with_header(head, d["text"].rstrip("\n") or "(no text content)", _page_note(d))


# ---------------------------------------------------------------- write functions
_BINARY_EXT = extract.OOXML_EXTENSIONS | extract.CONVERT_TO_PDF | {
    "pdf", "png", "jpg", "jpeg", "gif", "bmp", "zip", "msg", "vsdx", "mpp", "one", "mp4", "mp3"}


def _content_bytes(content: str, encoding: str, name: str) -> bytes:
    if encoding == "text":
        if extract.extension(name) in _BINARY_EXT:
            raise ToolError(f"{name!r} is a binary format: build the file and pass it with "
                            "encoding='base64' (sending text would corrupt it).")
        data = (content or "").encode("utf-8")
    elif encoding == "base64":
        try:
            data = base64.b64decode(re.sub(r"\s+", "", content or ""), validate=True)
        except (binascii.Error, ValueError):
            raise ToolError("content is not valid base64.") from None
    else:
        raise ToolError("encoding must be 'text' or 'base64'.")
    if len(data) > MAX_UPLOAD:
        raise ToolError(f"The file is {fmt.size(len(data))}, over the {MAX_UPLOAD // 2**20} MB "
                        "upload limit (MAX_UPLOAD_MB).")
    return data


async def _put_content(api_base: str, data: bytes, name: str, conflict: str | None) -> dict:
    """Upload to '/drives/{d}/items/{parent}:/{name}:' (new) or '/drives/{d}/items/{id}' (replace)."""
    ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
    if len(data) <= SIMPLE_UPLOAD_MAX:
        params = {"@microsoft.graph.conflictBehavior": conflict} if conflict else None
        return await graph.request("PUT", f"{api_base}/content", params=params, content=data,
                                   headers={"Content-Type": ctype})
    body = {"item": {"@microsoft.graph.conflictBehavior": conflict}} if conflict else {}
    session = await graph.post(f"{api_base}/createUploadSession", body)
    return await graph.upload_session(session.get("uploadUrl") or "", data)


def _written(verb: str, it: dict) -> str:
    d = fmt.item(it, full=True)
    return (f"{verb} {d.get('path') or d.get('name')} ({fmt.size(d.get('size'))}) · ref {d.get('ref')}"
            + (f" · {d['url']}" if d.get("url") else ""))


async def write_file_data(content: str, name: str, item: str, folder: str, site: str, library: str,
                          path: str, encoding: str, conflict: str) -> str:
    if item.strip():  # replace an existing file's content (new version)
        target = await resolve_item(item)
        if target.get("folder") is not None or target.get("root") is not None:
            raise ToolError("item is a folder: pass it as folder= with a name to create a file in it.")
        drive_id, iid = _ids(target)
        data = _content_bytes(content, encoding, target.get("name") or "")
        it = await _put_content(f"/drives/{drive_id}/items/{iid}", data, target.get("name") or "", None)
        _text_cache.clear()
        return _written("Replaced the content of", it or target)
    fname = _check_name(name)
    data = _content_bytes(content, encoding, fname)
    parent = await resolve_folder(folder, site, library, path)
    drive_id, pid = _ids(parent)
    try:
        it = await _put_content(f"/drives/{drive_id}/items/{pid}:/{quote(fname)}:", data, fname, _conflict(conflict))
    except GraphError as e:
        if e.status == 409:
            raise ToolError(f"{fname!r} already exists there. Pass conflict='replace' to overwrite it "
                            "(SharePoint keeps the old version) or conflict='rename'.") from None
        raise
    return _written("Uploaded", it)


async def create_folder_data(name: str, folder: str, site: str, library: str, path: str, conflict: str) -> str:
    fname = _check_name(name, "folder name")
    parent = await resolve_folder(folder, site, library, path)
    drive_id, pid = _ids(parent)
    behavior = _conflict(conflict)
    try:
        it = await graph.post(f"/drives/{drive_id}/items/{pid}/children",
                              {"name": fname, "folder": {}, "@microsoft.graph.conflictBehavior": behavior})
    except GraphError as e:
        if e.status == 409:
            raise ToolError(f"A folder or file named {fname!r} already exists there.") from None
        raise
    d = fmt.item(it, full=True)
    return f"Created folder {d.get('path') or d.get('name')} · ref {d.get('ref')}"


async def move_item_data(item: str, destination: str, new_name: str) -> str:
    it = await resolve_item(item)
    drive_id, iid = _ids(it)
    body: dict = {}
    if new_name.strip():
        body["name"] = _check_name(new_name)
    if destination.strip():
        dest = await resolve_folder(destination, "", "", "")
        dd, did = _ids(dest)
        if dd != drive_id:
            raise ToolError("SharePoint can't move items between libraries or sites in one step: "
                            "use copy_item to the destination, then delete the original if wanted.")
        body["parentReference"] = {"id": did}
    if not body:
        raise ToolError("Give destination (a folder ref or URL) and/or new_name.")
    try:
        moved = await graph.patch(f"/drives/{drive_id}/items/{iid}", body)
    except GraphError as e:
        if e.status == 409:
            raise ToolError("An item with that name already exists in the destination.") from None
        raise
    return _written("Moved/renamed to", moved)


async def copy_item_data(item: str, destination: str, new_name: str, conflict: str) -> str:
    it = await resolve_item(item)
    drive_id, iid = _ids(it)
    dest = await resolve_folder(destination, "", "", "")
    dd, did = _ids(dest)
    body: dict = {"parentReference": {"driveId": dd, "id": did}}
    if new_name.strip():
        body["name"] = _check_name(new_name)
    r = await graph.send("POST", f"/drives/{drive_id}/items/{iid}/copy",
                         params={"@microsoft.graph.conflictBehavior": _conflict(conflict)}, json=body)
    monitor = r.headers.get("Location")
    target = f"{fmt.item(dest, full=True).get('path') or dest.get('name')}"
    if not monitor:
        return f"Copy of {it.get('name')} to {target} accepted."
    status: dict = {}
    for delay in (1, 1, 2, 2, 3, 3, 4, 4, 5, 5):
        await asyncio.sleep(delay)
        status = await graph.monitor(monitor)
        if status.get("status") in ("completed", "failed"):
            break
    st = status.get("status")
    if st == "completed":
        rid = status.get("resourceId")
        return f"Copied {it.get('name')} to {target}" + (f" · ref {dd}/{rid}" if rid else "")
    if st == "failed":
        err = (status.get("error") or {}).get("message") or status
        raise ToolError(f"Copy failed: {err}")
    pct = status.get("percentageComplete")
    return (f"Copy of {it.get('name')} to {target} is still running"
            + (f" ({pct:.0f}%)" if isinstance(pct, (int, float)) else "")
            + ": check the destination with list_folder in a minute.")


async def update_file_fields_data(item: str, fields: Any) -> str:
    values = _fields_arg(fields)
    it = await resolve_item(item)
    drive_id, iid = _ids(it)
    try:
        out = await graph.patch(f"/drives/{drive_id}/items/{iid}/listItem/fields", values)
    except GraphError as e:
        if e.status == 400:
            raise ToolError(f"{e}\nUse internal column names from get_list_schema on the library.") from None
        raise
    shown = fmt.clean_fields(out, set(values))
    return f"Updated {it.get('name')}: " + " · ".join(f"{k}: {fmt._value(v, 80)}" for k, v in shown.items())


async def create_list_item_data(site: str, lst: str, fields: Any) -> str:
    values = _fields_arg(fields)
    s, row = await resolve_list(site, lst)
    try:
        li = await graph.post(f"/sites/{s['id']}/lists/{row['id']}/items", {"fields": values})
    except GraphError as e:
        if e.status == 400:
            raise ToolError(f"{e}\nCheck the internal column names and types with get_list_schema.") from None
        raise
    return f"Created item #{li.get('id')} in {row.get('displayName')}" + (
        f" · {li['webUrl']}" if li.get("webUrl") else "")


async def update_list_item_data(site: str, lst: str, item_id: str, fields: Any) -> str:
    values = _fields_arg(fields)
    s, row = await resolve_list(site, lst)
    iid = _list_item_id(item_id)
    try:
        out = await graph.patch(f"/sites/{s['id']}/lists/{row['id']}/items/{iid}/fields", values)
    except GraphError as e:
        if e.status == 400:
            raise ToolError(f"{e}\nCheck the internal column names and types with get_list_schema.") from None
        raise
    shown = fmt.clean_fields(out, set(values))
    return f"Updated {row.get('displayName')} #{iid}: " + " · ".join(
        f"{k}: {fmt._value(v, 80)}" for k, v in shown.items())


async def delete_item_data(item: str) -> str:
    it = await resolve_item(item)
    if it.get("root") is not None:
        raise ToolError("Refusing to delete a library root.")
    drive_id, iid = _ids(it)
    await graph.delete(f"/drives/{drive_id}/items/{iid}")
    _text_cache.clear()
    kind = "folder" if it.get("folder") is not None else "file"
    return f"Deleted {kind} {fmt.item_path(it) or it.get('name')} (it is in the site's recycle bin)."


async def delete_list_item_data(site: str, lst: str, item_id: str) -> str:
    s, row = await resolve_list(site, lst)
    iid = _list_item_id(item_id)
    await graph.delete(f"/sites/{s['id']}/lists/{row['id']}/items/{iid}")
    return f"Deleted {row.get('displayName')} #{iid} (it is in the site's recycle bin)."


# ---------------------------------------------------------------- read tools
@mcp.tool(**READ)
async def whoami(output: str = "text") -> str:
    """The Microsoft 365 account this server acts as, and the state of its sign-in."""
    return await _run(whoami_data(), whoami_text, output)


@mcp.tool(**READ)
async def list_sites(query: str = "", limit: int = 50, output: str = "text") -> str:
    """SharePoint sites. Without query: the sites the user follows. With query:
    sites whose name or description match (query='*' lists every site the user
    can access). Each line gives the URL and site id to use as site= elsewhere."""
    return await _run(list_sites_data(query, limit), list_sites_text, output)


@mcp.tool(**READ)
async def get_site(site: str, output: str = "text") -> str:
    """A site's details and its document libraries, each with the ref of its root
    folder. site is a URL (https://contoso.sharepoint.com/sites/Finance), a site id
    from list_sites, 'root', or a site name."""
    return await _run(get_site_data(site), get_site_text, output)


@mcp.tool(**READ)
async def list_folder(item: str = "", site: str = "", library: str = "", path: str = "",
                      limit: int = 200, output: str = "text") -> str:
    """Contents of a folder, folders first: name, size, last modified and a ref per
    entry. Address the folder by item (a folder ref or SharePoint URL) or by site +
    library (name; default the site's main library) + path ('Reports/2026';
    empty = library root)."""
    return await _run(list_folder_data(item, site, library, path, limit), list_folder_text, output)


@mcp.tool(**READ)
async def get_item(item: str = "", site: str = "", library: str = "", path: str = "",
                   output: str = "text") -> str:
    """Details of one file or folder: path, type, size, created/modified by, version,
    check-out state, the library's custom column values, and links."""
    return await _run(get_item_data(item, site, library, path), fmt.render_item, output)


@mcp.tool(**READ)
async def search_files(query: str = "", site: str = "", folder: str = "", library: str = "",
                       file_types: str = "", sort: str = "relevance", limit: int = 25, offset: int = 0,
                       output: str = "text") -> str:
    """Full-text search over files the user can access in SharePoint (and OneDrive),
    matching names, content and metadata (Microsoft Search; KQL allowed, e.g.
    'budget AND author:"Alice"' or 'title:"Q4 forecast"'). Narrow it with site
    (URL/id/name; plus library), folder (ref or URL), file_types ('docx,pdf,xlsx').
    sort='modified' puts the newest first; with an empty query and a site or folder
    it lists the most recently changed files. Page with offset. Each hit gives a
    ref for read_file. The index can lag new files by a few minutes."""
    return await _run(search_files_data(query, site, folder, library, file_types, sort, limit, offset),
                      search_files_text, output)


@mcp.tool(**READ)
async def read_file(item: str = "", site: str = "", library: str = "", path: str = "",
                    offset: int = 0, max_chars: int = 0, version: str = "", output: str = "text") -> str:
    """The text of a file: Word, PowerPoint (with notes), Excel (every sheet,
    tab-separated with row numbers), PDF, legacy Office/OpenDocument (converted via
    PDF), and text formats (txt, md, csv, json, html, code...). Long files come in
    pages of max_chars (default 20000): the footer gives the offset for the next
    call. version reads an older version (ids from list_versions)."""
    return await _run(read_file_data(item, site, library, path, offset, max_chars, version),
                      read_file_text, output)


@mcp.tool(**READ)
async def list_versions(item: str, limit: int = 50, output: str = "text") -> str:
    """Version history of a file (ref or URL), newest first: version id, time, author, size."""
    return await _run(list_versions_data(item, limit), list_versions_text, output)


@mcp.tool(**READ)
async def list_lists(site: str, include_hidden: bool = False, output: str = "text") -> str:
    """Lists and document libraries in a site, with their template (genericList,
    documentLibrary, events, ...) and id."""
    return await _run(list_lists_data(site, include_hidden), list_lists_text, output)


@mcp.tool(**READ)
async def get_list_schema(site: str, list: str, include_hidden: bool = False, output: str = "text") -> str:
    """Columns of a list or library: internal name, display name, type, choices,
    required. Use the internal names in list_items filters and when writing fields."""
    return await _run(get_list_schema_data(site, list, include_hidden), get_list_schema_text, output)


@mcp.tool(**READ)
async def list_items(site: str, list: str, filter: str = "", order_by: str = "", fields: str = "",
                     limit: int = 50, output: str = "text") -> str:
    """Items of a list (list = name or id). filter and order_by are OData on
    fields, e.g. filter="fields/Status eq 'Open' and fields/Priority eq 'High'",
    order_by='fields/Modified desc'. fields = comma-separated internal column names
    to return (default: all non-system columns). limit is 1-500."""
    return await _run(list_items_data(site, list, filter, order_by, fields, limit), list_items_text, output)


@mcp.tool(**READ)
async def get_list_item(site: str, list: str, item_id: str, output: str = "text") -> str:
    """One list item with all its column values in full."""
    return await _run(get_list_item_data(site, list, item_id), get_list_item_text, output)


@mcp.tool(**READ)
async def list_pages(site: str, limit: int = 100, output: str = "text") -> str:
    """Site pages and news posts in a site, most recently modified first."""
    return await _run(list_pages_data(site, limit), list_pages_text, output)


@mcp.tool(**READ)
async def read_page(site: str = "", page: str = "", offset: int = 0, max_chars: int = 0,
                    output: str = "text") -> str:
    """The text content of a site page or news post. page is its id, its name
    ('Home.aspx') or title, or the page URL (then site may be omitted)."""
    return await _run(read_page_data(site, page, offset, max_chars), read_page_text, output)


# ---------------------------------------------------------------- write tools
async def write_file(content: str, name: str = "", folder: str = "", site: str = "", library: str = "",
                     path: str = "", item: str = "", encoding: str = "text", conflict: str = "fail") -> str:
    """Create a file, or replace a file's content. To create: give name and the
    target folder (folder = ref or URL, or site + library + path). To overwrite an
    existing file: give item (its ref) instead; SharePoint keeps the previous
    version. encoding='text' for text formats (txt, md, csv, json, html...);
    binary files (docx, xlsx, pdf, images) must be sent with encoding='base64'.
    conflict when the name exists: 'fail' (default), 'replace' or 'rename'.
    Confirm the destination and content with the user first."""
    return await _call(write_file_data(content, name, item, folder, site, library, path, encoding, conflict))


async def create_folder(name: str, folder: str = "", site: str = "", library: str = "", path: str = "",
                        conflict: str = "fail") -> str:
    """Create a folder inside a folder (folder = ref or URL, or site + library + path)."""
    return await _call(create_folder_data(name, folder, site, library, path, conflict))


async def move_item(item: str, destination: str = "", new_name: str = "") -> str:
    """Move a file or folder to another folder in the same library (destination =
    folder ref or URL) and/or rename it (new_name). Across libraries or sites, use
    copy_item. Confirm with the user first."""
    return await _call(move_item_data(item, destination, new_name))


async def copy_item(item: str, destination: str, new_name: str = "", conflict: str = "fail") -> str:
    """Copy a file or folder to a folder anywhere the user can write, including
    other libraries and sites (destination = folder ref or URL). Large copies
    continue in the background."""
    return await _call(copy_item_data(item, destination, new_name, conflict))


async def update_file_fields(item: str, fields: dict[str, Any] | str) -> str:
    """Set a file's or folder's library column values (metadata), e.g.
    {"Status": "Approved", "Department": "Finance"}. Use internal column names
    from get_list_schema(site, library name)."""
    return await _call(update_file_fields_data(item, fields))


async def create_list_item(site: str, list: str, fields: dict[str, Any] | str) -> str:
    """Add an item to a list. fields maps internal column names (get_list_schema)
    to values, e.g. {"Title": "New request", "Priority": "High"}. Person columns
    take '<Name>LookupId' with the user's site id; multi-choice columns need
    {"Col@odata.type": "Collection(Edm.String)", "Col": ["a","b"]}."""
    return await _call(create_list_item_data(site, list, fields))


async def update_list_item(site: str, list: str, item_id: str, fields: dict[str, Any] | str) -> str:
    """Change column values of a list item (#id from list_items). Only the given
    fields change."""
    return await _call(update_list_item_data(site, list, item_id, fields))


async def delete_item(item: str) -> str:
    """Delete a file or folder (with everything in it). It goes to the site's
    recycle bin. Always confirm with the user first, naming the exact item."""
    return await _call(delete_item_data(item))


async def delete_list_item(site: str, list: str, item_id: str) -> str:
    """Delete a list item (to the recycle bin). Confirm with the user first."""
    return await _call(delete_list_item_data(site, list, item_id))


if ENABLE_WRITE:
    for _fn, _ann in ((write_file, OVERWRITES), (create_folder, WRITES), (move_item, WRITES),
                      (copy_item, WRITES), (update_file_fields, WRITES), (create_list_item, WRITES),
                      (update_list_item, WRITES)):
        mcp.tool(annotations=_ann, output_schema=None)(_fn)
if ENABLE_DELETE:
    for _fn in (delete_item, delete_list_item):
        mcp.tool(annotations=DELETES, output_schema=None)(_fn)


# ---------------------------------------------------------------- http + auth
@mcp.custom_route("/health", methods=["GET"])
async def health(_: Request):
    return JSONResponse({"status": "ok", "auth": tokens.status()})


class BearerAuth:
    """Minimal ASGI middleware: every request except /health needs the token."""

    def __init__(self, app, token: str):
        self.app = app
        self.expected = f"Bearer {token}".encode()

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http" and scope.get("path") != "/health":
            supplied = dict(scope.get("headers") or []).get(b"authorization", b"")
            if not hmac.compare_digest(supplied, self.expected):
                await send({"type": "http.response.start", "status": 401,
                            "headers": [(b"content-type", b"text/plain")]})
                await send({"type": "http.response.body", "body": b"Unauthorized"})
                return
        await self.app(scope, receive, send)


def build_app():
    return BearerAuth(mcp.http_app(path="/mcp"), MCP_TOKEN)


async def _keepalive_loop() -> None:
    """Refresh periodically so the stored refresh token keeps rotating even when
    the server is idle; otherwise it would expire after ~90 days of no use."""
    if KEEPALIVE_HOURS <= 0:
        return
    await asyncio.sleep(60)
    while True:
        await tokens.keepalive()
        await asyncio.sleep(KEEPALIVE_HOURS * 3600)


async def main() -> None:
    log.info("Starting on %s:%d (write tools %s, delete tools %s, vault %s, sign-in %s)", HOST, PORT,
             "enabled" if ENABLE_WRITE else "disabled", "enabled" if ENABLE_DELETE else "disabled",
             VAULT_PATH, tokens.status())
    server = uvicorn.Server(uvicorn.Config(build_app(), host=HOST, port=PORT, proxy_headers=True))
    keepalive = asyncio.create_task(_keepalive_loop())
    try:
        await server.serve()
    finally:
        keepalive.cancel()
        await graph.aclose()
        await tokens.aclose()


if __name__ == "__main__":
    asyncio.run(main())
