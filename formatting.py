"""Turn Graph SharePoint objects into compact output for the model.

Two layers, as in msteams-mcp:
- normalisers (site, drive, item, list_, column, list_item, ...) turn Graph JSON
  into small dicts with only what is useful; these are also the output="json" form;
- renderers (render_*) turn those dicts into compact text, the default output.

Files and folders are identified by an item ref "<drive_id>/<item_id>", which
every tool that takes `item` or `folder` accepts back.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from html.parser import HTMLParser
from urllib.parse import unquote

_BLOCK = {"p", "div", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "ul", "ol",
          "table", "section", "article", "header", "footer"}
_SKIP = {"script", "style", "head", "noscript"}


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in _SKIP:
            self.skip += 1
            return
        a = dict(attrs)
        if tag == "br":
            self.out.append("\n")
        elif tag == "li":
            self.out.append("\n- ")
        elif tag == "img":
            self.out.append(f"[{a.get('alt') or 'image'}]")
        elif tag in ("td", "th"):
            self.out.append(" | ")
        elif tag in ("h1", "h2", "h3", "h4"):
            self.out.append("\n" + "#" * int(tag[1]) + " ")
        elif tag in _BLOCK:
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in _SKIP:
            self.skip = max(0, self.skip - 1)
        elif tag in _BLOCK:
            self.out.append("\n")

    def handle_data(self, data):
        if not self.skip:
            self.out.append(data)


def html_to_text(content: str) -> str:
    p = _TextExtractor()
    p.feed(content or "")
    p.close()
    return tidy("".join(p.out))


_WS = " \t\u00a0\u2007\u202f\u200b\u2060\ufeff"


def tidy(text: str) -> str:
    """Trim every line and drop blank lines."""
    lines = (ln.strip(_WS) for ln in (text or "").replace("\r\n", "\n").split("\n"))
    return "\n".join(ln for ln in lines if ln)


def strip_search_markup(s: str) -> str:
    """Search summaries wrap hits in <c0>..</c0> and use <ddd/> for ellipses."""
    return " ".join(html_to_text((s or "").replace("<ddd/>", "…").replace("<c0>", "")
                                 .replace("</c0>", "")).split())


def parse_time(value: str) -> datetime:
    s = value.strip().replace("Z", "+00:00").replace("z", "+00:00")
    m = re.match(r"^(.*T\d\d:\d\d:\d\d)(\.\d+)?(.*)$", s)
    if m and m.group(2):
        s = m.group(1) + m.group(2)[:7] + m.group(3)
    dt = datetime.fromisoformat(s)
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def utc(value: str | None) -> datetime | None:
    try:
        return parse_time(value).astimezone(timezone.utc) if value else None
    except ValueError:
        return None


def short_time(value: str | None) -> str:
    """'2026-11-13T10:03:21Z' -> '2026-11-13 10:03' (UTC)."""
    t = utc(value)
    return t.strftime("%Y-%m-%d %H:%M") if t else (value or "?")


def shorten(text: str, n: int) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= n else text[: n - 1].rstrip() + "…"


def size(n: int | None) -> str:
    if n is None:
        return "?"
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return "?"


def who(identity: dict | None) -> str | None:
    """Display name from a Graph identitySet."""
    if not identity:
        return None
    for k in ("user", "application", "device", "group"):
        v = identity.get(k)
        if v:
            return v.get("displayName") or v.get("email") or v.get("id")
    return None


def compact(d: dict) -> dict:
    return {k: v for k, v in d.items() if v not in (None, "", [], {})}


def ref(drive_id: str | None, item_id: str | None) -> str | None:
    return f"{drive_id}/{item_id}" if drive_id and item_id else None


# ---------------------------------------------------------------- normalisers
def site(s: dict) -> dict:
    return compact({
        "site_id": s.get("id"),
        "name": s.get("displayName") or s.get("name"),
        "url": s.get("webUrl"),
        "description": s.get("description"),
        "modified": s.get("lastModifiedDateTime"),
    })


def drive(d: dict) -> dict:
    quota = d.get("quota") or {}
    return compact({
        "drive_id": d.get("id"),
        "name": d.get("name"),
        "url": d.get("webUrl"),
        "description": d.get("description"),
        "modified": d.get("lastModifiedDateTime"),
        "used": quota.get("used"),
    })


def item_path(it: dict) -> str | None:
    """Path of an item inside its library: 'Reports/2026/Q4.xlsx'."""
    parent = (it.get("parentReference") or {}).get("path") or ""
    # parentReference.path = '/drives/{id}/root:/Reports/2026' (or '/drive/root:')
    if "root:" in parent:
        base = unquote(parent.split("root:", 1)[1]).strip("/")
        return f"{base}/{it.get('name')}" if base else it.get("name")
    if it.get("root") is not None:
        return "/"
    return None


def item(it: dict, full: bool = False) -> dict:
    pr = it.get("parentReference") or {}
    drive_id = pr.get("driveId")
    folder = it.get("folder")
    out: dict = {
        "ref": ref(drive_id, it.get("id")),
        "name": it.get("name") or ("(library root)" if it.get("root") is not None else None),
        "kind": "folder" if folder is not None or it.get("root") is not None else "file",
        "size": it.get("size"),
        "modified": it.get("lastModifiedDateTime"),
        "modified_by": who(it.get("lastModifiedBy")),
    }
    if folder is not None:
        out["children"] = folder.get("childCount")
    if it.get("package"):
        out["package"] = (it["package"] or {}).get("type")   # e.g. oneNote
    if full:
        f = it.get("file") or {}
        out.update({
            "path": item_path(it),
            "url": it.get("webUrl"),
            "mime": f.get("mimeType"),
            "created": it.get("createdDateTime"),
            "created_by": who(it.get("createdBy")),
            "description": it.get("description"),
            "site_id": pr.get("siteId"),
            "parent_ref": ref(drive_id, pr.get("id")),
            "etag": it.get("eTag"),
        })
        if it.get("publication"):
            out["version"] = (it["publication"] or {}).get("versionId")
            out["level"] = (it["publication"] or {}).get("level")  # published / checkout
        if (it.get("shared") or {}).get("scope"):
            out["shared"] = it["shared"]["scope"]
        li = it.get("listItem") or {}
        fields = clean_fields(li.get("fields") or {}, drop_doc_fields=True)
        if fields:
            out["fields"] = fields
        if li.get("id"):
            out["list_item_id"] = li["id"]
    return compact(out)


def search_hit(h: dict) -> dict:
    r = h.get("resource") or {}
    pr = r.get("parentReference") or {}
    return compact({
        "ref": ref(pr.get("driveId"), r.get("id")),
        "name": r.get("name"),
        "kind": "folder" if r.get("folder") is not None else "file",
        "size": r.get("size"),
        "modified": r.get("lastModifiedDateTime"),
        "modified_by": who(r.get("lastModifiedBy")),
        "url": r.get("webUrl"),
        "summary": strip_search_markup(h.get("summary") or ""),
    })


def list_(lst: dict) -> dict:
    info = lst.get("list") or {}
    return compact({
        "list_id": lst.get("id"),
        "name": lst.get("displayName") or lst.get("name"),
        "template": info.get("template"),
        "hidden": info.get("hidden") or None,
        "url": lst.get("webUrl"),
        "description": lst.get("description"),
        "modified": lst.get("lastModifiedDateTime"),
    })


_COLUMN_TYPES = ("text", "number", "choice", "dateTime", "boolean", "lookup", "personOrGroup",
                 "currency", "calculated", "hyperlinkOrPicture", "term", "contentApprovalStatus",
                 "thumbnail", "geolocation")


def column(c: dict) -> dict:
    ctype = next((t for t in _COLUMN_TYPES if c.get(t) is not None), None)
    extra = None
    if ctype == "choice":
        extra = "choices: " + ", ".join((c["choice"] or {}).get("choices") or [])
        if (c["choice"] or {}).get("allowTextEntry"):
            extra += " (free text allowed)"
    elif ctype == "lookup":
        extra = f"lookup to list {(c['lookup'] or {}).get('listId')} column {(c['lookup'] or {}).get('columnName')}"
    elif ctype == "text" and (c["text"] or {}).get("allowMultipleLines"):
        extra = "multi-line"
    elif ctype == "personOrGroup" and (c["personOrGroup"] or {}).get("allowMultipleSelection"):
        extra = "multiple"
    elif ctype == "dateTime":
        extra = (c["dateTime"] or {}).get("format")
    return compact({
        "name": c.get("name"),
        "display_name": c.get("displayName") if c.get("displayName") != c.get("name") else None,
        "type": ctype,
        "required": c.get("required") or None,
        "read_only": c.get("readOnly") or None,
        "indexed": c.get("indexed") or None,
        "details": extra,
        "description": c.get("description"),
    })


# Internal SharePoint fields that are noise to an agent.
_SYSTEM_FIELDS = {
    "ContentType", "Edit", "LinkTitleNoMenu", "LinkTitle", "ItemChildCount", "FolderChildCount",
    "Attachments", "AppAuthorLookupId", "AppEditorLookupId", "DocIcon", "LinkFilenameNoMenu",
    "LinkFilename", "FileSizeDisplay", "SelectTitle", "ContentTypeId", "AuthorLookupId",
    "EditorLookupId", "Created", "Modified", "id", "ParentVersionStringLookupId",
    "ParentLeafNameLookupId", "CheckoutUserLookupId", "MediaServiceImageTags",
}
_DOC_FIELDS = {"FileLeafRef", "FileRef", "FileDirRef", "FileSizeDisplay", "FSObjType", "Created_x0020_Date",
               "Last_x0020_Modified", "File_x0020_Type", "HTML_x0020_File_x0020_Type", "_CheckinComment",
               "LinkCheckedOutTitle", "FileSize"}


def clean_fields(fields: dict, keep: set[str] | None = None, drop_doc_fields: bool = False) -> dict:
    out = {}
    for k, v in fields.items():
        if keep is not None:
            if k in keep:
                out[k] = v
            continue
        if (k.startswith(("_", "@", "OData_")) or k in _SYSTEM_FIELDS
                or (drop_doc_fields and k in _DOC_FIELDS) or v in (None, "", [])):
            continue
        out[k] = v
    return out


def list_item(li: dict, keep: set[str] | None = None) -> dict:
    return compact({
        "item_id": li.get("id"),
        "fields": clean_fields(li.get("fields") or {}, keep),
        "modified": li.get("lastModifiedDateTime"),
        "modified_by": who(li.get("lastModifiedBy")),
        "url": li.get("webUrl"),
    })


def version(v: dict) -> dict:
    return compact({
        "version": v.get("id"),
        "modified": v.get("lastModifiedDateTime"),
        "modified_by": who(v.get("lastModifiedBy")),
        "size": v.get("size"),
    })


def page(p: dict) -> dict:
    return compact({
        "page_id": p.get("id"),
        "title": p.get("title") or p.get("name"),
        "name": p.get("name"),
        "kind": p.get("promotionKind") if p.get("promotionKind") not in (None, "page") else None,
        "description": p.get("description"),
        "modified": p.get("lastModifiedDateTime"),
        "modified_by": who(p.get("lastModifiedBy")),
        "url": p.get("webUrl"),
    })


# ---------------------------------------------------------------- renderers
def _value(v, n: int = 120) -> str:
    if isinstance(v, (dict, list)):
        v = json.dumps(v, ensure_ascii=False, separators=(",", ":"))
    return shorten(str(v), n)


def _mod(d: dict) -> str:
    if not d.get("modified"):
        return ""
    return f"modified {short_time(d['modified'])}" + (f" {d['modified_by']}" if d.get("modified_by") else "")


def render_sites(sites: list[dict]) -> str:
    return "\n".join(
        "• " + " · ".join(x for x in (
            s.get("name"), s.get("url"),
            shorten(s["description"], 80) if s.get("description") else None,
            f"site {s.get('site_id')}") if x)
        for s in sites) or "No sites found."


def render_items(items: list[dict], show_path: bool = False) -> str:
    lines = []
    for d in items:
        if d.get("kind") == "folder":
            head = f"• {d.get('name')}/" + (f" ({d['children']} items)" if d.get("children") is not None else "")
        else:
            head = f"• {d.get('name')} · {size(d.get('size'))}"
        parts = [head]
        if show_path and d.get("path"):
            parts.append(d["path"])
        if _mod(d):
            parts.append(_mod(d))
        if d.get("package"):
            parts.append(d["package"])
        parts.append(f"ref {d.get('ref')}")
        lines.append(" · ".join(parts))
        if d.get("summary"):
            lines.append(f"    {d['summary']}")
    return "\n".join(lines)


def render_hits(hits: list[dict]) -> str:
    lines = []
    for d in hits:
        head = f"• {d.get('name')}{'/' if d.get('kind') == 'folder' else ''}"
        if d.get("kind") != "folder":
            head += f" · {size(d.get('size'))}"
        parts = [head]
        if _mod(d):
            parts.append(_mod(d))
        if d.get("url"):
            parts.append(d["url"])
        parts.append(f"ref {d.get('ref')}")
        lines.append(" · ".join(parts))
        if d.get("summary"):
            lines.append(f"    {d['summary']}")
    return "\n".join(lines) or "No matches."


def render_item(d: dict) -> str:
    kind = "Folder" if d.get("kind") == "folder" else "File"
    lines = [f"{kind}: {d.get('name')}" + (f" · {size(d.get('size'))}" if d.get("kind") == "file" else
                                           f" · {d.get('children', '?')} items")]
    for label, key in (("Path", "path"), ("Type", "mime"), ("Version", "version"), ("State", "level"),
                       ("Shared", "shared"), ("Description", "description")):
        if d.get(key):
            lines.append(f"{label}: {d[key]}")
    if d.get("created"):
        lines.append(f"Created: {short_time(d['created'])} {d.get('created_by') or ''}".rstrip())
    if d.get("modified"):
        lines.append(f"Modified: {short_time(d['modified'])} {d.get('modified_by') or ''}".rstrip())
    if d.get("fields"):
        lines.append("Columns: " + " · ".join(f"{k}: {_value(v, 80)}" for k, v in d["fields"].items()))
    if d.get("url"):
        lines.append(f"Open: {d['url']}")
    lines.append(f"ref {d.get('ref')}" + (f" · parent {d['parent_ref']}" if d.get("parent_ref") else ""))
    return "\n".join(lines)


def render_lists(lists: list[dict]) -> str:
    return "\n".join(
        f"• {x.get('name')} ({x.get('template', '?')})"
        + (f" · {shorten(x['description'], 60)}" if x.get("description") else "")
        + (f" · {_mod(x)}" if x.get("modified") else "")
        + f" · list {x.get('list_id')}"
        for x in lists) or "No lists."


def render_columns(cols: list[dict]) -> str:
    lines = []
    for c in cols:
        flags = [f for f, k in (("required", "required"), ("read-only", "read_only"), ("indexed", "indexed"))
                 if c.get(k)]
        line = f"• {c.get('name')}"
        if c.get("display_name"):
            line += f' ("{c["display_name"]}")'
        line += f": {c.get('type') or '?'}"
        if c.get("details"):
            line += f" ({c['details']})"
        if flags:
            line += " [" + ", ".join(flags) + "]"
        lines.append(line)
    return "\n".join(lines)


def render_list_items(items: list[dict]) -> str:
    lines = []
    for li in items:
        fields = li.get("fields") or {}
        body = " · ".join(f"{k}: {_value(v)}" for k, v in fields.items())
        line = f"• #{li.get('item_id')}" + (f" · {body}" if body else "")
        if li.get("modified"):
            line += f" · {_mod(li)}"
        lines.append(line)
    return "\n".join(lines) or "No items."


def render_versions(vs: list[dict]) -> str:
    return "\n".join(
        f"• v{v.get('version')} · {short_time(v.get('modified'))} {v.get('modified_by') or ''}".rstrip()
        + (f" · {size(v['size'])}" if v.get("size") is not None else "")
        for v in vs) or "No versions."


def render_pages(pages: list[dict]) -> str:
    return "\n".join(
        f"• {p.get('title')}" + (f" ({p['kind']})" if p.get("kind") else "")
        + (f" · {_mod(p)}" if p.get("modified") else "")
        + (f" · {shorten(p['description'], 80)}" if p.get("description") else "")
        + f" · page {p.get('page_id')}"
        for p in pages) or "No pages."


def with_header(header: str, body: str, note: str | None = None) -> str:
    return "\n".join(x for x in (header, body, f"({note})" if note else None) if x)
