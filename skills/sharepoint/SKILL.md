---
name: sharepoint
description: Find, read and update SharePoint content through the sharepoint-mcp server, acting as the signed-in user. Use whenever the user asks about documents, files or folders in SharePoint or a team site, wants to find a document by topic, summarise or extract data from a Word/Excel/PowerPoint/PDF file stored in SharePoint, upload or save a file to SharePoint, organise folders, or read or update a SharePoint list or site page. Also use when the user pastes a sharepoint.com link.
---

# SharePoint (sharepoint-mcp)

The sharepoint-mcp server acts in SharePoint **as the signed-in user**,
through delegated Microsoft Graph permissions. It can reach exactly what the
user can reach. Anything it writes appears as done by the user personally and
is visible to everyone with access to that site.

## Tool names

Clients add their own prefix (for example `mcp__sharepoint_mcp__read_file`).
This skill uses the bare names.

| Tool | Use for |
|---|---|
| `search_files(query, site?, folder?, file_types?, sort?)` | Finding documents by topic, name or content. Usually the first call |
| `list_sites(query?)` → `get_site(site)` → `list_folder(...)` | Browsing: which sites, which libraries, what's in a folder |
| `read_file(item, offset?, max_chars?)` | The text of a file (Word, Excel, PowerPoint, PDF, text, legacy Office) |
| `get_item(item)` | Metadata: path, version, who changed it, check-out state, custom columns, link |
| `list_versions(item)` + `read_file(item, version=)` | "What changed?" / older versions |
| `list_lists` → `get_list_schema` → `list_items` / `get_list_item` | SharePoint lists (trackers, registers, requests) |
| `list_pages` / `read_page` | Site pages and news posts |
| `whoami()` | Which account the server acts as; sign-in status |
| `write_file`, `create_folder`, `move_item`, `copy_item`, `update_file_fields`, `create_list_item`, `update_list_item` | Writing. Absent if the server is read-only |
| `delete_item`, `delete_list_item` | Only if the admin enabled delete tools |

## Addressing things

- Every listing and search hit ends with `ref <drive_id>/<item_id>`. Pass that
  ref as `item`, `folder` or `destination`. Copy it exactly; never build one.
- A **SharePoint URL** works anywhere a ref does: a link the user pasted, a
  sharing link, a folder or library view URL.
- Or use `site` + `library` + `path`: `site` is a URL, site id, or site name;
  `library` is a library name (default: the site's main "Documents" library);
  `path` is relative to the library root (`Reports/2026/Q4.xlsx`).
- `get_site` shows a library's root ref, which `list_folder` accepts.

## Workflow

### Finding a document

1. If the user gave a link, use it directly as `item`.
2. Otherwise `search_files` with the key words. Narrow with `site=` when the
   user names a site or team, `file_types="docx,pdf"` when they name a format.
   KQL works: `title:"Travel policy"`, `author:"Alice Johnson"`,
   `budget AND forecast`, `filename:contract*`.
3. "Latest / recent files in X": `search_files(site=X, sort="modified")` with
   no query.
4. Several plausible hits: show name, location and modified date, and ask which
   one. Don't guess between similar documents (drafts vs final).
5. Search lags a few minutes behind new files. If a file the user just created
   isn't found, browse with `list_folder`.

### Reading

- `read_file` returns up to `max_chars` (default 20000). If the footer says
  `call again with offset=N`, there is more. Read further pages only when the
  question needs them, for example skim page 1, then continue if the answer
  isn't there.
- Word keeps structure: `#` headings, `- ` list items, `| a | b |` tables.
- Excel comes as `## Sheet: Name` blocks; each row starts with its row number,
  then tab-separated cells (empty columns kept as empty cells). Dates are ISO.
  Values are the last calculated values, not formulas.
- PowerPoint comes as `## Slide N`, with `Notes:` for speaker notes.
- PDF pages are separated by `--- page N ---`. "No text layer" means a
  scanned document: tell the user it can't be read as text.
- Quote what the document says; don't present inferences as its content.
  Give the file name (and link from `get_item` if useful) as the source.

### Lists

1. `list_lists(site)` to find the list.
2. `get_list_schema(site, list)` before filtering or writing: display names
   differ from internal names (`"Due Date"` may be `DueDate` or `Due_x0020_Date`).
3. `list_items` with `filter="fields/Status eq 'Open'"`,
   `order_by="fields/Modified desc"`, `fields="Title,Status,DueDate"` to keep
   output small. Dates in filters: `fields/DueDate lt '2026-12-01'`.

### Writing (only if the write tools are present)

Before **every** write, move, copy, metadata update or list change:

1. Resolve the exact target (folder ref or URL, or list and item id).
2. Show the user what will happen: target location, file name, and the
   content or field values. For a replacement, say the existing file will get a
   new version. Then **wait for explicit confirmation**, unless the user
   already said to go ahead without asking.
3. Do it once and report the result line (path, ref, link). Never claim
   success if the tool returned an error.

Details:
- `write_file` with `name` + `folder` creates a file. It fails if the name
  exists; ask before retrying with `conflict="replace"` (overwrites, old
  version kept) or `conflict="rename"`.
- `write_file` with `item` replaces that file's content (new version).
- Text formats (`.md`, `.txt`, `.csv`, `.json`, `.html`) go with
  `encoding="text"`. Binary formats (`.docx`, `.xlsx`, `.pptx`, `.pdf`,
  images) must be built as a real file first and sent with
  `encoding="base64"`. Never send plain text into a `.docx` name.
- To edit a text file: `read_file`, change it, `write_file(item=…)` with the
  full new content.
- `move_item` only works inside one library; across libraries or sites,
  `copy_item` and then (if delete tools exist and the user agrees) delete the original.
- Deletes go to the recycle bin, but always name the exact item and confirm.

## Errors

- "Not signed in" / "sign-in is no longer valid … sharepoint-mcp-authorize":
  tell the user to run `sudo sharepoint-mcp-authorize` on the server, then stop.
- "rejected the app's client secret": the app's secret expired or is wrong; the
  message has the fix. No new sign-in is needed. Stop.
- 403 Forbidden: the user has no access (or only read access, for a write). Say so.
- 404 Not found: wrong ref or path, or no access. Re-list or re-search rather
  than guessing.
- 423 Locked: the file is checked out or open for editing by someone. Tell the user.
- "Several sites match": pick from the list shown, or ask the user.

## Data handling

SharePoint content can be confidential. Quote only what the question needs,
don't copy document content into other tools, files, sites or messages unless
the user asks, and don't move or share content to a location with wider
access than its source without the user's explicit instruction.
