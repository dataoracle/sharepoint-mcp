"""Offline tests for sharepoint-mcp. Microsoft Graph, the SharePoint download
host and Entra ID are mocked; nothing leaves the machine.

Run:  python tests/test_sharepoint.py      (or: pytest tests/)
Pass -v to print sample output.
"""
import asyncio
import io
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import parse_qsl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import httpx  # noqa: E402

from vault import Vault, load_key, new_key_text, vault_context  # noqa: E402

TENANT, CLIENT = "contoso.onmicrosoft.com", "11111111-2222-3333-4444-555555555555"
_work = Path(tempfile.mkdtemp())
_key = new_key_text()
os.environ.update(MS_TENANT_ID=TENANT, MS_CLIENT_ID=CLIENT, MCP_TOKEN="test-token",
                  MS_CLIENT_SECRET="test-secret", VAULT_KEY=_key, KEEPALIVE_HOURS="0",
                  VAULT_PATH=str(_work / "vault.json.enc"), MAX_READ_CHARS="20000")
os.environ.pop("ENABLE_DELETE_TOOLS", None)
Vault(_work / "vault.json.enc", load_key(_key), vault_context(TENANT, CLIENT)).save(
    {"refresh_token": "RT", "user": {"id": "me-id", "name": "Me", "upn": "me@contoso.example"}})

import extract  # noqa: E402
import server  # noqa: E402
from fastmcp import Client  # noqa: E402

VERBOSE = "-v" in sys.argv
HOST = "contoso.sharepoint.com"
SITE_URL = f"https://{HOST}/sites/Finance"
SID = f"{HOST},aaaaaaaa-0000-0000-0000-000000000001,bbbbbbbb-0000-0000-0000-000000000002"
D, D2 = "b!docsDRIVE_id-1", "b!contractsDRIVE-2"
LID = "cccccccc-0000-0000-0000-000000000003"


# ------------------------------------------------------------------ sample files
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def _zip(parts: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, text in parts.items():
            z.writestr(name, text)
    return buf.getvalue()


def _wp(text, style=None, num=False):
    ppr = ""
    if style or num:
        ppr = "<w:pPr>" + (f'<w:pStyle w:val="{style}"/>' if style else "") + \
              ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if num else "") + "</w:pPr>"
    return f"<w:p>{ppr}<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>"


def _tc(text):
    return f"<w:tc>{_wp(text)}</w:tc>"


DOCX = _zip({"word/document.xml": f"<w:document {W}><w:body>"
             + _wp("Quarterly Report", "Title") + _wp("Summary", "Heading1")
             + "<w:p><w:r><w:t xml:space=\"preserve\">Revenue grew </w:t></w:r><w:r><w:t>12%.</w:t></w:r></w:p>"
             + _wp("Item one", num=True)
             + f"<w:tbl><w:tr>{_tc('Region')}{_tc('Sales')}</w:tr><w:tr>{_tc('EMEA')}{_tc('40')}</w:tr></w:tbl>"
             + "</w:body></w:document>"})

XLSX = _zip({
    "xl/workbook.xml": f'<workbook {S} {R}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/>'
                       f'<sheet name="Empty" sheetId="2" r:id="rId2"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": f'<Relationships {RELS}>'
                                  f'<Relationship Id="rId1" Type="{RT}worksheet" Target="worksheets/sheet1.xml"/>'
                                  f'<Relationship Id="rId2" Type="{RT}worksheet" Target="/xl/worksheets/sheet2.xml"/>'
                                  f'</Relationships>',
    "xl/sharedStrings.xml": f"<sst {S}><si><t>Name</t></si><si><t>Joined</t></si><si><r><t>Ra</t></r><r><t>te</t></r></si></sst>",
    "xl/styles.xml": f'<styleSheet {S}><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>',
    "xl/worksheets/sheet1.xml": f'<worksheet {S}><sheetData>'
                                '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="D1" t="s"><v>2</v></c></row>'
                                '<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c><c r="B2" s="1"><v>46339</v></c>'
                                '<c r="D2"><v>3.5</v></c></row></sheetData></worksheet>',
    "xl/worksheets/sheet2.xml": f"<worksheet {S}><sheetData/></worksheet>",
})

PPTX = _zip({
    "ppt/presentation.xml": f'<p:presentation {P} {R}><p:sldIdLst><p:sldId id="256" r:id="rId7"/>'
                            f'</p:sldIdLst></p:presentation>',
    "ppt/_rels/presentation.xml.rels": f'<Relationships {RELS}><Relationship Id="rId7" Type="{RT}slide" '
                                       f'Target="slides/slide1.xml"/></Relationships>',
    "ppt/slides/slide1.xml": f"<p:sld {P} {A}><a:p><a:r><a:t>Roadmap 2027</a:t></a:r></a:p>"
                             f"<a:p><a:r><a:t>Ship v2</a:t></a:r></a:p></p:sld>",
    "ppt/slides/_rels/slide1.xml.rels": f'<Relationships {RELS}><Relationship Id="rId1" Type="{RT}notesSlide" '
                                        f'Target="../notesSlides/notesSlide9.xml"/></Relationships>',
    "ppt/notesSlides/notesSlide9.xml": f"<p:notes {P} {A}><a:p><a:r><a:t>Speak slowly</a:t></a:r></a:p>"
                                       f"<a:p><a:r><a:t>1</a:t></a:r></a:p></p:notes>",
})


def _pdf(text: str) -> bytes:
    stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode()
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
            b"/Resources << /Font << /F1 5 0 R >> >> >>",
            b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    out, offsets = bytearray(b"%PDF-1.4\n"), []
    for i, o in enumerate(objs, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    out += b"".join(b"%010d 00000 n \n" % off for off in offsets)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    return bytes(out)


PDF = _pdf("Signed contract total 1200 EUR")
LONG = "\n".join(f"line {i:04d} " + "x" * 40 for i in range(1000)).encode()

FILES = {  # item id -> (name, bytes)
    "01DOCX": ("Report.docx", DOCX), "01XLSX": ("Staff.xlsx", XLSX), "01PPTX": ("Plan.pptx", PPTX),
    "01PDF": ("Contract.pdf", PDF), "01LONG": ("big.log", LONG), "01DOC": ("Legacy.doc", b"\xd0\xcf\x11\xe0"),
}


# ------------------------------------------------------------------ mock Graph
def _item(iid, name, folder=None, path="/drives/%s/root:/Reports" % D, size=None, drive=D):
    it = {"id": iid, "name": name, "eTag": f'"{iid},1"', "webUrl": f"{SITE_URL}/Shared%20Documents/Reports/{name}",
          "lastModifiedDateTime": "2026-11-10T09:30:00Z",
          "lastModifiedBy": {"user": {"displayName": "Alice Johnson"}},
          "parentReference": {"driveId": drive, "id": "01FOLDER", "path": path, "siteId": SID}}
    if folder is not None:
        it["folder"] = {"childCount": folder}
    else:
        it["file"] = {"mimeType": "application/octet-stream"}
        it["size"] = size if size is not None else len(FILES.get(iid, ("", b""))[1])
    return it


FOLDER = {**_item("01FOLDER", "Reports", folder=7, path=f"/drives/{D}/root:"),
          "webUrl": f"{SITE_URL}/Shared%20Documents/Reports"}
CHILDREN = [_item("01XLSX", "Staff.xlsx"), _item("01ARCH", "Archive", folder=3), _item("01DOCX", "Report.docx")]
requests: list[httpx.Request] = []


def graph(req: httpx.Request) -> httpx.Response:
    requests.append(req)
    host, p = req.url.host, req.url.path
    q = dict(req.url.params)
    if host == HOST:  # pre-authenticated download URL
        assert "authorization" not in req.headers, "Graph token sent to the download host"
        iid = p.rsplit("/", 1)[-1]
        if iid.endswith(".pdf"):
            return httpx.Response(200, content=PDF)
        return httpx.Response(200, content=FILES[iid][1], headers={"Content-Length": str(len(FILES[iid][1]))})
    assert req.headers.get("authorization") == "Bearer AT", req.headers
    p = p.removeprefix("/v1.0")
    if p == "/me":
        return httpx.Response(200, json={"id": "me-id", "displayName": "Me", "mail": "me@contoso.example"})
    if p in (f"/sites/{HOST}:/sites/Finance", f"/sites/{SID}"):
        return httpx.Response(200, json={"id": SID, "displayName": "Finance", "webUrl": SITE_URL})
    if p.startswith(f"/sites/{HOST}:"):
        return httpx.Response(404, json={"error": {"code": "itemNotFound", "message": "no site"}})
    if p == f"/sites/{SID}/drives":
        return httpx.Response(200, json={"value": [
            {"id": D, "name": "Documents", "webUrl": f"{SITE_URL}/Shared%20Documents", "driveType": "documentLibrary"},
            {"id": D2, "name": "Contracts", "webUrl": f"{SITE_URL}/Contracts", "driveType": "documentLibrary"}]})
    if p == f"/sites/{SID}/drive":
        return httpx.Response(200, json={"id": D, "name": "Documents", "webUrl": f"{SITE_URL}/Shared%20Documents"})
    if p == f"/drives/{D}/root:/Reports" or p == f"/drives/{D}/items/01FOLDER":
        return httpx.Response(200, json=FOLDER)
    if p == f"/drives/{D}/items/01FOLDER/children":
        return httpx.Response(200, json={"value": CHILDREN})
    if p.startswith(f"/drives/{D}/items/01FOLDER:/") and req.method == "PUT":
        name = p.split(":/", 1)[1].removesuffix(":/content")
        if name == "exists.txt":
            return httpx.Response(409, json={"error": {"code": "nameAlreadyExists", "message": "exists"}})
        return httpx.Response(201, json={**_item("01NEW", name, size=len(req.content)),
                                         "parentReference": {"driveId": D, "path": f"/drives/{D}/root:/Reports"}})
    if p.startswith(f"/drives/{D}/items/") and p.endswith("/content"):
        iid = p.split("/")[4]
        if q.get("format") == "pdf":
            return httpx.Response(302, headers={"Location": f"https://{HOST}/dl/{iid}.pdf"})
        return httpx.Response(302, headers={"Location": f"https://{HOST}/dl/{iid}"})
    if p.startswith(f"/drives/{D}/items/"):
        iid = p.split("/")[4]
        if iid in FILES:
            return httpx.Response(200, json=_item(iid, FILES[iid][0]))
    if p == "/search/query":
        body = json.loads(req.content)
        return httpx.Response(200, json={"value": [{"hitsContainers": [{"total": 42, "moreResultsAvailable": True,
            "hits": [{"summary": "the <c0>budget</c0> for Q4<ddd/>", "resource": {
                **_item("01XLSX", "Budget.xlsx", size=2048), "webUrl": f"{SITE_URL}/Shared%20Documents/Budget.xlsx"}}],
            "_echo": body}]}]})
    if p == f"/sites/{SID}/lists":
        return httpx.Response(200, json={"value": [
            {"id": LID, "displayName": "Tasks", "name": "Tasks", "list": {"template": "genericList"}},
            {"id": "dddddddd-0000-0000-0000-000000000004", "displayName": "Documents",
             "list": {"template": "documentLibrary"}}]})
    if p == f"/sites/{SID}/lists/{LID}/items":
        return httpx.Response(200, json={"value": [{"id": "7", "lastModifiedDateTime": "2026-11-12T08:00:00Z",
            "lastModifiedBy": {"user": {"displayName": "Bob"}},
            "fields": {"@odata.etag": "x", "Title": "Renew licence", "Status": "Open", "_UIVersionString": "1.0",
                       "ContentType": "Item", "Edit": "", "AuthorLookupId": "12"}}]})
    return httpx.Response(404, json={"error": {"code": "itemNotFound", "message": f"mock: {req.method} {p}"}})


def token(req: httpx.Request) -> httpx.Response:
    form = dict(parse_qsl(req.content.decode()))
    assert form["client_secret"] == "test-secret"
    return httpx.Response(200, json={"access_token": "AT", "expires_in": 3600, "refresh_token": "RT"})


server.tokens._http = httpx.AsyncClient(transport=httpx.MockTransport(token))
server.graph._http = httpx.AsyncClient(base_url="https://graph.microsoft.com/v1.0",
                                       transport=httpx.MockTransport(graph))
server.graph._plain = httpx.AsyncClient(transport=httpx.MockTransport(graph))


async def _call(name: str, args: dict) -> str:
    async with Client(server.mcp) as c:
        res = await c.call_tool(name, args, raise_on_error=False)
    return res.content[0].text


def call(tool: str, /, **args) -> str:
    text = asyncio.run(_call(tool, args))
    if VERBOSE:
        print(f"\n--- {tool}({', '.join(f'{k}={v!r}' for k, v in args.items())})\n{text}")
    return text


def tool_names() -> set[str]:
    async def go():
        async with Client(server.mcp) as c:
            return {t.name for t in await c.list_tools()}
    return asyncio.run(go())


# ------------------------------------------------------------------ tests
def test_tools_registered_and_delete_off_by_default():
    names = tool_names()
    assert {"read_file", "search_files", "list_folder", "write_file", "copy_item", "update_list_item"} <= names
    assert "delete_item" not in names and "delete_list_item" not in names, names


def test_get_site_by_url_lists_libraries():
    out = call("get_site", site=SITE_URL)
    assert out.splitlines()[0] == f"Finance · {SITE_URL}", out
    assert f"• Documents · ref {D}/root" in out and f"• Contracts · ref {D2}/root" in out, out


def test_list_folder_by_site_library_path_folders_first():
    out = call("list_folder", site=SITE_URL, library="Shared Documents", path="/Reports/")
    lines = out.splitlines()
    assert lines[0].startswith("Reports: 3 items (folders first"), out
    assert lines[1].startswith("• Archive/ (3 items) · modified 2026-11-10 09:30 Alice Johnson"), out
    assert lines[2].startswith("• Report.docx · ") and lines[2].endswith(f"ref {D}/01DOCX"), out


def test_list_folder_by_plain_url():
    requests.clear()
    out = call("list_folder", item=f"{SITE_URL}/Shared%20Documents/Reports")
    assert out.startswith("Reports: 3 items"), out
    # Site found from the URL (cut at the library segment), no sharing-link lookup needed.
    assert not any("/shares/" in r.url.path for r in requests), [r.url.path for r in requests]


def test_read_docx_keeps_structure_and_never_leaks_token():
    out = call("read_file", item=f"{D}/01DOCX")
    assert out.splitlines()[1:] == ["# Quarterly Report", "## Summary", "Revenue grew 12%.", "- Item one",
                                    "| Region | Sales |", "| EMEA | 40 |"], out


def test_read_xlsx_sheets_dates_and_shared_strings():
    out = call("read_file", item=f"{D}/01XLSX")
    assert out.splitlines()[1:] == ["## Sheet: Data", "1\tName\tJoined\t\tRate", "2\tAlice\t2026-11-13\t\t3.5",
                                    "", "## Sheet: Empty", "(empty)"], out


def test_read_pptx_slides_and_notes():
    out = call("read_file", item=f"{D}/01PPTX")
    assert out.splitlines()[1:] == ["## Slide 1", "Roadmap 2027", "Ship v2", "Notes: Speak slowly"], out


def test_read_pdf():
    out = call("read_file", item=f"{D}/01PDF")
    assert "--- page 1 ---" in out and "Signed contract total 1200 EUR" in out, out


def test_read_legacy_doc_converts_via_pdf():
    requests.clear()
    out = call("read_file", item=f"{D}/01DOC")
    assert "converted from .doc via PDF" in out and "Signed contract" in out, out
    assert any(r.url.params.get("format") == "pdf" for r in requests)


def test_read_file_pages_long_text():
    first = call("read_file", item=f"{D}/01LONG", max_chars=5000)
    assert first.endswith("for more)"), first[-200:]
    nxt = int(first.rsplit("offset=", 1)[1].split()[0])
    second = call("read_file", item=f"{D}/01LONG", offset=nxt, max_chars=5000)
    # Page breaks fall on line boundaries, and page 2 starts with the line after page 1.
    last = first.splitlines()[-2]
    assert last.endswith("x" * 40), last
    n = int(last.split()[1])
    assert second.splitlines()[1] == f"line {n + 1:04d} " + "x" * 40, second[:200]


def test_search_scoped_to_site_with_types():
    requests.clear()
    out = call("search_files", query="budget", site=SITE_URL, file_types="xlsx,.pdf", sort="modified")
    body = json.loads([r for r in requests if r.url.path.endswith("/search/query")][0].content)
    req = body["requests"][0]
    assert req["query"]["queryString"] == f'budget AND (filetype:xlsx OR filetype:pdf) AND path:"{SITE_URL}"', req
    assert req["sortProperties"][0]["name"] == "lastModifiedDateTime"
    assert "• Budget.xlsx · 2.0 KB" in out and "    the budget for Q4…" in out, out
    assert out.endswith("(more results: call again with offset=1)"), out


def test_write_text_file_and_guards():
    out = call("write_file", content="# Notes\nhello", name="notes.md", folder=f"{D}/01FOLDER")
    assert out.startswith("Uploaded Reports/notes.md (13 B) · ref"), out
    put = [r for r in requests if r.method == "PUT"][-1]
    assert put.url.params["@microsoft.graph.conflictBehavior"] == "fail" and put.content == b"# Notes\nhello"
    bad = call("write_file", content="plain text", name="memo.docx", folder=f"{D}/01FOLDER")
    assert "binary format" in bad, bad
    clash = call("write_file", content="x", name="exists.txt", folder=f"{D}/01FOLDER")
    assert "already exists" in clash and "conflict='replace'" in clash, clash
    bad_name = call("write_file", content="x", name="a:b.txt", folder=f"{D}/01FOLDER")
    assert "Invalid name" in bad_name, bad_name


def test_list_items_filter_sends_prefer_header_and_hides_system_fields():
    requests.clear()
    out = call("list_items", site=SITE_URL, list="tasks", filter="fields/Status eq 'Open'")
    r = [r for r in requests if r.url.path.endswith("/items")][0]
    assert r.headers.get("prefer") == "HonorNonIndexedQueriesWarningMayFailRandomly"
    assert r.url.params["$filter"] == "fields/Status eq 'Open'"
    assert out.splitlines()[1] == "• #7 · Title: Renew licence · Status: Open · modified 2026-11-12 08:00 Bob", out


def test_bad_ref_explains_what_to_pass():
    out = call("read_file", item="not a ref")
    assert "neither a ref" in out, out


def test_extract_rejects_zip_bomb_and_garbage():
    try:
        extract.to_text(b"not a zip", "x.docx")
        raise AssertionError("expected ExtractError")
    except extract.ExtractError as e:
        assert "not a valid Office document" in str(e)


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_")]
    for name, fn in tests:
        fn()
        print(f"ok  {name}")
    print(f"{len(tests)} passed")
