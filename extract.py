"""Turn file bytes into plain text an AI agent can read.

- Word (.docx/.docm), PowerPoint (.pptx/.pptm) and Excel (.xlsx/.xlsm) are
  parsed from their Office Open XML parts with the standard library (through
  defusedxml), keeping headings, list bullets, tables, slide notes and sheets.
- PDF uses pypdf.
- Text-like files (txt, md, csv, json, code, ...) are decoded; HTML is reduced to text.
- Legacy and OpenDocument formats (.doc, .ppt, .xls, .rtf, .odt, ...) are not
  parsed here: the server asks Graph to convert them to PDF first (CONVERT_TO_PDF).

Every parser is defensive: zip members are size-checked before they are read.
"""
from __future__ import annotations

import io
import re
import zipfile
from datetime import datetime, timedelta
from pathlib import PurePosixPath

from defusedxml import ElementTree as ET

import formatting as fmt

TEXT_EXTENSIONS = {
    "txt", "md", "markdown", "csv", "tsv", "json", "jsonl", "xml", "yaml", "yml", "log", "ini",
    "cfg", "conf", "toml", "properties", "env", "sql", "py", "js", "mjs", "ts", "tsx", "jsx",
    "java", "cs", "c", "h", "cpp", "hpp", "go", "rs", "rb", "php", "sh", "bash", "ps1", "psm1",
    "bat", "cmd", "css", "scss", "less", "vb", "vbs", "kql", "r", "m", "swift", "kt", "scala",
    "tex", "srt", "vtt", "ics", "vcf", "eml", "rst", "adoc", "gitignore", "dockerfile", "tf",
    "bicep", "graphql", "proto", "svg",
}
HTML_EXTENSIONS = {"html", "htm", "aspx", "xhtml", "mht"}
OOXML_EXTENSIONS = {"docx", "docm", "dotx", "pptx", "pptm", "potx", "ppsx", "xlsx", "xlsm", "xltx"}
# Graph can render these as PDF (GET .../content?format=pdf), which we then parse.
CONVERT_TO_PDF = {"doc", "dot", "ppt", "pps", "pot", "xls", "xlsb", "rtf", "odt", "ods", "odp",
                  "ppsm", "potm", "dotm"}

MAX_UNZIPPED = 300 * 2**20       # total uncompressed bytes we accept from an Office file
MAX_MEMBER = 100 * 2**20         # any single part
MAX_SHEET_ROWS = 5000            # per worksheet
MAX_PDF_PAGES = 500


class ExtractError(Exception):
    """The file can't be turned into text (unsupported, corrupt, encrypted)."""


def extension(name: str) -> str:
    n = (name or "").lower()
    if n in ("dockerfile", "makefile"):
        return n
    return PurePosixPath(n).suffix.lstrip(".")


def can_extract(name: str, mime: str = "") -> bool:
    ext = extension(name)
    return (ext in TEXT_EXTENSIONS or ext in HTML_EXTENSIONS or ext in OOXML_EXTENSIONS
            or ext == "pdf" or (mime or "").startswith("text/"))


def to_text(data: bytes, name: str, mime: str = "") -> str:
    ext = extension(name)
    if ext in ("docx", "docm", "dotx"):
        return docx_text(data)
    if ext in ("pptx", "pptm", "potx", "ppsx"):
        return pptx_text(data)
    if ext in ("xlsx", "xlsm", "xltx"):
        return xlsx_text(data)
    if ext == "pdf" or data[:5] == b"%PDF-":
        return pdf_text(data)
    if ext in HTML_EXTENSIONS:
        return fmt.html_to_text(decode(data))
    if ext in TEXT_EXTENSIONS or (mime or "").startswith("text/") or _looks_textual(data):
        return decode(data)
    raise ExtractError(f"Can't extract text from a .{ext or '?'} file ({mime or 'unknown type'}).")


def decode(data: bytes) -> str:
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16", errors="replace")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def _looks_textual(data: bytes) -> bool:
    sample = data[:4096]
    if not sample or b"\x00" in sample:
        return False
    try:
        sample.decode("utf-8")
        return True
    except UnicodeDecodeError as e:
        return e.start > len(sample) - 4  # only a multi-byte char cut at the end


# ---------------------------------------------------------------- OOXML plumbing
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


class _Package:
    def __init__(self, data: bytes):
        try:
            self.zip = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile:
            raise ExtractError("The file is not a valid Office document (it may be encrypted "
                               "or password protected).") from None
        infos = self.zip.infolist()
        if sum(i.file_size for i in infos) > MAX_UNZIPPED:
            raise ExtractError("The document expands to more than 300 MB; refusing to unpack it.")
        self.names = {i.filename for i in infos}

    def xml(self, name: str):
        if name not in self.names:
            return None
        if self.zip.getinfo(name).file_size > MAX_MEMBER:
            raise ExtractError(f"Part {name} is too large to parse.")
        try:
            return ET.fromstring(self.zip.read(name))
        except Exception as e:  # defusedxml raises its own classes for hostile XML
            raise ExtractError(f"Part {name} is not valid XML ({type(e).__name__}).") from None

    def rels(self, part: str) -> dict[str, tuple[str, str]]:
        """rId -> (type suffix, absolute part name) for a part's relationships."""
        p = PurePosixPath(part)
        root = self.xml(str(p.parent / "_rels" / (p.name + ".rels")))
        out: dict[str, tuple[str, str]] = {}
        if root is None:
            return out
        for r in root.findall(_q("rel", "Relationship")):
            target = r.get("Target") or ""
            if r.get("TargetMode") == "External":
                continue
            full = target.lstrip("/") if target.startswith("/") else _norm(p.parent / target)
            out[r.get("Id") or ""] = ((r.get("Type") or "").rsplit("/", 1)[-1], full)
        return out


def _norm(path: PurePosixPath) -> str:
    parts: list[str] = []
    for seg in path.parts:
        if seg == "..":
            if parts:
                parts.pop()
        elif seg not in (".", ""):
            parts.append(seg)
    return "/".join(parts)


# ---------------------------------------------------------------- Word
_HEADING = re.compile(r"^(?:heading|kop|titre|überschrift|título|titolo)\s*(\d)$", re.I)


def _w_paragraph(p) -> str:
    out: list[str] = []
    for el in p.iter():
        tag = el.tag
        if tag == _q("w", "t"):
            out.append(el.text or "")
        elif tag == _q("w", "tab"):
            out.append("\t")
        elif tag in (_q("w", "br"), _q("w", "cr")):
            out.append("\n")
        elif tag == _q("w", "noBreakHyphen"):
            out.append("-")
    text = "".join(out).strip()
    if not text:
        return ""
    ppr = p.find(_q("w", "pPr"))
    if ppr is not None:
        style = ppr.find(_q("w", "pStyle"))
        sval = (style.get(_q("w", "val")) if style is not None else "") or ""
        m = _HEADING.match(sval.replace("_", ""))
        if sval.lower() == "title":
            return "# " + text
        if m:
            return "#" * min(int(m.group(1)) + 1, 6) + " " + text
        if ppr.find(_q("w", "numPr")) is not None or sval.lower().startswith("list"):
            return "- " + text
    return text


def _w_table(tbl) -> list[str]:
    rows = []
    for tr in tbl.findall(_q("w", "tr")):
        cells = []
        for tc in tr.findall(_q("w", "tc")):
            parts = [_w_paragraph(p) for p in tc.iter(_q("w", "p"))]
            cells.append(" ".join(x for x in parts if x).replace("|", "/"))
        rows.append("| " + " | ".join(cells) + " |")
    return rows


def _w_block(body) -> list[str]:
    lines: list[str] = []
    for el in body:
        if el.tag == _q("w", "p"):
            t = _w_paragraph(el)
            if t:
                lines.append(t)
        elif el.tag == _q("w", "tbl"):
            lines.extend(_w_table(el))
        elif el.tag == _q("w", "sdt"):  # content controls wrap ordinary blocks
            content = el.find(_q("w", "sdtContent"))
            if content is not None:
                lines.extend(_w_block(content))
    return lines


def docx_text(data: bytes) -> str:
    pkg = _Package(data)
    doc = pkg.xml("word/document.xml")
    if doc is None:
        raise ExtractError("No word/document.xml: not a Word document.")
    body = doc.find(_q("w", "body"))
    lines = _w_block(body) if body is not None else []
    for part, label in (("word/footnotes.xml", "Footnotes"), ("word/comments.xml", "Comments")):
        root = pkg.xml(part)
        if root is None:
            continue
        extra = [t for t in (_w_paragraph(p) for p in root.iter(_q("w", "p"))) if t]
        if extra:
            lines += ["", f"## {label}"] + extra
    return "\n".join(lines)


# ---------------------------------------------------------------- PowerPoint
def _a_paragraphs(root) -> list[str]:
    out = []
    for p in root.iter(_q("a", "p")):
        text = "".join((t.text or "") if t.tag == _q("a", "t") else "\n"
                       for t in p.iter() if t.tag in (_q("a", "t"), _q("a", "br"))).strip()
        if text:
            out.append(text)
    return out


def pptx_text(data: bytes) -> str:
    pkg = _Package(data)
    pres = pkg.xml("ppt/presentation.xml")
    if pres is None:
        raise ExtractError("No ppt/presentation.xml: not a PowerPoint file.")
    rels = pkg.rels("ppt/presentation.xml")
    order = []
    lst = pres.find(_q("p", "sldIdLst"))
    for sid in (lst if lst is not None else []):
        rid = sid.get(_q("r", "id"))
        if rid in rels:
            order.append(rels[rid][1])
    lines: list[str] = []
    for n, part in enumerate(order, 1):
        slide = pkg.xml(part)
        if slide is None:
            continue
        lines.append(f"## Slide {n}")
        lines.extend(_a_paragraphs(slide))
        for kind, target in pkg.rels(part).values():
            if kind == "notesSlide":
                notes = pkg.xml(target)
                if notes is not None:
                    # Skip the slide-number placeholder text Office puts in notes.
                    body = [t for t in _a_paragraphs(notes) if not t.isdigit()]
                    if body:
                        lines.append("Notes: " + "\n".join(body))
        lines.append("")
    return "\n".join(lines).strip()


# ---------------------------------------------------------------- Excel
_BUILTIN_DATE_FORMATS = set(range(14, 23)) | {27, 30, 36, 45, 46, 47, 50, 57}
_EPOCH = datetime(1899, 12, 30)


def _is_date_format(code: str) -> bool:
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "").lower()
    return bool(re.search(r"[dmyh]", code)) and "general" not in code


def _col_index(ref: str) -> int:
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def _excel_date(value: float) -> str:
    dt = _EPOCH + timedelta(days=value)
    if dt.hour == dt.minute == dt.second == 0:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M")


def xlsx_text(data: bytes, max_rows: int = MAX_SHEET_ROWS) -> str:
    pkg = _Package(data)
    wb = pkg.xml("xl/workbook.xml")
    if wb is None:
        raise ExtractError("No xl/workbook.xml: not an Excel workbook.")
    rels = pkg.rels("xl/workbook.xml")

    shared: list[str] = []
    sst = pkg.xml("xl/sharedStrings.xml")
    if sst is not None:
        for si in sst.findall(_q("s", "si")):
            shared.append("".join(t.text or "" for t in si.iter(_q("s", "t"))))

    date_styles: set[int] = set()
    styles = pkg.xml("xl/styles.xml")
    if styles is not None:
        custom = {}
        nf = styles.find(_q("s", "numFmts"))
        for f in (nf if nf is not None else []):
            custom[int(f.get("numFmtId", "0"))] = f.get("formatCode", "")
        xfs = styles.find(_q("s", "cellXfs"))
        for i, xf in enumerate(xfs if xfs is not None else []):
            fid = int(xf.get("numFmtId", "0"))
            if fid in _BUILTIN_DATE_FORMATS or (fid in custom and _is_date_format(custom[fid])):
                date_styles.add(i)

    out: list[str] = []
    sheets = wb.find(_q("s", "sheets"))
    for sh in (sheets if sheets is not None else []):
        name = sh.get("name") or "?"
        rid = sh.get(_q("r", "id"))
        if rid not in rels:
            continue
        root = pkg.xml(rels[rid][1])
        if root is None:
            continue
        state = sh.get("state")
        out.append(f"## Sheet: {name}" + (f" ({state})" if state and state != "visible" else ""))
        rows_out = 0
        sheet_data = root.find(_q("s", "sheetData"))
        for row in (sheet_data if sheet_data is not None else []):
            cells: dict[int, str] = {}
            for i, c in enumerate(row.findall(_q("s", "c"))):
                col = _col_index(c.get("r") or "") if c.get("r") else i
                t = c.get("t")
                v = c.find(_q("s", "v"))
                raw = v.text if v is not None else None
                if t == "s" and raw is not None:
                    try:
                        val = shared[int(raw)]
                    except (ValueError, IndexError):
                        val = ""
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(_q("s", "t")))
                elif t == "b":
                    val = "TRUE" if raw == "1" else "FALSE"
                elif raw is None:
                    continue
                elif t in (None, "n") and int(c.get("s", "0")) in date_styles:
                    try:
                        val = _excel_date(float(raw))
                    except (ValueError, OverflowError):
                        val = raw
                else:
                    val = raw
                if val != "":
                    cells[col] = val.replace("\t", " ").replace("\n", " ")
            if not cells:
                continue
            rnum = row.get("r") or "?"
            width = max(cells) + 1
            out.append(f"{rnum}\t" + "\t".join(cells.get(i, "") for i in range(width)))
            rows_out += 1
            if rows_out >= max_rows:
                out.append(f"(sheet cut at {max_rows} non-empty rows)")
                break
        if rows_out == 0:
            out.append("(empty)")
        out.append("")
    return "\n".join(out).strip()


# ---------------------------------------------------------------- PDF
def pdf_text(data: bytes) -> str:
    try:
        from pypdf import PdfReader
        from pypdf.errors import PdfReadError
    except ImportError:  # pragma: no cover
        raise ExtractError("PDF support needs the pypdf package.") from None
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise ExtractError("The PDF is password protected.") from None
        pages = []
        for n, page in enumerate(reader.pages, 1):
            if n > MAX_PDF_PAGES:
                pages.append(f"(stopped after {MAX_PDF_PAGES} pages)")
                break
            text = fmt.tidy(page.extract_text() or "")
            pages.append(f"--- page {n} ---\n{text}" if text else f"--- page {n} --- (no text)")
    except ExtractError:
        raise
    except (PdfReadError, ValueError, KeyError, TypeError) as e:
        raise ExtractError(f"Can't read the PDF ({type(e).__name__}: {e}).") from None
    text = "\n".join(pages)
    if not re.search(r"\w", re.sub(r"--- page \d+ ---|\(no text\)", "", text)):
        raise ExtractError("The PDF has no text layer (it is probably scanned images).")
    return text
