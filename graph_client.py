"""Minimal Microsoft Graph client using delegated tokens from TokenManager.

Besides JSON calls it handles the two kinds of pre-authenticated URL that
SharePoint file transfers use:
- downloads: GET .../content answers 302 with a short-lived download URL;
- large uploads: createUploadSession returns an upload URL for the chunks.
Those URLs carry their own authorization, so they are fetched with a separate
client that never sends the Graph bearer token.
"""
from __future__ import annotations

import asyncio
import logging
from urllib.parse import urlparse

import httpx

from auth import TokenManager

log = logging.getLogger("sharepoint-mcp.graph")

GRAPH = "https://graph.microsoft.com/v1.0"


class GraphError(Exception):
    """A Graph call failed in a way worth explaining to the user."""

    def __init__(self, message: str, status: int = 0, code: str = ""):
        super().__init__(message)
        self.status = status
        self.code = code


class TooLarge(GraphError):
    """A download went over the configured size cap."""


class GraphClient:
    def __init__(self, tokens: TokenManager, timeout: float = 60,
                 transport: httpx.AsyncBaseTransport | None = None):
        self.tokens = tokens
        self._http = httpx.AsyncClient(
            base_url=GRAPH,
            headers={"Accept": "application/json"},
            timeout=timeout,
            follow_redirects=False,
            transport=transport,
        )
        # For pre-authenticated download/upload/monitor URLs: no Graph token, ever.
        self._plain = httpx.AsyncClient(timeout=timeout, follow_redirects=False, transport=transport)

    async def aclose(self) -> None:
        await self._http.aclose()
        await self._plain.aclose()

    # ------------------------------------------------------------ core
    async def send(self, method: str, path: str, *, params: dict | None = None,
                   json: dict | None = None, content: bytes | None = None,
                   headers: dict | None = None) -> httpx.Response:
        """Authenticated Graph request with 401 refresh and 429/503 retry.
        Raises GraphError on 4xx/5xx; returns 2xx and 3xx responses."""
        # Absolute URLs only come from @odata.nextLink; never send the token elsewhere.
        if path.startswith("http") and not path.startswith(GRAPH + "/"):
            raise GraphError(f"Refusing to follow a link outside Microsoft Graph: {path[:100]}")
        force, retried_401 = False, False
        for attempt in range(4):
            token = await self.tokens.access_token(force=force)
            force = False
            try:
                r = await self._http.request(
                    method, path, params=params, json=json, content=content,
                    headers={**(headers or {}), "Authorization": f"Bearer {token}"},
                )
            except httpx.HTTPError as e:
                raise GraphError(f"Could not reach Microsoft Graph: {e}") from e
            if r.status_code == 401 and not retried_401:
                retried_401, force = True, True
                continue
            if r.status_code in (429, 503, 504) and attempt < 3:
                await asyncio.sleep(_retry_delay(r))
                log.warning("Graph %s on %s, retried", r.status_code, path[:80])
                continue
            break
        if r.status_code >= 400:
            raise self._error(r, method, path)
        return r

    async def request(self, method: str, path: str, *, params: dict | None = None,
                      json: dict | None = None, content: bytes | None = None,
                      headers: dict | None = None) -> dict:
        r = await self.send(method, path, params=params, json=json, content=content, headers=headers)
        if r.status_code == 204 or not r.content:
            return {}
        try:
            return r.json()
        except ValueError:
            return {}

    @staticmethod
    def _error(r: httpx.Response, method: str, path: str) -> GraphError:
        try:
            err = r.json().get("error") or {}
        except ValueError:
            err = {}
        code = err.get("code") or ""
        msg = (err.get("message") or r.text[:300] or "").strip()
        where = f"{method} {path.split('?')[0][:160]}"
        if r.status_code == 401:
            text = f"401 Unauthorized on {where}: {msg}"
        elif r.status_code == 403:
            text = (f"403 Forbidden on {where}: {msg} (the signed-in user, or the app's "
                    f"delegated permissions, don't allow this)")
        elif r.status_code == 404:
            text = f"404 Not found on {where}: {msg} (check the id, path or URL; the user may lack access)"
        elif r.status_code == 409:
            text = f"409 Conflict on {where}: {msg}"
        elif r.status_code == 423:
            text = f"423 Locked on {where}: {msg} (the file is checked out or open for editing)"
        else:
            text = f"Microsoft Graph returned {r.status_code} {code} on {where}: {msg}"
        return GraphError(text, r.status_code, code)

    # ------------------------------------------------------------ helpers
    async def get(self, path: str, params: dict | None = None, headers: dict | None = None) -> dict:
        return await self.request("GET", path, params=params, headers=headers)

    async def post(self, path: str, json: dict, params: dict | None = None) -> dict:
        return await self.request("POST", path, json=json, params=params)

    async def patch(self, path: str, json: dict, headers: dict | None = None) -> dict:
        return await self.request("PATCH", path, json=json, headers=headers)

    async def delete(self, path: str, headers: dict | None = None) -> None:
        await self.request("DELETE", path, headers=headers)

    async def get_all(self, path: str, params: dict | None = None, limit: int = 100,
                      headers: dict | None = None) -> tuple[list[dict], bool]:
        """Follow @odata.nextLink until `limit` items. Returns (items, more_available)."""
        items: list[dict] = []
        url: str | None = path
        while url and len(items) < limit:
            page = await self.request("GET", url, params=params, headers=headers)
            items.extend(page.get("value") or [])
            url = page.get("@odata.nextLink")
            params = None  # the next link already carries the query
        return items[:limit], bool(url) or len(items) > limit

    # ------------------------------------------------------------ pre-authenticated URLs
    @staticmethod
    def _check_url(url: str) -> str:
        u = urlparse(url or "")
        if u.scheme != "https" or not u.hostname:
            raise GraphError(f"Microsoft Graph returned an unexpected transfer URL: {url[:100]!r}")
        return url

    async def download(self, path: str, max_bytes: int, params: dict | None = None) -> bytes:
        """GET a /content endpoint: Graph answers 302 to a pre-authenticated URL,
        which is fetched without the Graph token. Stops at max_bytes."""
        r = await self.send("GET", path, params=params, headers={"Accept": "*/*"})
        if r.status_code in (301, 302, 303, 307, 308):
            url = self._check_url(r.headers.get("Location", ""))
            return await self._fetch(url, max_bytes)
        if len(r.content) > max_bytes:
            raise TooLarge(f"The file is larger than the {max_bytes // 2**20} MB limit.")
        return r.content

    async def _fetch(self, url: str, max_bytes: int) -> bytes:
        for attempt in range(3):
            try:
                async with self._plain.stream("GET", url) as r:
                    if r.status_code in (429, 503) and attempt < 2:
                        await asyncio.sleep(_retry_delay(r))
                        continue
                    if r.status_code >= 400:
                        await r.aread()
                        raise GraphError(f"Download failed with HTTP {r.status_code}: {r.text[:200]}",
                                         r.status_code)
                    size = int(r.headers.get("Content-Length") or 0)
                    if size > max_bytes:
                        raise TooLarge(f"The file is {size / 2**20:.1f} MB, over the "
                                       f"{max_bytes // 2**20} MB limit.")
                    buf = bytearray()
                    async for chunk in r.aiter_bytes():
                        buf.extend(chunk)
                        if len(buf) > max_bytes:
                            raise TooLarge(f"The file is over the {max_bytes // 2**20} MB limit.")
                    return bytes(buf)
            except httpx.HTTPError as e:
                raise GraphError(f"Download failed: {e}") from e
        raise GraphError("Download failed: the server kept throttling the request.")

    async def upload_session(self, upload_url: str, data: bytes,
                             chunk: int = 10 * 327680) -> dict:
        """PUT data to an upload session URL in chunks (multiples of 320 KiB)."""
        self._check_url(upload_url)
        total = len(data)
        start = 0
        result: dict = {}
        while start < total:
            end = min(start + chunk, total)
            for attempt in range(4):
                try:
                    r = await self._plain.put(upload_url, content=data[start:end], headers={
                        "Content-Length": str(end - start),
                        "Content-Range": f"bytes {start}-{end - 1}/{total}",
                    })
                except httpx.HTTPError as e:
                    if attempt < 3:
                        await asyncio.sleep(2 ** attempt)
                        continue
                    raise GraphError(f"Upload failed: {e}") from e
                if r.status_code in (429, 500, 502, 503, 504) and attempt < 3:
                    await asyncio.sleep(_retry_delay(r))
                    continue
                break
            if r.status_code >= 400:
                try:
                    await self._plain.delete(upload_url)
                except httpx.HTTPError:
                    pass
                raise GraphError(f"Upload failed at byte {start} with HTTP {r.status_code}: {r.text[:200]}",
                                 r.status_code)
            if r.status_code in (200, 201):
                result = r.json()
            start = end
        return result

    async def monitor(self, url: str) -> dict:
        """Read an async-operation monitor URL (from POST .../copy). No token needed."""
        self._check_url(url)
        try:
            r = await self._plain.get(url)
        except httpx.HTTPError as e:
            raise GraphError(f"Could not check the operation status: {e}") from e
        if r.status_code == 303 or (r.status_code in (200, 202) and not r.content):
            return {"status": "completed" if r.status_code == 303 else "inProgress"}
        try:
            return r.json()
        except ValueError:
            return {"status": f"http_{r.status_code}"}


def _retry_delay(r: httpx.Response) -> float:
    try:
        delay = float(r.headers.get("Retry-After", "2"))
    except ValueError:
        delay = 2.0
    return min(max(delay, 1.0), 15.0)
