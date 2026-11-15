"""Delegated OAuth2 against Entra ID (Microsoft identity platform v2).

authorize.py does the one-time interactive sign-in (authorization code +
PKCE) and stores the refresh token in the encrypted vault. During normal
operation TokenManager turns that refresh token into short-lived access
tokens, keeps them in memory only, and writes each rotated refresh token
back to the vault.
"""
from __future__ import annotations

import asyncio
import logging
import re
import time
from datetime import datetime, timezone

import httpx

from vault import Vault, VaultError

log = logging.getLogger("sharepoint-mcp.auth")

AUTHORITY = "https://login.microsoftonline.com"
# .default = every delegated Graph permission already consented for the app.
GRAPH_SCOPES = "https://graph.microsoft.com/.default offline_access openid"
# Must be registered on the app under the "Web" platform (confidential client).
DEFAULT_REDIRECT_URI = "http://localhost"
REAUTH_COMMAND = "sudo sharepoint-mcp-authorize"
SECRET_COMMAND = ("sudo /opt/sharepoint-mcp/bin/seal-credential.sh client_secret "
                  "&& sudo systemctl restart sharepoint-mcp")

GUID_RE = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
TENANT_RE = re.compile(r"^[A-Za-z0-9.\-]{1,100}$")

# Errors that mean the refresh token can no longer be used: a person has to sign in again.
_REAUTH_ERRORS = {"invalid_grant", "interaction_required", "consent_required", "login_required"}


class AuthError(Exception):
    """Getting a token failed (network, configuration, Entra ID error)."""


class ReauthRequired(AuthError):
    """No usable sign-in: run the interactive authorization on the server."""


class ClientSecretError(AuthError):
    """Entra ID rejected the app's client secret (wrong or expired).
    Fix: seal the current secret and restart. The sign-in itself is still valid."""


def check_client_id(value: str) -> str:
    value = value.strip()
    if not GUID_RE.match(value):
        raise SystemExit(f"MS_CLIENT_ID must be the app (client) id GUID, got {value!r}")
    return value


def check_tenant(value: str) -> str:
    value = value.strip()
    if not TENANT_RE.match(value):
        raise SystemExit(f"MS_TENANT_ID must be a tenant GUID or domain, got {value!r}")
    return value


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


async def token_request(http: httpx.AsyncClient, tenant: str, form: dict) -> dict:
    """POST to the v2 token endpoint. Never logs or echoes the form or response tokens."""
    url = f"{AUTHORITY}/{tenant}/oauth2/v2.0/token"
    try:
        r = await http.post(url, data=form, headers={"Accept": "application/json"})
    except httpx.HTTPError as e:
        raise AuthError(f"Could not reach Entra ID: {e}") from e
    try:
        body = r.json()
    except ValueError:
        body = {}
    if r.status_code == 200 and body.get("access_token"):
        return body
    err = body.get("error") or f"http_{r.status_code}"
    # error_description = "AADSTS70008: ...\r\nTrace ID: ...": keep the first line.
    desc = (body.get("error_description") or r.text[:300] or "").splitlines()
    desc = desc[0] if desc else ""
    if err in _REAUTH_ERRORS:
        raise ReauthRequired(f"{err}: {desc}")
    if err in ("invalid_client", "unauthorized_client"):
        # AADSTS7000222 = secret expired, AADSTS7000215 = wrong secret
        raise ClientSecretError(
            f"Entra ID rejected the app's client secret ({err}: {desc}). Create a new secret "
            f"in the app registration if it expired, then on the server: {SECRET_COMMAND}. "
            f"No new sign-in is needed.")
    raise AuthError(f"Entra ID token endpoint returned {r.status_code} {err}: {desc}")


class TokenManager:
    REFRESH_MARGIN = 300  # seconds before expiry at which a new access token is fetched

    def __init__(self, vault: Vault, tenant: str, client_id: str, client_secret: str,
                 http: httpx.AsyncClient | None = None):
        self.vault = vault
        self.tenant = tenant
        self.client_id = client_id
        self._client_secret = client_secret
        self._http = http or httpx.AsyncClient(timeout=30, follow_redirects=False)
        self._lock = asyncio.Lock()
        self._state: dict | None = None
        self._access: str | None = None
        self._expires = 0.0
        self._problem: str | None = None   # set when the refresh token is dead
        self._last_refresh: str | None = None

    # ---------------------------------------------------------------- state
    def _reload_if_changed(self) -> None:
        if self._state is None or self.vault.changed():
            try:
                state = self.vault.load()
            except VaultError as e:
                self._state = None
                raise ReauthRequired(f"{e} Run: {REAUTH_COMMAND}") from None
            if state is not None and self._state is not None:
                log.info("Vault replaced on disk: using the new sign-in")
            self._state = state
            self._access, self._expires, self._problem = None, 0.0, None

    @property
    def user(self) -> dict:
        return (self._state or {}).get("user") or {}

    def status(self) -> str:
        if self._problem:
            return "reauth_required"
        if self._state is None and not self.vault.path.exists():
            return "not_signed_in"
        return "ok"

    def info(self) -> dict:
        s = self._state or {}
        return {
            "status": self.status(),
            "signed_in_at": s.get("signed_in_at"),
            "refresh_token_rotated_at": s.get("refreshed_at"),
            "last_access_token_at": self._last_refresh,
            "problem": self._problem,
        }

    # ---------------------------------------------------------------- tokens
    async def access_token(self, force: bool = False) -> str:
        async with self._lock:
            self._reload_if_changed()
            if not self._state:
                raise ReauthRequired(f"Not signed in to Microsoft yet. Run on the server: {REAUTH_COMMAND}")
            if self._problem:
                raise ReauthRequired(self._problem)
            if not force and self._access and time.time() < self._expires - self.REFRESH_MARGIN:
                return self._access
            await self._refresh()
            return self._access  # type: ignore[return-value]

    async def _refresh(self) -> None:
        assert self._state is not None
        try:
            body = await token_request(self._http, self.tenant, {
                "client_id": self.client_id,
                "client_secret": self._client_secret,
                "grant_type": "refresh_token",
                "refresh_token": self._state["refresh_token"],
                "scope": GRAPH_SCOPES,
            })
        except ReauthRequired as e:
            self._access, self._expires = None, 0.0
            self._problem = (f"The saved Microsoft sign-in is no longer valid ({e}). "
                             f"Run on the server: {REAUTH_COMMAND}")
            log.error("Refresh token rejected: %s", e)
            raise ReauthRequired(self._problem) from None

        self._access = body["access_token"]
        self._expires = time.time() + int(body.get("expires_in", 3600))
        self._last_refresh = now_iso()

        new_rt = body.get("refresh_token")
        if new_rt and new_rt != self._state["refresh_token"]:
            self._state = {**self._state, "refresh_token": new_rt, "refreshed_at": now_iso()}
            try:
                self.vault.save(self._state)
            except OSError as e:
                # The new token still works from memory, and the previous one stays valid
                # until it expires, so this is not fatal, but it needs attention.
                log.error("Could not store the rotated refresh token in %s: %s", self.vault.path, e)

    async def keepalive(self) -> None:
        """Refresh now so the stored refresh token keeps rotating while the server is idle."""
        if self.status() != "ok":
            return
        try:
            await self.access_token(force=True)
            log.info("Keepalive refresh OK")
        except ReauthRequired as e:
            log.warning("Keepalive: %s", e)
        except ClientSecretError as e:
            log.error("Keepalive: %s", e)
        except AuthError as e:
            log.warning("Keepalive refresh failed (will retry later): %s", e)

    async def aclose(self) -> None:
        await self._http.aclose()
