"""Encrypted on-disk store for the Microsoft refresh token.

Entra ID rotates refresh tokens: every refresh returns a new one, and each is
only valid for a limited time. The service therefore has to rewrite the stored
token while it runs, which it cannot do with TPM sealing directly (the sandboxed
service has no TPM access). Instead:

- a random 256-bit vault key is TPM-sealed with systemd-creds (like mcp_token)
  and handed to the service in memory by systemd at start;
- the refresh token and sign-in metadata live in a small file encrypted with
  AES-256-GCM under that key, bound to the tenant and app id (as AAD).

Without the TPM-sealed key the file is useless, so a copy of the disk or a
backup does not expose the token.
"""
from __future__ import annotations

import base64
import json
import os
import tempfile
from pathlib import Path

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

FORMAT_VERSION = 1


class VaultError(Exception):
    """The vault exists but cannot be used (wrong key, corrupt, unreadable)."""


def _b64e(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode().rstrip("=")


def _b64d(text: str) -> bytes:
    text = text.strip()
    return base64.urlsafe_b64decode(text + "=" * (-len(text) % 4))


def new_key_text() -> str:
    """A fresh vault key in the text form stored in the sealed credential."""
    return _b64e(os.urandom(32))


def load_key(text: str) -> bytes:
    try:
        key = _b64d(text)
    except ValueError:
        key = b""
    if len(key) != 32:
        raise VaultError("vault_key must be 32 random bytes, base64url-encoded "
                         "(create it with: seal-credential.sh vault_key --generate).")
    return key


def vault_context(tenant: str, client_id: str) -> str:
    """Associated data: a vault written for one tenant/app won't open for another."""
    return f"sharepoint-mcp/v{FORMAT_VERSION}|{tenant.strip().lower()}|{client_id.strip().lower()}"


class Vault:
    def __init__(self, path: str | os.PathLike, key: bytes, context: str):
        self.path = Path(path)
        self._aead = AESGCM(key)
        self._aad = context.encode()
        self._seen: tuple[int, int] | None = None  # (inode, mtime_ns) last read or written

    def _stamp(self) -> tuple[int, int] | None:
        try:
            st = os.stat(self.path)
        except FileNotFoundError:
            return None
        return (st.st_ino, st.st_mtime_ns)

    def changed(self) -> bool:
        """True if the file was replaced since we last read or wrote it
        (for example by a new sign-in with authorize.sh)."""
        return self._stamp() != self._seen

    def load(self) -> dict | None:
        """Decrypt and return the stored state, or None if there is no vault yet."""
        try:
            with open(self.path, "rb") as f:
                st = os.fstat(f.fileno())
                blob = f.read()
        except FileNotFoundError:
            self._seen = None
            return None
        except OSError as e:
            raise VaultError(f"Cannot read {self.path}: {e}") from e
        try:
            doc = json.loads(blob)
            if doc.get("v") != FORMAT_VERSION:
                raise VaultError(f"{self.path} has unsupported format {doc.get('v')!r}.")
            plain = self._aead.decrypt(_b64d(doc["nonce"]), _b64d(doc["ct"]), self._aad)
            state = json.loads(plain)
        except InvalidTag:
            raise VaultError(
                "The stored sign-in cannot be decrypted: the vault key was changed, or it "
                "was written for a different tenant or app id. Sign in again."
            ) from None
        except (ValueError, KeyError, TypeError) as e:
            raise VaultError(f"{self.path} is corrupt ({type(e).__name__}). Sign in again.") from None
        if not isinstance(state, dict) or not state.get("refresh_token"):
            raise VaultError(f"{self.path} holds no refresh token. Sign in again.")
        self._seen = (st.st_ino, st.st_mtime_ns)
        return state

    def save(self, state: dict, owner: tuple[int, int] | None = None) -> None:
        """Encrypt and atomically replace the vault file (mode 0600)."""
        nonce = os.urandom(12)
        ct = self._aead.encrypt(nonce, json.dumps(state, separators=(",", ":")).encode(), self._aad)
        blob = json.dumps({"v": FORMAT_VERSION, "alg": "AES-256-GCM",
                           "nonce": _b64e(nonce), "ct": _b64e(ct)}).encode()
        fd, tmp = tempfile.mkstemp(dir=self.path.parent, prefix=".vault.", suffix=".tmp")
        try:
            os.fchmod(fd, 0o600)
            if owner:
                os.fchown(fd, *owner)
            with os.fdopen(fd, "wb") as f:
                f.write(blob)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp, self.path)
        except BaseException:
            try:
                os.unlink(tmp)
            except FileNotFoundError:
                pass
            raise
        self._seen = self._stamp()
