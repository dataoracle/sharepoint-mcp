"""One-time interactive Microsoft sign-in for sharepoint-mcp.

Normally run through deploy/authorize.sh (installed as sharepoint-mcp-authorize):

  1. prints a Microsoft sign-in URL (authorization code flow with PKCE);
  2. you open it in any browser and sign in; the browser is redirected to
     http://localhost/?code=... and shows "can't connect", which is expected;
  3. you paste that URL from the address bar here (input is not echoed);
  4. the code is exchanged for tokens using the app's client secret (TPM-sealed,
     decrypted into memory only), and the refresh token is written to the vault,
     encrypted with the TPM-sealed vault key. Nothing is written in plain text.

The running service notices the new vault on its next request; no restart needed.
"""
from __future__ import annotations

import argparse
import asyncio
import base64
import getpass
import hashlib
import os
import pwd
import secrets
import subprocess
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse

import httpx

from auth import (AUTHORITY, DEFAULT_REDIRECT_URI, GRAPH_SCOPES, AuthError,
                  check_client_id, check_tenant, now_iso, token_request)
from vault import Vault, VaultError, load_key, vault_context


def die(msg: str) -> None:
    print(f"error: {msg}", file=sys.stderr)
    sys.exit(1)


def read_env_file(path: str | None) -> dict:
    """Parse a systemd EnvironmentFile (KEY=VALUE lines, # comments)."""
    out: dict = {}
    if not path:
        return out
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        out[k.strip()] = v.strip().strip('"').strip("'")
    return out


def read_credential(cred_path: str | None, name: str, dev_env: str) -> str:
    """Decrypt a TPM-sealed systemd credential straight into this process's memory
    (root only). Without a credential path, fall back to an env var (development)."""
    if cred_path:
        r = subprocess.run(["systemd-creds", "decrypt", f"--name={name}", cred_path, "-"],
                           capture_output=True)
        if r.returncode != 0:
            die(f"cannot decrypt {cred_path}: {r.stderr.decode(errors='replace').strip()}")
        value = r.stdout.decode().rstrip("\r\n")
    else:
        value = os.environ.get(dev_env, "")
    if not value:
        die(f"no {name}: pass its sealed credential, or set {dev_env} for local development")
    return value


def get_vault_key(cred_path: str | None) -> bytes:
    text = read_credential(cred_path, "vault_key", "VAULT_KEY")
    try:
        return load_key(text)
    except VaultError as e:
        die(str(e))
        raise


def pkce_pair() -> tuple[str, str]:
    verifier = secrets.token_urlsafe(64)  # 86 chars, within the 43-128 allowed
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
    return verifier, challenge


def extract_code(pasted: str, state: str) -> str:
    pasted = pasted.strip()
    if not pasted:
        die("nothing pasted")
    if "code=" in pasted or "error=" in pasted:
        query = urlparse(pasted).query or pasted.split("?", 1)[-1]
        q = parse_qs(query)
        if "error" in q:
            die(f"sign-in failed: {q['error'][0]}: {q.get('error_description', [''])[0]}")
        if q.get("state", [None])[0] != state:
            die("state mismatch: paste the URL from *this* sign-in attempt")
        if not q.get("code"):
            die("no code= found in the pasted URL")
        return q["code"][0]
    print("note: a bare code was pasted, so the state check was skipped")
    return pasted


async def run(args: argparse.Namespace) -> None:
    env = read_env_file(args.env_file)

    def setting(name: str, default: str | None = None) -> str:
        value = os.environ.get(name) or env.get(name) or default
        if not value:
            die(f"{name} is not set (in {args.env_file or 'the environment'})")
        return value  # type: ignore[return-value]

    tenant = check_tenant(setting("MS_TENANT_ID"))
    client_id = check_client_id(setting("MS_CLIENT_ID"))
    redirect_uri = setting("MS_REDIRECT_URI", DEFAULT_REDIRECT_URI)
    client_secret = read_credential(args.secret_cred, "client_secret", "MS_CLIENT_SECRET")

    owner = None
    if args.owner:
        pw = pwd.getpwnam(args.owner)
        owner = (pw.pw_uid, pw.pw_gid)
    vault = Vault(args.vault, get_vault_key(args.key_cred), vault_context(tenant, client_id))
    if not vault.path.parent.is_dir():
        die(f"{vault.path.parent} does not exist (run install.sh first)")
    try:
        previous = vault.load()
    except VaultError:
        previous = None
    if previous:
        u = previous.get("user") or {}
        print(f"Replacing the current sign-in ({u.get('upn') or u.get('name') or 'unknown user'}, "
              f"signed in {previous.get('signed_in_at')}).")

    verifier, challenge = pkce_pair()
    state = secrets.token_urlsafe(16)
    url = f"{AUTHORITY}/{tenant}/oauth2/v2.0/authorize?" + urlencode({
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": GRAPH_SCOPES,
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
        "prompt": "select_account",
    })
    print("\n1. Open this URL in a browser and sign in with the account the server should use:\n")
    print(url)
    print(f"\n2. The browser is sent to {redirect_uri}?code=... and will most likely show")
    print("   \"can't connect\": that is expected, nothing needs to listen there.")
    print("   Copy the full URL from the address bar. The code expires in about 10 minutes.\n")
    pasted = getpass.getpass("3. Paste it here (input hidden): ")
    code = extract_code(pasted, state)

    async with httpx.AsyncClient(timeout=30, follow_redirects=False) as http:
        try:
            body = await token_request(http, tenant, {
                "client_id": client_id,
                "client_secret": client_secret,
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": redirect_uri,
                "code_verifier": verifier,
                "scope": GRAPH_SCOPES,
            })
        except AuthError as e:
            die(f"code exchange failed: {e}")
        if not body.get("refresh_token"):
            die("no refresh token returned: is offline_access granted to the app?")
        r = await http.get("https://graph.microsoft.com/v1.0/me",
                           params={"$select": "id,displayName,userPrincipalName"},
                           headers={"Authorization": f"Bearer {body['access_token']}"})
        if r.status_code != 200:
            die(f"signed in, but GET /me failed with {r.status_code}: {r.text[:200]}")
        me = r.json()

    vault.save({
        "refresh_token": body["refresh_token"],
        "user": {"id": me.get("id"), "name": me.get("displayName"), "upn": me.get("userPrincipalName")},
        "scope": body.get("scope"),
        "signed_in_at": now_iso(),
        "refreshed_at": now_iso(),
    }, owner=owner)

    print(f"\nSigned in as {me.get('displayName')} <{me.get('userPrincipalName')}>")
    print(f"Granted scopes: {body.get('scope')}")
    print(f"Refresh token encrypted into {vault.path}")
    print("The running service will use it on its next request; no restart needed.")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--env-file", help="systemd EnvironmentFile with MS_TENANT_ID / MS_CLIENT_ID")
    p.add_argument("--key-cred", help="TPM-sealed vault_key credential (systemd-creds)")
    p.add_argument("--secret-cred", help="TPM-sealed client_secret credential (systemd-creds)")
    p.add_argument("--vault", default="vault.json.enc", help="vault file to write")
    p.add_argument("--owner", help="system user that should own the vault file")
    args = p.parse_args()
    try:
        asyncio.run(run(args))
    except KeyboardInterrupt:
        die("cancelled")


if __name__ == "__main__":
    main()
