#!/usr/bin/env bash
# Install or update sharepoint-mcp as a hardened systemd service.
# Safe to re-run: keeps /etc/sharepoint-mcp/sharepoint-mcp.env, sealed credentials
# and the stored Microsoft sign-in.
set -euo pipefail
[[ $EUID -eq 0 ]] || { echo "run with sudo" >&2; exit 1; }

SRC="$(cd "$(dirname "$0")/.." && pwd)"
APP=/opt/sharepoint-mcp

python3 -c 'import venv, ensurepip' 2>/dev/null \
  || { echo "Missing venv support: sudo apt install python3-venv" >&2; exit 1; }

echo "==> Service user"
if ! id sharepoint-mcp >/dev/null 2>&1; then
  useradd --system --no-create-home --home-dir /nonexistent --shell /usr/sbin/nologin sharepoint-mcp
  echo "    Created system user sharepoint-mcp"
fi

echo "==> Copying application to $APP"
install -d -m 0755 "$APP" "$APP/app" "$APP/bin"
install -m 0644 "$SRC"/server.py "$SRC"/auth.py "$SRC"/vault.py "$SRC"/graph_client.py \
  "$SRC"/formatting.py "$SRC"/extract.py "$SRC"/authorize.py "$SRC"/requirements.txt "$APP/app/"
install -m 0755 "$SRC"/deploy/authorize.sh "$SRC"/deploy/seal-credential.sh "$APP/bin/"
ln -sf "$APP/bin/authorize.sh" /usr/local/sbin/sharepoint-mcp-authorize

echo "==> Python environment"
[[ -x "$APP/venv/bin/python" ]] || python3 -m venv "$APP/venv"
"$APP/venv/bin/pip" install --quiet --disable-pip-version-check -r "$APP/app/requirements.txt"
"$APP/venv/bin/python" -m compileall -q "$APP/app"

echo "==> Configuration"
install -d -m 0755 /etc/sharepoint-mcp
if [[ ! -f /etc/sharepoint-mcp/sharepoint-mcp.env ]]; then
  install -m 0644 "$SRC/deploy/sharepoint-mcp.env.example" /etc/sharepoint-mcp/sharepoint-mcp.env
  echo "    Created /etc/sharepoint-mcp/sharepoint-mcp.env: set MS_TENANT_ID and MS_CLIENT_ID"
fi
install -d -o sharepoint-mcp -g sharepoint-mcp -m 0700 /var/lib/sharepoint-mcp

echo "==> Credentials"
if [[ ! -f /etc/credstore.encrypted/sharepoint-mcp.vault_key.cred ]]; then
  "$APP/bin/seal-credential.sh" vault_key --generate
fi

echo "==> systemd unit"
install -m 0644 "$SRC/deploy/sharepoint-mcp.service" /etc/systemd/system/sharepoint-mcp.service
systemctl daemon-reload

missing=0
for c in "client_secret" "mcp_token --generate"; do
  if [[ ! -f "/etc/credstore.encrypted/sharepoint-mcp.${c%% *}.cred" ]]; then
    echo "    Missing credential: sudo $APP/bin/seal-credential.sh $c"
    missing=1
  fi
done
if [[ $missing -eq 0 ]] && systemctl is-active --quiet sharepoint-mcp; then
  systemctl restart sharepoint-mcp
  echo "==> Restarted sharepoint-mcp"
fi
[[ -f /var/lib/sharepoint-mcp/vault.json.enc ]] \
  || echo "    Not signed in to Microsoft yet: sudo sharepoint-mcp-authorize"
echo "Done."
