#!/usr/bin/env bash
# Seal a secret for sharepoint-mcp with this machine's TPM.
#
#   sudo ./seal-credential.sh client_secret          # prompts for the Entra app's client secret
#   sudo ./seal-credential.sh mcp_token              # prompts (paste your existing token)
#   sudo ./seal-credential.sh mcp_token --generate   # creates a new random token
#   sudo ./seal-credential.sh vault_key --generate   # creates the key that encrypts the
#                                                    # stored Microsoft sign-in (install.sh
#                                                    # does this for you)
#
# The value is read without echo and never touches disk or shell history.
set -euo pipefail

name="${1:-}"
mode="${2:-}"
case "$name" in
  client_secret|mcp_token|vault_key) ;;
  *) echo "usage: sudo $0 client_secret | mcp_token [--generate] | vault_key --generate [--force]" >&2; exit 2 ;;
esac
[[ $EUID -eq 0 ]] || { echo "run with sudo" >&2; exit 1; }
if [[ "$name" == "client_secret" && -n "$mode" ]]; then
  echo "client_secret can't be generated: paste the secret Value from the app registration" >&2
  exit 2
fi

store=/etc/credstore.encrypted
dest="$store/sharepoint-mcp.$name.cred"

if [[ "$name" == "vault_key" ]]; then
  [[ "$mode" == "--generate" ]] || { echo "vault_key can only be generated: $0 vault_key --generate" >&2; exit 2; }
  if [[ -f "$dest" && "${3:-}" != "--force" ]]; then
    echo "$dest already exists. Replacing it makes the stored Microsoft sign-in unreadable" >&2
    echo "(you then have to run sharepoint-mcp-authorize again). Add --force to do it anyway." >&2
    exit 1
  fi
fi

install -d -m 0700 "$store"
tmp="$(mktemp "$store/.sharepoint-mcp.$name.XXXXXX")"
trap 'rm -f "$tmp"' EXIT

if [[ "$mode" == "--generate" ]]; then
  # 32 random bytes, base64url: an access token for mcp_token, an AES-256 key for vault_key
  python3 -c 'import base64, os; print(base64.urlsafe_b64encode(os.urandom(32)).decode().rstrip("="), end="")' \
    | systemd-creds encrypt --with-key=tpm2 --name="$name" - "$tmp"
else
  systemd-ask-password -n "Value for $name:" \
    | systemd-creds encrypt --with-key=tpm2 --name="$name" - "$tmp"
fi

# Refuse to install an empty secret
len="$(systemd-creds decrypt --name="$name" "$tmp" - | wc -c)"
[[ "$len" -gt 0 ]] || { echo "empty value, nothing sealed" >&2; exit 1; }

chmod 0600 "$tmp"
mv -f "$tmp" "$dest"
trap - EXIT
echo "Sealed $name -> $dest ($len characters, TPM-bound)"

if [[ "$name" == "mcp_token" && "$mode" == "--generate" ]]; then
  echo "Show the new token once (for the MCP client config) with:"
  echo "  sudo systemd-creds decrypt --name=mcp_token $dest -; echo"
fi
if systemctl is-active --quiet sharepoint-mcp; then
  echo "Restart to apply: sudo systemctl restart sharepoint-mcp"
fi
