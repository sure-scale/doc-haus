#!/usr/bin/env bash
# Copied into the app bundle. Existing stacks keep their Terminal ownership.
set -euo pipefail
resources="$(cd "$(dirname "$0")" && pwd)"
. "$resources/installation.sh"
cd "$DOCHAUS_INSTALL_DIR"
. scripts/data-paths.sh
configure_data_paths "$PWD"
. scripts/start-runtime.sh

listener_belongs_to() {
  local child=$1 ancestor=$2 attempts=0
  while [[ "$child" =~ ^[0-9]+$ ]] && [ "$child" -gt 1 ] && [ "$attempts" -lt 64 ]; do
    if [ "$child" = "$ancestor" ]; then return 0; fi
    child=$(ps -p "$child" -o ppid= 2>/dev/null | tr -d '[:space:]')
    attempts=$((attempts + 1))
  done
  return 1
}

owner=$(cat "$DOCHAUS_RUNTIME_DIR/engine.lock/pid" 2>/dev/null || true)
if runtime_launcher "$owner"; then
  port="${DOC_HAUS_WEB_PORT:-5174}"
  listeners=$(lsof -nP -tiTCP:"$port" -sTCP:LISTEN 2>/dev/null || true)
  # An app refreshed with different ports must never open another installation.
  for listener in $listeners; do
    if ! listener_belongs_to "$listener" "$owner"; then
      echo "This app's web port belongs to another process. Refresh the app with the running installation's original port settings." >&2
      exit 1
    fi
  done
  if [ -z "$listeners" ]; then exit 0; fi
  url="http://localhost:$port"
  # The launcher opens the browser itself once startup is ready. A click during
  # warmup must leave it alone instead of opening an unavailable page.
  if curl --fail --silent --max-time 1 --output /dev/null "$url"; then open "$url"; fi
  exit 0
fi
if runtime_pid_running "$owner"; then
  echo "Another process owns this doc.haus runtime. Leave it running and check its Terminal window before opening this installation." >&2
  exit 1
fi
engine=$(cat "$DOCHAUS_RUNTIME_DIR/engine-process.lock" 2>/dev/null || true)
if runtime_pid_running "$engine"; then
  echo "A doc.haus engine is still running without its launcher. Check its Terminal window before opening this installation." >&2
  exit 1
fi

if [ "${1:-launch}" = waiting ]; then
  # The first Terminal may still be preparing dependencies, before it claims
  # the runtime. The applet briefly debounces repeated cold-start requests.
  exit 0
fi
open -a Terminal "$resources/doc.haus.command"
echo started
