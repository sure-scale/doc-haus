#!/usr/bin/env bash
# doc.haus launcher — starts the three processes the stack needs as one command:
#   1. the native OpenCode server graph with the isolated doc.haus config layer
#   2. the ingest service (matters + DOCX embedding + redlines)
#   3. the web app
# This is additive (upstream has no start.sh) so it never affects `git merge upstream/dev`.
set -euo pipefail
cd "$(dirname "$0")"
# Select Bun before anything else runs, offering to install it when it is
# missing or too old. Optional host tools are checked later, before dependencies.
. "$PWD/scripts/requirements.sh"
ensure_bun || exit 1
export BUN_BIN
if [ "$(basename "$BUN_BIN")" = "bun" ]; then
  export PATH="$(dirname "$BUN_BIN"):$PATH"
fi

# Load provider credentials from dochaus/.env.local if present so all three
# processes inherit them. The engine runs with cwd = repo root, so Bun's
# auto-.env loading (which only reads the cwd) never picks up dochaus/.env.local;
# sourcing it here with auto-export covers the engine, ingest, and web alike.
# Gitignored, so real keys never enter version control.
if [ -f dochaus/.env.local ]; then
  set -a
  . dochaus/.env.local
  set +a
fi

# --demo seeds the fictional Aldgate Mills letter of engagement before launch, so a
# first run lands on a matter with cited Q&A and a legal review without uploading
# anything. The seed is idempotent — pass it on every boot, it only ingests once.
DEMO=
for arg in "$@"; do
  [ "$arg" = "--demo" ] && DEMO=1
done

# WORKSPACE_ROOT is where matters live. Defaults to ./workspace at the repo root
# (gitignored) like opencode defaults its project dir to process.cwd() — set it to
# an absolute path to keep matters outside the repo.
# DOCHAUS_DATA_ROOT opts into a common workspace/runtime backup boundary.
. "$PWD/scripts/data-paths.sh"
configure_data_paths "$PWD"
# The engine scans these optional firm-library paths at startup. Create the
# empty libraries before it starts so a new workspace is valid without first
# importing a skill or playbook.
mkdir -p "$WORKSPACE_ROOT" "$WORKSPACE_ROOT/.skills" "$WORKSPACE_ROOT/.playbooks"

# The web product owns a dedicated OpenCode runtime boundary under
# .dochaus-runtime. This is always on: arbitrary global OpenCode plugins, MCPs,
# agents, instructions, skills, conversations, and provider accounts must never
# become part of a legal-matter model request. The data directory is durable, so
# a restart resumes the same legal sessions and provider connections.
mkdir -p \
  "$DOCHAUS_RUNTIME_DIR/data" \
  "$DOCHAUS_RUNTIME_DIR/config" \
  "$DOCHAUS_RUNTIME_DIR/state" \
  "$DOCHAUS_RUNTIME_DIR/cache"

# No provider env is required to launch: the engine reads provider credentials
# lazily, and you connect a provider + pick a model in the web UI (Settings) on
# first run. These two are only consumed if you choose the bundled Google Vertex
# provider — set GOOGLE_VERTEX_PROJECT then, or leave them unset for any other.
# Provider region/location choices are saved explicitly through Settings.

# The stack ports are configurable to support parallel copies. Defaults are set to
# non-standard values to avoid clashing with a separate default-opencode stack.
ENGINE_PORT="${DOC_HAUS_ENGINE_PORT:-4097}"
INGEST_PORT="${DOC_HAUS_INGEST_PORT:-4501}"
WEB_PORT="${DOC_HAUS_WEB_PORT:-5174}"
OPENCODE_HOST="${OPENCODE_HOSTNAME:-127.0.0.1}"
INGEST_HOST="${INGEST_HOST:-127.0.0.1}"
WEB_HOST="${WEB_HOST:-$OPENCODE_HOST}"
VITE_OPENCODE_URL="${VITE_OPENCODE_URL:-http://localhost:${ENGINE_PORT}}"
VITE_INGEST_URL="${VITE_INGEST_URL:-http://localhost:${INGEST_PORT}}"
PORTS=("$ENGINE_PORT" "$INGEST_PORT" "$WEB_PORT")

# Provider retry timing and attempts belong to OpenCode. Do not impose a
# launcher ceiling that converts a brief transport throttle into a workflow STOP.
# Drafting checkpoints, cancellation and substantive retry budgets remain separate.

# pids holds our direct children. `bun run dev`/`bun run start` fork grandchildren
# (vite, the ingest server) that a plain `kill $pids` would orphan — those orphans
# are what squat the ports across restarts. So tear down the whole subtree.
pids=()
kill_tree() {
  local pid=$1 child
  for child in $(pgrep -P "$pid" 2>/dev/null); do kill_tree "$child"; done
  kill "$pid" 2>/dev/null || true
}
cleanup() { for pid in ${pids[@]+"${pids[@]}"}; do kill_tree "$pid"; done; }
# Ctrl-C and TERM exit the launcher; the EXIT trap then tears every process down.
# Exiting (rather than resuming) keeps the child monitor at the bottom from
# reporting the processes cleanup just stopped as unexpected exits.
trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM
update_requested=
trap 'update_requested=1' USR1
export DOCHAUS_LAUNCHER_PID=$$

# A second stack must never open the same durable session database. The engine's
# native restart sweep assumes sole ownership after its previous process exits.
ENGINE_LOCK="$DOCHAUS_RUNTIME_DIR/engine.lock"
. scripts/start-runtime.sh
claim_status=0
claim_engine_runtime "$DOCHAUS_RUNTIME_DIR" || claim_status=$?
if [ "$claim_status" = 2 ]; then exit 0; fi
if [ "$claim_status" != 0 ]; then exit "$claim_status"; fi
cleanup() {
  for pid in ${pids[@]+"${pids[@]}"}; do kill_tree "$pid"; done
  if [ -n "${BUN_PATH_DIR:-}" ]; then
    rm -f "$BUN_PATH_DIR/bun"
    rmdir "$BUN_PATH_DIR"
  fi
  if [ "$(cat "$ENGINE_LOCK/pid" 2>/dev/null || true)" = "$$" ]; then
    rm -f "$ENGINE_LOCK/pid"
    rmdir "$ENGINE_LOCK" 2>/dev/null || true
  fi
}

# Refuse occupied ports instead of terminating another process.
for port in "${PORTS[@]}"; do
  stale=$(lsof -nP -tiTCP:"$port" -sTCP:LISTEN 2>/dev/null || true)
  if [ -n "$stale" ]; then
    echo "doc.haus: port $port is already in use (pid $stale)" >&2
    exit 1
  fi
done

# A custom executable name needs a bun alias for children that look it up on
# PATH. Keep that alias private to this launcher, and clean it up on exit.
if [ "$(basename "$BUN_BIN")" != "bun" ]; then
  BUN_PATH_DIR="$(mktemp -d "$DOCHAUS_RUNTIME_DIR/bun-path.XXXXXX")"
  ln -s "$BUN_BIN" "$BUN_PATH_DIR/bun"
  export PATH="$BUN_PATH_DIR:$PATH"
fi

# Check ownership and ports before installing, so an already-running stack can
# be left alone without changing tools or dependencies underneath it.
add_tool_paths
if [ -z "${SKIP_INSTALL:-}" ]; then
  ensure_host_tools "$DOCHAUS_RUNTIME_DIR"
  for dir in . services/ingest; do
    "$BUN_BIN" run scripts/install-dependencies.ts "$dir"
  done
fi

# Make sure the local embedding model loads before anything is declared ready:
# every upload and every document search needs it, and its first download comes
# from huggingface.co. The model is cached once per user, outside node_modules,
# in ~/.cache/doc-haus/models (set DOCHAUS_MODEL_CACHE_DIR to move it, for
# example onto a persistent volume); ingest and the engine share that cache, so
# after the first launch this check is a fast local load. A transient network
# failure gets two more attempts before the launch stops.
export DOCHAUS_MODEL_CACHE_DIR="${DOCHAUS_MODEL_CACHE_DIR:-$HOME/.cache/doc-haus/models}"
echo "doc.haus: checking the embedding model ($DOCHAUS_MODEL_CACHE_DIR)"
model_attempt=1
until (cd services/ingest && "$BUN_BIN" run src/warm-embedding-model.ts); do
  if [ "$model_attempt" -ge 3 ]; then
    cat >&2 <<EOF
doc.haus: the embedding model could not be loaded after $model_attempt attempts.
  It downloads once (about 50 MB) from https://huggingface.co into
  $DOCHAUS_MODEL_CACHE_DIR. Check that this machine can reach huggingface.co
  (HTTPS_PROXY is honoured) and that the directory is writable, then run
  ./start.sh again. Set DOCHAUS_MODEL_CACHE_DIR to use a different directory.
EOF
    exit 1
  fi
  echo "doc.haus: retrying the embedding model in $((model_attempt * 5))s" >&2
  sleep $((model_attempt * 5))
  model_attempt=$((model_attempt + 1))
done

# Seed the demo matter before the stack comes up so it is already selectable on the
# first page load. Runs inline (not backgrounded) via the ingest pipeline's own
# imports — no server needed — and inherits the WORKSPACE_ROOT exported above.
if [ -n "$DEMO" ]; then
  echo "doc.haus: seeding demo matter"
  (cd services/ingest && "$BUN_BIN" run src/seed.ts)
fi

WEB_URL="http://localhost:${WEB_PORT}"
echo "doc.haus: starting engine, ingest, web (workspace: $WORKSPACE_ROOT)"
echo "doc.haus: ports engine=$ENGINE_PORT ingest=$INGEST_PORT web=$WEB_PORT"
echo "doc.haus: web UI will be at $WEB_URL"

# Pin the engine to the configured port so it fails loudly if the port is still
# taken rather than silently drifting to a random port the web app cannot reach.
# Bind loopback by default; set OPENCODE_HOSTNAME=0.0.0.0 to expose it (Docker,
# reverse proxy) — the same loopback-default the engine, ingest, and vite all keep.
# Export the ingest URL and ports so the engine and any tools can dynamically
# target the correct port start.sh is configured to use, instead of falling back
# to stale defaults.
export INGEST_URL="http://127.0.0.1:${INGEST_PORT}"
export DOC_HAUS_INGEST_PORT="${INGEST_PORT}"
export DOC_HAUS_ENGINE_PORT="${ENGINE_PORT}"

(
  # Preserve only the native credential source for explicit, selected reuse.
  export DOCHAUS_NATIVE_OPENCODE_DATA_DIR="${DOCHAUS_NATIVE_OPENCODE_DATA_DIR:-${XDG_DATA_HOME:-$HOME/.local/share}/opencode}"
  export XDG_DATA_HOME="$DOCHAUS_RUNTIME_DIR/data"
  export XDG_CONFIG_HOME="$DOCHAUS_RUNTIME_DIR/config"
  export XDG_STATE_HOME="$DOCHAUS_RUNTIME_DIR/state"
  export XDG_CACHE_HOME="$DOCHAUS_RUNTIME_DIR/cache"
  # The legal engine accepts configuration only from its isolated runtime and
  # the repository-shipped doc.haus config directory. Do not let a developer's
  # shell inject an alternate config body or permission map.
  unset OPENCODE_CONFIG OPENCODE_CONFIG_CONTENT OPENCODE_PERMISSION
  export DOCHAUS_MANAGED_RESTART=1
  exec "$BUN_BIN" run services/engine/src/server.ts
) &
engine_pid=$!
pids+=("$engine_pid")

# Native inventories populate asynchronously after HTTP starts listening. Wait
# for the legal plugin before exposing the rest of this stack.
"$BUN_BIN" run services/engine/script/wait-ready.ts

(cd services/ingest && INGEST_HOST="$INGEST_HOST" INGEST_PORT="$INGEST_PORT" "$BUN_BIN" run src/server.ts) &
ingest_pid=$!
pids+=("$ingest_pid")

(cd apps/web && VITE_OPENCODE_URL="$VITE_OPENCODE_URL" VITE_INGEST_URL="$VITE_INGEST_URL" "$BUN_BIN" run dev -- --host "$WEB_HOST" --port "$WEB_PORT") &
web_pid=$!
pids+=("$web_pid")

# Retained native work starts only after the whole managed stack can serve it.
# The exec-owned PID is our direct child; no durable release marker is needed.
"$BUN_BIN" run services/engine/script/wait-stack-ready.ts
if [ "$(cat "$DOCHAUS_RUNTIME_DIR/engine-process.lock")" != "$engine_pid" ]; then
  echo "doc.haus: managed restart release lost its exact engine process owner" >&2
  exit 1
fi
kill -USR2 "$engine_pid"

# Open the web UI once vite is actually serving (poll the port via bash's /dev/tcp
# so it works without curl). Vite binds IPv6 localhost (::1) by default, so probe
# both stacks. Opt out with NO_OPEN=1; harmless if no opener exists.
if [ -z "${NO_OPEN:-}" ]; then
  (
    for _ in $(seq 1 60); do
      if (exec 3<>/dev/tcp/127.0.0.1/"$WEB_PORT") 2>/dev/null || (exec 3<>/dev/tcp/::1/"$WEB_PORT") 2>/dev/null; then
        break
      fi
      sleep 0.5
    done
    echo "doc.haus: opening $WEB_URL"
    if command -v open >/dev/null 2>&1; then open "$WEB_URL"
    elif command -v xdg-open >/dev/null 2>&1; then xdg-open "$WEB_URL" >/dev/null 2>&1
    fi
  ) &
  pids+=($!)
fi

# Run until interrupted (Ctrl-C) or until a managed process exits on its own. The
# stack cannot serve without all three, so one exiting stops the launch with a
# non-zero status and the EXIT trap tears the rest down. Polling with kill -0
# keeps this portable to the bash 3.2 that ships on macOS (`wait -n` is 4.0+).
while true; do
  if [ -n "$update_requested" ]; then
    echo "doc.haus: installing the requested update..."
    # Keep the launcher lock until its updater finishes. No second stack may
    # start against the session database during installation.
    for pid in ${pids[@]+"${pids[@]}"}; do kill_tree "$pid"; done
    for pid in ${pids[@]+"${pids[@]}"}; do wait "$pid" 2>/dev/null || true; done
    pids=()
    "$BUN_BIN" run scripts/updater.ts --apply || true
    cleanup
    trap - EXIT
    export NO_OPEN=1
    exec bash ./start.sh "$@"
  fi
  for name in engine ingest web; do
    case "$name" in
      engine) pid=$engine_pid ;;
      ingest) pid=$ingest_pid ;;
      web) pid=$web_pid ;;
    esac
    if ! kill -0 "$pid" 2>/dev/null; then
      status=0
      wait "$pid" || status=$?
      echo "doc.haus: $name exited (status $status); stopping the stack" >&2
      exit 1
    fi
  done
  sleep 1
done
