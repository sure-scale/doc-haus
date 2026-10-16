#!/usr/bin/env bash

# Claim this launcher runtime, or let an interactive caller restart its owner.
# Return 2 when the existing stack should simply be left running.
claim_engine_runtime() {
  local runtime=$1 owner engine_owner answer attempts
  local lock="$runtime/engine.lock"
  while ! mkdir "$lock" 2>/dev/null; do
    owner=$(cat "$lock/pid" 2>/dev/null || true)
    if [[ "$owner" =~ ^[0-9]+$ ]] && [ "$owner" -gt 1 ] && ! runtime_pid_running "$owner"; then
      engine_owner=$(cat "$runtime/engine-process.lock" 2>/dev/null || true)
      if runtime_pid_running "$engine_owner"; then
        echo "doc.haus: the engine is still running (pid $engine_owner), but its launcher has exited. Stop that engine before restarting." >&2
        return 1
      fi
      echo "doc.haus: clearing a stale launcher lock" >&2
      rm -f "$lock/pid"
      rmdir "$lock" || return 1
      continue
    fi
    if ! runtime_launcher "$owner"; then
      echo "doc.haus: another process owns this runtime ($lock); it could not be identified as this stack's launcher." >&2
      return 1
    fi
    echo "doc.haus: already running (launcher pid $owner)." >&2
    if [ ! -t 0 ]; then
      echo "doc.haus: leaving it running. Run ./start.sh in a terminal to choose whether to restart it." >&2
      return 2
    fi
    printf 'Stop that stack and restart it? [y/N] ' >&2
    read -r answer || answer=
    case "$answer" in
      y|Y|yes|YES|Yes) ;;
      *) echo "doc.haus: leaving the existing stack running." >&2; return 2 ;;
    esac
    # Re-check ownership after the answer; never signal a replacement owner.
    if [ "$(cat "$lock/pid" 2>/dev/null || true)" != "$owner" ] || ! runtime_launcher "$owner"; then
      echo "doc.haus: runtime ownership changed; run ./start.sh again." >&2
      return 1
    fi
    echo "doc.haus: stopping the existing stack..." >&2
    engine_owner=$(cat "$runtime/engine-process.lock" 2>/dev/null || true)
    kill -TERM "$owner" || return 1
    attempts=0
    while { runtime_pid_running "$owner" || runtime_pid_running "$engine_owner"; } && [ "$attempts" -lt 60 ]; do
      sleep 0.25
      attempts=$((attempts + 1))
    done
    if runtime_pid_running "$owner" || runtime_pid_running "$engine_owner"; then
      echo "doc.haus: the existing stack has not stopped yet. Retry once shutdown finishes." >&2
      return 1
    fi
  done
  engine_owner=$(cat "$runtime/engine-process.lock" 2>/dev/null || true)
  if runtime_pid_running "$engine_owner"; then
    rmdir "$lock"
    echo "doc.haus: the engine is still running (pid $engine_owner), but its launcher has exited. Stop that engine before restarting." >&2
    return 1
  fi
  # The engine uses exclusive file creation, so a dead PID's record must also
  # be removed before launch. Preserve malformed or changed ownership records.
  if [ -e "$runtime/engine-process.lock" ]; then
    if ! [[ "$engine_owner" =~ ^[0-9]+$ ]] || [ "$engine_owner" -le 1 ] ||
      [ "$(cat "$runtime/engine-process.lock" 2>/dev/null || true)" != "$engine_owner" ]; then
      rmdir "$lock"
      echo "doc.haus: engine ownership could not be verified ($runtime/engine-process.lock)." >&2
      return 1
    fi
    echo "doc.haus: clearing a stale engine lock (pid $engine_owner)" >&2
    if ! rm "$runtime/engine-process.lock"; then
      rmdir "$lock"
      return 1
    fi
  fi
  echo "$$" > "$lock/pid"
}

runtime_pid_running() {
  [[ "$1" =~ ^[0-9]+$ ]] && [ "$1" -gt 1 ] && kill -0 "$1" 2>/dev/null
}

runtime_launcher() {
  local owner=$1 command directory
  runtime_pid_running "$owner" || return 1
  command=$(ps -p "$owner" -o args= 2>/dev/null || true)
  [[ "$command" =~ (^|[[:space:]/])start\.sh([[:space:]]|$) ]] || return 1
  directory=$(lsof -a -p "$owner" -d cwd -Fn 2>/dev/null | sed -n 's/^n//p')
  [ "$directory" = "$(pwd -P)" ]
}
