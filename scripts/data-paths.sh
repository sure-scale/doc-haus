# Sourced by the launcher; keep precedence aligned with services/engine/src/data-paths.ts.
configure_data_paths() {
  local repository="$1"
  if [ -n "${DOCHAUS_DATA_ROOT:-}" ]; then
    [[ "$DOCHAUS_DATA_ROOT" = /* ]] || DOCHAUS_DATA_ROOT="$repository/$DOCHAUS_DATA_ROOT"
    export DOCHAUS_DATA_ROOT
    WORKSPACE_ROOT="${WORKSPACE_ROOT:-$DOCHAUS_DATA_ROOT/workspace}"
    DOCHAUS_RUNTIME_DIR="${DOCHAUS_RUNTIME_DIR:-$DOCHAUS_DATA_ROOT/runtime}"
  fi
  WORKSPACE_ROOT="${WORKSPACE_ROOT:-$repository/workspace}"
  DOCHAUS_RUNTIME_DIR="${DOCHAUS_RUNTIME_DIR:-$repository/.dochaus-runtime}"
  [[ "$WORKSPACE_ROOT" = /* ]] || WORKSPACE_ROOT="$repository/$WORKSPACE_ROOT"
  [[ "$DOCHAUS_RUNTIME_DIR" = /* ]] || DOCHAUS_RUNTIME_DIR="$repository/$DOCHAUS_RUNTIME_DIR"
  export WORKSPACE_ROOT DOCHAUS_RUNTIME_DIR
}
