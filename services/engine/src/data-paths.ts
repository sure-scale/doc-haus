import path from "node:path"

// Opt-in layout; existing installations keep their paths until deliberately
// migrated. Explicit workspace/runtime overrides continue to take precedence.
export function dataPaths(
  repository: string,
  environment: Record<string, string | undefined> = process.env,
  legacyWorkspace = path.join(repository, "workspace"),
) {
  const root = environment.DOCHAUS_DATA_ROOT ? path.resolve(repository, environment.DOCHAUS_DATA_ROOT) : undefined
  return {
    workspace: path.resolve(
      repository,
      environment.WORKSPACE_ROOT || (root ? path.join(root, "workspace") : legacyWorkspace),
    ),
    runtime: path.resolve(
      repository,
      environment.DOCHAUS_RUNTIME_DIR ||
        (root ? path.join(root, "runtime") : path.join(repository, ".dochaus-runtime")),
    ),
  }
}
