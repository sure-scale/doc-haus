import os from "node:os"
import path from "node:path"

// Where the local embedding model's weights are cached. The ingest service
// (services/ingest/src/embed.ts) and the search-document tool
// (query-embedding.ts) load the same model, and its bytes are identical across
// checkouts, so both share one per-user cache instead of the transformers.js
// default inside each package's node_modules — a fresh install or a new
// worktree then reuses the download rather than fetching it again.
// DOCHAUS_MODEL_CACHE_DIR overrides the location (for example a mounted volume).
// Caches left under node_modules by earlier versions are not migrated; the
// model downloads once into this location.
export function embeddingModelCacheDir() {
  return process.env.DOCHAUS_MODEL_CACHE_DIR || path.join(os.homedir(), ".cache", "doc-haus", "models")
}
