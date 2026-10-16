import { Database } from "bun:sqlite"
import { mkdir, rename, rm } from "node:fs/promises"
import path from "node:path"
import { getMeta, setMeta } from "./db"
import { embeddingModelCacheDir } from "../../../dochaus/lib/search/model-cache"

// Local embedding model, shared shape with dochaus/lib/search/query-embedding.ts.
// Both sides MUST use the same model id and pooling so query and chunk vectors
// are comparable. granite-small-r2 replaces MiniLM-L6: MiniLM's
// 256-token window silently truncated half of every 2000-char chunk; granite
// reads 8192 tokens and scores ~14 points higher on MTEB retrieval at the same
// 384 dims, so stored vectors keep their schema. The model card specifies CLS
// pooling, not mean.
//
// Multilingual limitation: this is an English embedding model ("...-english-...").
// Retrieval quality degrades for non-English matters (civil-law jurisdictions,
// cross-border documents with non-English text). Supporting a multilingual
// matter requires a coordinated model swap here AND in
// dochaus/lib/search/query-embedding.ts (same MODEL id, same DIM, same pooling), plus
// a re-embed of every affected matter — migrateEmbeddings() handles the re-embed
// when embedding_model meta changes, but the model id and dims must match on
// both sides first or search-document refuses the index.

export const MODEL = "onnx-community/granite-embedding-small-english-r2-ONNX"
export const DIM = 384

// The model could not be loaded (for example its first download from
// huggingface.co failed), so nothing was embedded. The message is safe to show
// the lawyer: it names the model and the network cause, never document text.
export class EmbeddingModelUnavailableError extends Error {}

let extractor: ReturnType<typeof loadExtractor> | undefined

// One load per process, shared by every concurrent embed so a many-chunk
// document never starts parallel downloads. A failed load clears the cached
// promise, so the next upload retries instead of replaying the old failure.
export function loadEmbeddingModel() {
  extractor ??= loadExtractor().catch((error: unknown) => {
    extractor = undefined
    throw new EmbeddingModelUnavailableError(
      `The embedding model ${MODEL} could not be loaded (${
        error instanceof Error ? error.message : String(error)
      }). It downloads once from huggingface.co into ${embeddingModelCacheDir()}; check the network connection and try again.`,
      { cause: error },
    )
  })
  return extractor
}

async function loadExtractor() {
  const transformers = await import("@huggingface/transformers")
  const env = transformers.env
  env.cacheDir = embeddingModelCacheDir()
  // transformers.js 4.2 never settles a pipeline load whose external-data
  // (.onnx_data) download fails: the error is thrown inside a Promise executor,
  // so the load stays pending forever and the rejection goes unhandled. Fetch
  // any uncached model file first, one at a time, so a network failure rejects
  // this load cleanly; the pipeline then reads its weights from the cache.
  const cached = await transformers.ModelRegistry.is_pipeline_cached_files("feature-extraction", MODEL, { dtype: "q8" })
  for (const entry of cached.files.filter((file) => !file.cached)) await downloadModelFile(env, entry.file)
  return transformers.pipeline("feature-extraction", MODEL, { dtype: "q8" })
}

// Writes to the same cache path transformers.js reads, through a temporary
// name in its "<file>.tmp.<pid>.<random>" form and an atomic rename, so an
// interrupted download is never mistaken for a complete file.
async function downloadModelFile(env: { remoteHost: string; remotePathTemplate: string; cacheDir: string | null }, file: string) {
  const response = await fetch(
    new URL(env.remotePathTemplate.replaceAll("{model}", MODEL).replaceAll("{revision}", "main") + file, env.remoteHost),
  )
  if (!response.ok) throw new Error(`downloading ${file} returned HTTP ${response.status}`)
  const target = path.join(env.cacheDir ?? embeddingModelCacheDir(), MODEL, file)
  const partial = `${target}.tmp.${process.pid}.${Math.random().toString(36).slice(2)}`
  await mkdir(path.dirname(target), { recursive: true })
  try {
    await Bun.write(partial, response)
    await rename(partial, target)
  } finally {
    await rm(partial, { force: true })
  }
}

export async function embed(text: string): Promise<Float32Array> {
  const output = await (await loadEmbeddingModel())(text, { pooling: "cls", normalize: true })
  return output.data as Float32Array
}

// Chunk-side embedding text: a document › structural-path breadcrumb prepended
// to the chunk body. On matters where dozens of contracts share near-identical
// boilerplate, the breadcrumb is what lets the vector channel keep clauses from
// different documents apart, and the full path ("§7 Indemnification › §7.2 Cap")
// situates a clause within its parent better than the bare section label.
// Queries embed raw — only the document side carries context. The breadcrumb
// exists solely in the vector; stored chunk text stays the verbatim document text.
export function embedChunk(docName: string, breadcrumb: string, text: string) {
  return embed(`${docName} › ${breadcrumb}\n${text}`)
}

// Re-embed every chunk written by an older embedding model. Vectors from two
// models are not comparable, so a stale matter would silently return garbage
// rankings; this runs at service boot (server.ts walks all matters) and before
// any ingest into an existing db, so a matter is never served mixed vectors.
export async function migrateEmbeddings(db: Database): Promise<number> {
  if (getMeta(db, "embedding_model") === MODEL) return 0
  const rows = db.query("SELECT id, doc_name, section, structural_path, text FROM chunks").all() as Array<{
    id: number
    doc_name: string
    section: string
    structural_path: string
    text: string
  }>
  // Embed first, write after: bun:sqlite transactions are synchronous, so all
  // awaits happen before the write. The transaction makes the swap atomic — a
  // crash mid-migration leaves the old vectors and old meta intact rather than
  // a mixed-model db (which search-document would refuse, but never serve).
  // Re-embed on the structural path when one exists (chunks written by the
  // structure-aware ingest), falling back to the bare section label for chunks
  // that predate it so the breadcrumb still matches how they were first embedded.
  const vectors = await Promise.all(
    rows.map((row) => embedChunk(row.doc_name, row.structural_path || row.section, row.text)),
  )
  db.transaction(() => {
    rows.forEach((row, i) => {
      const vector = vectors[i]!
      db.run("UPDATE chunks SET embedding = ? WHERE id = ?", [
        Buffer.from(vector.buffer, vector.byteOffset, vector.byteLength),
        row.id,
      ])
    })
    setMeta(db, "embedding_model", MODEL)
  })()
  return rows.length
}
