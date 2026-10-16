// Query-side embedding for the search-document tool.
//
// Must stay in lockstep with services/ingest/src/embed.ts (same model id and
// pooling), or query and chunk vectors stop being comparable. The model card
// specifies CLS pooling. Queries embed raw; chunks were embedded with a
// document › section breadcrumb prepended (see embed.ts), which only the
// document side carries.
//
// Multilingual limitation: this is an English embedding model, so retrieval
// quality degrades on non-English matters. A coordinated model swap (same id,
// dim, and pooling here and in embed.ts) plus a matter re-embed is required to
// serve multijurisdictional non-English work; see embed.ts for the migration
// notes. The model-refuse guard in tool/search-document.ts catches a stale
// mismatch but not a partial unilateral swap, so both files must change together.
import { embeddingModelCacheDir } from "./model-cache"

export const MODEL = "onnx-community/granite-embedding-small-english-r2-ONNX"
export const DIM = 384

// One shared load per process; a failed load clears the cached promise so the
// next search retries. The cache is shared with ingest (see model-cache.ts), so
// a stack launched by start.sh finds the model already downloaded.
let extractor: ReturnType<typeof loadExtractor> | undefined
export async function embed(text: string) {
  extractor ??= loadExtractor().catch((error: unknown) => {
    extractor = undefined
    throw error
  })
  const output = await (await extractor)(text, { pooling: "cls", normalize: true })
  return output.data as Float32Array
}

async function loadExtractor() {
  const { env, pipeline } = await import("@huggingface/transformers")
  env.cacheDir = embeddingModelCacheDir()
  return pipeline("feature-extraction", MODEL, { dtype: "q8" })
}

export function dot(a: Float32Array, b: Float32Array) {
  let sum = 0
  for (let i = 0; i < DIM; i++) sum += a[i] * b[i]
  return sum
}
