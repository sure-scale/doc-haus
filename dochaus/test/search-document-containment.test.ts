import { expect, test } from "bun:test"
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, renameSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { insertChunk, openDb, setMeta, upsertDocument } from "../../services/ingest/src/db"

// Keep the embedding stub in a subprocess: it must never replace the model
// module in another test file. SQLite, the tool, and its citation path are real.
async function search(directory: string) {
  const child = Bun.spawn([process.execPath, "--no-env-file", "-e", `
    import { Database } from "bun:sqlite"
    import { mock, spyOn } from "bun:test"
    let embeddings = 0
    mock.module("@huggingface/transformers", () => ({
      env: {},
      pipeline: async () => async () => {
        embeddings++
        const data = new Float32Array(384)
        data[0] = 1
        return { data }
      },
    }))
    const searchDocument = (await import("../tool/search-document")).default
    const reads = spyOn(Database.prototype, "query")
    const input = JSON.parse(await Bun.stdin.text())
    const result = await searchDocument.execute({ query: "scope canary" }, {
      directory: input.directory,
      sessionID: crypto.randomUUID(),
    })
    console.log(JSON.stringify({ result, embeddings, reads: reads.mock.calls.length }))
  `], {
    cwd: import.meta.dir,
    env: { PATH: process.env.PATH ?? "" },
    stdin: "pipe",
    stdout: "pipe",
    stderr: "pipe",
  })
  child.stdin.write(JSON.stringify({ directory }))
  child.stdin.end()
  const [stdout, stderr, code] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ])
  expect({ code, stderr }).toEqual({ code: 0, stderr: "" })
  return JSON.parse(stdout) as { result: string | { output: string }; embeddings: number; reads: number }
}

function index(directory: string, canary: string, standalone = false) {
  mkdirSync(directory, { recursive: true })
  const document = path.join(directory, "source.docx")
  writeFileSync(document, "synthetic source")
  const db = openDb(directory)
  const documentId = upsertDocument(db, document, "source.docx", 1, null)
  setMeta(db, "embedding_model", "onnx-community/granite-embedding-small-english-r2-ONNX")
  const text = `scope canary ${canary}`
  const embedding = new Float32Array(384)
  embedding[0] = 1
  insertChunk(db, {
    documentId,
    docPath: document,
    docName: "source.docx",
    section: "1",
    structuralPath: "Doc(source.docx) › §1 Source",
    party: "",
    chunkIndex: 0,
    text,
    charStart: 0,
    charEnd: text.length,
    embedding,
    flagged: false,
  })
  // A copied/renamed main file must be a complete SQLite snapshot, rather than
  // a WAL database whose current schema and rows still live in its sidecar.
  if (standalone) db.query("PRAGMA journal_mode = DELETE").get()
  db.close()
  return path.join(directory, ".dochaus", "legal.db")
}

test.each(["metadata directory", "database file"])("search rejects an out-of-matter %s alias before embedding or SQL reads", async (kind) => {
  const root = mkdtempSync(path.join(tmpdir(), "search-containment-"))
  const active = path.join(root, "matter")
  // A lexical prefix of the active path is not a physical descendant.
  const sibling = path.join(root, "matter-private")
  const canary = `sibling-${crypto.randomUUID()}`
  const localDb = index(active, "local")
  const siblingDb = index(sibling, canary)
  try {
    if (kind === "metadata directory") {
      rmSync(path.dirname(localDb), { recursive: true })
      symlinkSync(path.dirname(siblingDb), path.dirname(localDb))
    }
    if (kind === "database file") {
      rmSync(localDb)
      symlinkSync(siblingDb, localDb)
    }
    const denied = await search(active)
    expect(denied.result).toBe("This matter's search index is outside its matter directory. Search was not performed.")
    expect(denied.embeddings).toBe(0)
    expect(denied.reads).toBe(0)
    expect(JSON.stringify(denied.result)).not.toContain(canary)
    expect(JSON.stringify(denied.result)).not.toContain(sibling)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

test("a byte-identical database outside the matter is still denied through an alias", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "search-identical-index-"))
  const active = path.join(root, "matter")
  const localDb = index(active, "local", true)
  const outsideDb = path.join(root, "identical-matter", ".dochaus", "legal.db")
  try {
    mkdirSync(path.dirname(outsideDb), { recursive: true })
    copyFileSync(localDb, outsideDb)
    expect(readFileSync(outsideDb)).toEqual(readFileSync(localDb))
    rmSync(localDb)
    symlinkSync(outsideDb, localDb)
    const denied = await search(active)
    expect(denied.result).toBe("This matter's search index is outside its matter directory. Search was not performed.")
    expect(denied.embeddings).toBe(0)
    expect(denied.reads).toBe(0)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

test("an outside non-database file is denied without parsing it or embedding the query", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "search-invalid-index-"))
  const active = path.join(root, "matter")
  mkdirSync(path.join(active, ".dochaus"), { recursive: true })
  const outside = path.join(root, "not-a-database")
  writeFileSync(outside, `not sqlite ${crypto.randomUUID()}`)
  symlinkSync(outside, path.join(active, ".dochaus", "legal.db"))
  try {
    const denied = await search(active)
    expect(denied.result).toBe("This matter's search index is outside its matter directory. Search was not performed.")
    expect(denied.embeddings).toBe(0)
    expect(denied.reads).toBe(0)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

test.each(["regular index", "internal database alias", "internal metadata alias", "matter alias"])("search preserves a contained %s", async (kind) => {
  const root = mkdtempSync(path.join(tmpdir(), "search-contained-index-"))
  const active = path.join(root, "matter")
  const canary = `permitted-${crypto.randomUUID()}`
  const localDb = index(active, canary, kind === "internal database alias")
  const matterAlias = path.join(root, "matter-alias")
  try {
    if (kind === "internal database alias") {
      mkdirSync(path.join(path.dirname(localDb), "stored-index"))
      renameSync(localDb, path.join(path.dirname(localDb), "stored-index", "legal.db"))
      symlinkSync("stored-index/legal.db", localDb)
    }
    if (kind === "internal metadata alias") {
      renameSync(path.dirname(localDb), path.join(active, ".stored-index"))
      symlinkSync(".stored-index", path.dirname(localDb))
    }
    if (kind === "matter alias") symlinkSync(active, matterAlias)
    const permitted = await search(kind === "matter alias" ? matterAlias : active)
    expect(typeof permitted.result).toBe("object")
    expect((permitted.result as { output: string }).output).toContain(canary)
    expect(permitted.embeddings).toBe(1)
    expect(permitted.reads).toBeGreaterThan(0)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

test("an unindexed matter retains its existing response without database or model access", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "search-no-index-"))
  try {
    const absent = await search(directory)
    expect(absent.result).toBe("No documents have been indexed for this matter yet.")
    expect(absent.embeddings).toBe(0)
    expect(absent.reads).toBe(0)
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})
