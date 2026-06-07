// doc.haus issue #1 spike: validate Docxodus headless under Bun + offset alignment.
// Throwaway scratch. try/catch per phase so one failure still yields a full report.
import { readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { execSync } from "node:child_process"
import mammoth from "mammoth"
import * as dx from "docxodus"

const WASM = fileURLToPath(new URL("./node_modules/docxodus/dist/wasm/", import.meta.url))
const log = (...a) => console.log(...a)
const ok = (q, v) => log(`\n## ${q}\n${v}`)
const now = () => Number(process.hrtime.bigint() / 1000000n)

const report = {}

// --- Q1: load headless under Bun + cold-start ---
try {
  const t0 = now()
  await dx.initialize(WASM)
  const cold = now() - t0
  report.init = { ok: dx.isInitialized(), coldMs: cold, version: dx.getVersion?.() }
  ok("Q1 initialize() under Bun", `initialized=${dx.isInitialized()} coldStart=${cold}ms version=${JSON.stringify(dx.getVersion?.())}`)
} catch (e) {
  report.init = { ok: false, error: String(e?.stack || e) }
  ok("Q1 initialize() under Bun", `FAILED: ${e?.message || e}`)
  console.log(e?.stack)
  // hard gate: nothing else can run
  console.log("\n=== REPORT ===\n" + JSON.stringify(report, null, 2))
  process.exit(0)
}

const sample = readFileSync("sample.docx")

// --- sanity: convert to html (proves parse path) ---
try {
  const t0 = now()
  const html = await dx.convertDocxToHtml(new Uint8Array(sample))
  report.html = { ok: true, ms: now() - t0, len: html.length }
  ok("convertDocxToHtml", `len=${html.length} in ${report.html.ms}ms`)
} catch (e) {
  report.html = { ok: false, error: String(e?.message || e) }
  ok("convertDocxToHtml", `FAILED: ${e?.message || e}`)
}

// --- Q4 part A: docxodus text projection vs mammoth plaintext ---
let mammothText = ""
try {
  mammothText = (await mammoth.extractRawText({ buffer: sample })).value
  ok("mammoth.extractRawText (our ingest offset basis)", JSON.stringify(mammothText.slice(0, 200)) + ` ... total ${mammothText.length} chars`)
} catch (e) {
  ok("mammoth", `FAILED: ${e?.message || e}`)
}

let session
try {
  session = dx.openDocxSession(new Uint8Array(sample), { trackedChanges: "render_inline", revisionAuthor: "doc.haus" })
  const proj = session.project()
  const keys = Object.keys(proj || {})
  ok("Q4 DocxSession.project() shape", `top-level keys: ${keys.join(", ")}`)
  // try to surface anchors + their text + any char offsets
  const blocks = proj.blocks || proj.elements || proj.children || []
  log(`block count: ${Array.isArray(blocks) ? blocks.length : "n/a (not array)"}`)
  if (Array.isArray(blocks)) {
    for (const b of blocks.slice(0, 6)) {
      log("  block:", JSON.stringify({ id: b.anchorId || b.id, kind: b.kind, text: (b.text || b.markdown || "").slice(0, 60) }))
    }
  } else {
    log("proj sample:", JSON.stringify(proj).slice(0, 600))
  }
  report.project = { keys, blockCount: Array.isArray(blocks) ? blocks.length : null }
} catch (e) {
  report.project = { ok: false, error: String(e?.stack || e) }
  ok("Q4 project()", `FAILED: ${e?.message || e}`)
  console.log(e?.stack)
}

// --- searchTextOffsets: docxodus global offsets vs mammoth offsets ---
try {
  const needle = "thirty (30) days"
  const spans = await dx.searchTextOffsets(new Uint8Array(sample), needle, 5)
  const mammothIdx = mammothText.indexOf(needle)
  ok("Q4 offset alignment probe",
    `needle=${JSON.stringify(needle)}\n  docxodus searchTextOffsets -> ${JSON.stringify(spans)}\n  mammoth char index    -> ${mammothIdx}`)
  report.offsets = { needle, docxodus: spans, mammoth: mammothIdx, aligned: spans?.[0]?.start === mammothIdx }
} catch (e) {
  report.offsets = { ok: false, error: String(e?.message || e) }
  ok("Q4 offset alignment probe", `searchTextOffsets FAILED: ${e?.message || e}`)
}

// --- Q3 part A: find a clause by TEXT (re-anchoring path), edit, save ---
let editedBytes
try {
  const target = session.findByText("for convenience upon thirty (30) days")
  ok("Q3 findByText (text re-anchoring)", JSON.stringify(target))
  const anchorId = target?.id || target?.anchorId
  if (anchorId) {
    const res = session.replaceText(anchorId, "Either party may terminate this Agreement for convenience upon sixty (60) days written notice to the other party.")
    ok("Q3 replaceText EditResult", JSON.stringify({ success: res.success, error: res.error, modified: res.modified }))
    editedBytes = session.save()
    writeFileSync("edited.docx", editedBytes)
    log(`saved edited.docx (${editedBytes.length} bytes)`)
    report.edit = { ok: res.success }
  }
} catch (e) {
  report.edit = { ok: false, error: String(e?.stack || e) }
  ok("Q3 edit", `FAILED: ${e?.message || e}`)
  console.log(e?.stack)
}

// --- Q2/Q3: produce native w:ins/w:del via compareDocuments redline ---
try {
  if (editedBytes) {
    const redline = await dx.compareDocuments(new Uint8Array(sample), new Uint8Array(editedBytes), { author: "doc.haus" })
    writeFileSync("redline.docx", redline)
    // unzip word/document.xml and grep for tracked-change elements
    execSync("rm -rf _x && mkdir _x && cd _x && unzip -o ../redline.docx word/document.xml >/dev/null", { shell: "/bin/zsh" })
    const xml = readFileSync("_x/word/document.xml", "utf8")
    const ins = (xml.match(/<w:ins[ >]/g) || []).length
    const del = (xml.match(/<w:del[ >]/g) || []).length
    ok("Q2/Q3 native tracked changes (compareDocuments)", `redline.docx written. <w:ins>=${ins}  <w:del>=${del}`)
    report.redline = { ok: ins > 0 || del > 0, ins, del }
  }
} catch (e) {
  report.redline = { ok: false, error: String(e?.message || e) }
  ok("Q2/Q3 redline", `FAILED: ${e?.message || e}`)
}

try { session?.close() } catch {}

console.log("\n=== REPORT (json) ===\n" + JSON.stringify(report, null, 2))
