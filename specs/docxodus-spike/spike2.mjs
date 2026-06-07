import { readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import * as dx from "docxodus"
const WASM = fileURLToPath(new URL("./node_modules/docxodus/dist/wasm/", import.meta.url))
await dx.initialize(WASM)
const bytes = new Uint8Array(readFileSync("sample.docx"))

// PLAIN edit (no trackedChanges) — what #2 needs
const s = dx.openDocxSession(bytes, {})
const proj = s.project()
console.log("project keys:", Object.keys(proj), "markdown len:", proj.markdown?.length)
console.log("markdown head:", JSON.stringify(proj.markdown?.slice(0,140)))

const hit = s.findByText("thirty (30) days")
console.log("findByText:", hit?.id, hit?.kind)
const res = s.replaceTextRange(hit.id, "thirty (30) days", "ninety (90) days")
console.log("replaceTextRange results:", JSON.stringify(res?.map(r=>({ok:r.success,err:r.error}))))
writeFileSync("plain.docx", s.save())
s.close()
