// Render upload boundary guarantees. The renderer delegates writes to ingest, but
// the drafting matter is the scoped directory; a 200 response is not enough if
// the file is not visible there afterward.
// Run from dochaus/: bun test test/render-upload.test.ts
import { describe, expect, test } from "bun:test"
import { existsSync, mkdtempSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

describe("uploadDeliverable", () => {
  test("reports a successful ingest response that did not write into the scoped matter", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "render-upload-matter-"))
    writeFileSync(path.join(dir, "matter.json"), JSON.stringify({ id: "render-upload-test" }))

    const server = Bun.serve({
      port: 0,
      async fetch(req) {
        expect(new URL(req.url).pathname).toBe("/matters/render-upload-test/deliverables")
        await req.formData()
        return Response.json({ ok: true })
      },
    })
    const originalFetch = globalThis.fetch
    globalThis.fetch = ((input, init) => {
      const original = new URL(input instanceof Request ? input.url : String(input))
      return originalFetch(new URL(original.pathname, server.url), init)
    }) as typeof fetch

    try {
      const { uploadDeliverable } = await import("../lib/render")
      const result = await uploadDeliverable(dir, "agreement.docx", new Uint8Array([80, 75, 3, 4]))

      expect(result).toContain("did not write agreement.docx")
      expect(existsSync(path.join(dir, "agreement.docx"))).toBe(false)
    } finally {
      server.stop(true)
      globalThis.fetch = originalFetch
    }
  })
})
