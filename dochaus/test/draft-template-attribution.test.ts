import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"

test("template drafting stamps the new document and leaves template bytes untouched", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dh-template-attribution-"))
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify({ attorney: "Template Reviewing Lawyer", firm: "Template Reviewing Firm" }))
  writeFileSync(path.join(root, "matter.json"), JSON.stringify({ id: "template-test" }))
  const template = await Bun.file(path.resolve(import.meta.dir, "../templates/_base.docx")).bytes()
  const snapshot = new Uint8Array(template)
  const server = Bun.serve({ port: 0, async fetch(request) {
    if (new URL(request.url).pathname === "/templates/content") return new Response(new Uint8Array(template))
    const file = (await request.formData()).get("file") as File
    writeFileSync(path.join(root, file.name), new Uint8Array(await file.arrayBuffer()))
    return Response.json({})
  } })
  try {
    // Isolate import-time URLs so this test cannot change another test's tool
    // transport or contact a configured stack.
    const child = Bun.spawn([process.execPath, "-e", `
      import draft from ${JSON.stringify(path.resolve(import.meta.dir, "../tool/draft-document.ts"))};
      const directory = process.env.WORKSPACE_ROOT;
      const result = await draft.execute({name:"prepared.docx",template:"review.docx"}, {
        directory,worktree:directory,sessionID:"template-identity",messageID:"template-identity",agent:"draft",
        abort:new AbortController().signal,metadata() {},async ask() {}
      });
      if (typeof result === "string") throw new Error(result);
    `], { env: { ...process.env, WORKSPACE_ROOT: root, INGEST_URL: server.url.href.replace(/\/$/, "") }, stdout: "pipe", stderr: "pipe" })
    const error = await new Response(child.stderr).text()
    expect(await child.exited, error).toBe(0)
    const parts = unzipSync(readFileSync(path.join(root, "prepared.docx")))
    const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
    expect(parse(parts["docProps/core.xml"]).getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "creator")[0]?.textContent).toBe("Template Reviewing Lawyer")
    expect(parse(parts["docProps/app.xml"]).getElementsByTagNameNS("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")[0]?.textContent).toBe("Template Reviewing Firm")
    expect(template).toEqual(snapshot)
  } finally {
    server.stop(true)
    rmSync(root, { recursive: true, force: true })
  }
}, 120_000)
