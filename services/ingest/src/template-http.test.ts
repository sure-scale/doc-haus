import { expect, test } from "bun:test"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"

test("template upload initializes a fresh data root before any library read", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-template-http-"))
  const config = path.join(root, "config")
  await mkdir(config)
  await Bun.write(path.join(config, "opencode.json"), "{}")
  const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const child = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: {
      ...process.env,
      DOCHAUS_DATA_ROOT: root,
      WORKSPACE_ROOT: undefined,
      DOCHAUS_RUNTIME_DIR: undefined,
      DOCHAUS_DIR: config,
      INGEST_HOST: "127.0.0.1",
      INGEST_PORT: String(port),
    },
    stdout: Bun.file(path.join(root, "ingest.log")),
    stderr: Bun.file(path.join(root, "ingest.log")),
  })
  const request = (route: string, options?: RequestInit) =>
    fetch(`http://127.0.0.1:${port}${route}`, { ...options, signal: AbortSignal.timeout(15000) })
  try {
    const deadline = Date.now() + 20000
    while (!(await request("/matters").catch(() => undefined))?.ok) {
      if (Date.now() > deadline || child.exitCode !== null) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
    const bytes = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Template storage check")] }] }))
    const body = new FormData()
    body.set("file", new File([new Uint8Array(bytes)], "storage.docx"))
    body.set("description", "First template")
    const uploaded = await request("/templates", { method: "POST", body })
    expect(uploaded.status).toBe(200)
    expect(await uploaded.json()).toMatchObject({ name: "storage.docx", description: "First template" })
    const listed = await (await request("/templates")).json()
    expect(listed.dir).toBe(path.join(root, "workspace", ".templates"))
    expect(listed.templates).toMatchObject([{ name: "storage.docx", description: "First template" }])
    expect(new Uint8Array(await (await request("/templates/content?name=storage.docx")).arrayBuffer())).toEqual(new Uint8Array(bytes))
    expect((await (await request("/templates/text?name=storage.docx")).json()).text).toContain("Template storage check")
    expect((await request("/templates?name=storage.docx", {
      method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ description: "Revised template" }),
    })).status).toBe(200)
    expect((await request("/templates?name=storage.docx", { method: "DELETE" })).status).toBe(200)
    expect((await (await request("/templates")).json()).templates).toEqual([])
  } finally {
    child.kill("SIGTERM")
    await child.exited
    await rm(root, { recursive: true, force: true })
  }
}, 60000)
