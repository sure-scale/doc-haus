import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import os from "node:os"
import { startWordHttps, wordHttpsFailure, wordHttpsHandler } from "./word-https"
import { wordConfiguration } from "./word-documents"

const roots: string[] = []
afterEach(() => { for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true }) })

test("Word HTTPS surface forwards capability writes but cannot reach normal document or admin APIs", async () => {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-pane-"))
  roots.push(directory)
  writeFileSync(path.join(directory, "pane.html"), "<p>Word pane</p>")
  const forwarded: string[] = []
  const handler = wordHttpsHandler(request => { forwarded.push(new URL(request.url).pathname); return new Response("write") },
    { origin: "https://localhost:4511", external: false, host: "127.0.0.1", port: 4511 }, directory)
  for (const route of ["/matters", "/settings", "/matters/test/documents/content", "/word/saves/extra"])
    expect((await handler(new Request(`https://localhost:4511${route}`))).status).toBe(404)
  expect((await handler(new Request("https://localhost:4511/word/saves"))).status).toBe(405)
  expect(forwarded).toEqual([])
  expect((await handler(new Request("https://localhost:4511/word/saves", { method: "POST", body: "docx" }))).status).toBe(200)
  expect(forwarded).toEqual(["/word/saves"])
  expect(await (await handler(new Request("https://localhost:4511/word-addin/pane.html"))).text()).toContain("Word pane")
  expect((await handler(new Request("https://localhost:4511/word-addin/%2fetc%2fpasswd"))).status).toBe(404)
  const manifest = await handler(new Request("https://localhost:4511/word-addin/manifest.xml"))
  expect(manifest.headers.get("Content-Disposition")).toContain("attachment")
  expect(await manifest.text()).toContain("https://localhost:4511/word-addin/pane.html")
})

test("bad bridge TLS is recoverable and config reports unavailable without private paths", () => {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-bad-tls-"))
  roots.push(directory)
  const keys = ["DOCHAUS_DATA_ROOT", "DOCHAUS_RUNTIME_DIR", "DOCHAUS_WORD_ORIGIN", "DOCHAUS_WORD_TLS_CERT", "DOCHAUS_WORD_TLS_KEY", "DOCHAUS_WORD_EXTERNAL_HTTPS"]
  const prior = Object.fromEntries(keys.map(key => [key, process.env[key]]))
  const cert = path.join(directory, "private-certificate.crt"), key = path.join(directory, "private-key.key")
  writeFileSync(cert, "not a certificate")
  writeFileSync(key, "not a private key")
  try {
    process.env.DOCHAUS_DATA_ROOT = directory
    delete process.env.DOCHAUS_RUNTIME_DIR
    process.env.DOCHAUS_WORD_ORIGIN = "https://localhost:47613"
    process.env.DOCHAUS_WORD_TLS_CERT = cert
    process.env.DOCHAUS_WORD_TLS_KEY = key
    delete process.env.DOCHAUS_WORD_EXTERNAL_HTTPS
    expect(() => startWordHttps(() => new Response())).not.toThrow()
    expect(wordHttpsFailure()).toContain("could not start")
    expect(wordConfiguration()).toMatchObject({ available: false, origin: null, manifestUrl: null })
    expect(JSON.stringify(wordConfiguration())).not.toContain(directory)
    process.env.DOCHAUS_WORD_EXTERNAL_HTTPS = "1"
    startWordHttps(() => new Response())
    expect(wordHttpsFailure()).toBeUndefined()
    expect(wordConfiguration().available).toBe(true)
  } finally {
    for (const envKey of keys) {
      if (prior[envKey] === undefined) delete process.env[envKey]
      else process.env[envKey] = prior[envKey]
    }
  }
})
