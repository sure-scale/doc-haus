import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import os from "node:os"
import { startWordHttps, wordHttpsFailure, wordHttpsHandler } from "./word-https"
import { wordConfiguration } from "./word-documents"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { WORD_BOOTSTRAP_ORIGIN, wordBootstrapUrl } from "../../../apps/word-addin/src/bootstrap-contract"

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
  expect((await handler(new Request("https://localhost:4511/word/lease"))).status).toBe(405)
  expect((await handler(new Request("https://localhost:4511/word/lease", { method: "POST", body: "{}" }))).status).toBe(200)
  expect(forwarded).toEqual(["/word/saves", "/word/lease"])
  expect(await (await handler(new Request("https://localhost:4511/word-addin/pane.html"))).text()).toContain("Word pane")
  expect((await handler(new Request("https://localhost:4511/word-addin/%2fetc%2fpasswd"))).status).toBe(404)
  const manifest = await handler(new Request("https://localhost:4511/word-addin/manifest.xml"))
  expect(manifest.headers.get("Content-Disposition")).toContain("attachment")
  expect(await manifest.text()).toContain(wordBootstrapUrl("https://localhost:4511"))
  const handoff = `/word/open/${"a".repeat(43)}`
  expect((await handler(new Request(`https://localhost:4511${handoff}`))).status).toBe(200)
  expect(forwarded.at(-1)).toBe(handoff)
  expect((await handler(new Request(`https://localhost:4511${handoff}`, { headers: { Origin: "https://unrelated.example" } }))).status).toBe(403)
  expect((await handler(new Request(`https://localhost:4511${handoff}`, { method: "POST" }))).status).toBe(405)
  expect((await handler(new Request(`https://localhost:4511${handoff}/extra`))).status).toBe(404)
  expect((await handler(new Request(`https://localhost:4511${handoff}/result`, { method: "POST" }))).status).toBe(200)
  expect((await handler(new Request(`https://localhost:4511${handoff}/result`))).status).toBe(405)
})

test("bootstrap health is read-only, reports actual assets and limits cross-origin access", async () => {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-health-"))
  roots.push(directory)
  let forwarded = 0
  const handler = wordHttpsHandler(() => { forwarded++; return new Response("private") },
    { origin: "https://localhost:4511", external: false, host: "127.0.0.1", port: 4511 }, directory)
  const request = (method = "GET", origin: string | undefined = WORD_BOOTSTRAP_ORIGIN) =>
    new Request("https://localhost:4511/word-addin/health", { method, ...(origin ? { headers: { Origin: origin } } : {}) })
  expect((await handler(request())).status).toBe(503)
  for (const file of ["pane.html", "pane.js", "pane.css", "desk.html", "desk.js", "desk.css", "assets/icon.png"]) {
    mkdirSync(path.dirname(path.join(directory, file)), { recursive: true })
    writeFileSync(path.join(directory, file), "fixture")
  }
  const ready = await handler(request())
  expect(ready.status).toBe(200)
  expect(await ready.json()).toEqual({ ready: true, version: WORD_ADDIN_VERSION })
  expect(ready.headers.get("Access-Control-Allow-Origin")).toBe(WORD_BOOTSTRAP_ORIGIN)
  expect(ready.headers.get("Cache-Control")).toBe("no-store")
  expect(ready.headers.get("Access-Control-Allow-Credentials")).toBeNull()
  expect((await handler(request("HEAD"))).status).toBe(200)
  expect((await handler(request("OPTIONS"))).status).toBe(204)
  expect((await handler(request("POST"))).status).toBe(405)
  expect((await handler(request("GET", "https://unrelated.example"))).status).toBe(403)
  const native = await handler(new Request("https://localhost:4511/word-addin/health"))
  expect(native.status).toBe(200)
  expect(native.headers.get("Access-Control-Allow-Origin")).toBeNull()
  expect(forwarded).toBe(0)
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
