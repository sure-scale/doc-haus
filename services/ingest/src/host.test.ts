import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { Hono } from "hono"
import { vertexProbeRoutes } from "./host"

test("legacy Vertex checks forward only saved provider/model through the native engine", async () => {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-native-vertex-"))
  const environment = ["DOCHAUS_DATA_ROOT", "DOCHAUS_RUNTIME_DIR", "DOCHAUS_DIR", "OPENCODE_URL", "OPENCODE_SERVER_PASSWORD"] as const
  const previous = Object.fromEntries(environment.map(key => [key, process.env[key]]))
  const config = path.join(directory, "runtime", "config", "opencode", "opencode.json")
  mkdirSync(path.dirname(config), { recursive: true })
  const saved = JSON.stringify({ provider: {
    "google-vertex": { options: { project: "fixture-project", location: "us-central1" } },
    "google-vertex-anthropic": { options: { project: "fixture-project" } },
  } })
  writeFileSync(config, saved)
  writeFileSync(path.join(directory, "opencode.json"), "{}")
  const password = "synthetic-native-vertex"
  let status = 200
  let payload: unknown = { ok: true, privateDiagnostic: "must not escape" }
  const requests: { path: string; authorization: string | null; body: unknown }[] = []
  const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, async fetch(request) {
    requests.push({ path: new URL(request.url).pathname, authorization: request.headers.get("Authorization"), body: await request.json() })
    return Response.json(payload, { status })
  } })
  const ingest = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: new Hono().route("/host", vertexProbeRoutes()).fetch })
  const requestFetch = globalThis.fetch
  const external: string[] = []
  globalThis.fetch = Object.assign((input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(input instanceof Request ? input.url : String(input))
    if (![engine.url.origin, ingest.url.origin].includes(url.origin)) {
      external.push(url.origin)
      throw new Error("External network is forbidden in the native Vertex fixture.")
    }
    return requestFetch(input, init)
  }, requestFetch) as typeof fetch
  const input = { project: "fixture-project", location: "us-central1", publisher: "google", model: "fixture-model" }
  const post = (body: unknown) => fetch(new URL("/host/probe-vertex", ingest.url), {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  })
  try {
    process.env.DOCHAUS_DATA_ROOT = directory
    process.env.DOCHAUS_RUNTIME_DIR = path.join(directory, "runtime")
    process.env.DOCHAUS_DIR = directory
    process.env.OPENCODE_URL = engine.url.origin
    process.env.OPENCODE_SERVER_PASSWORD = password
    const success = await post({ ...input, project: " fixture-project ", location: " us-central1 ", model: " fixture-model " })
    expect(success.status).toBe(200)
    expect(await success.json()).toEqual({ ok: true })
    expect(requests).toEqual([{ path: "/api/dochaus/provider-test", authorization: `Basic ${btoa(`opencode:${password}`)}`, body: { providerID: "google-vertex", modelID: "fixture-model" } }])
    delete process.env.OPENCODE_SERVER_PASSWORD
    expect((await post({ ...input, publisher: "anthropic", location: "global" })).status).toBe(200)
    expect(requests[1]).toEqual({ path: "/api/dochaus/provider-test", authorization: null, body: { providerID: "google-vertex-anthropic", modelID: "fixture-model" } })

    // The same valid request receives the native model/policy/auth/readiness
    // failures unchanged in meaning; none may be reported as a passing check.
    for (const failure of [
      { status: 400, body: { ok: false, kind: "model", restriction: "free-tier", error: "Fixture model access unavailable." } },
      { status: 403, body: { ok: false, kind: "model", error: "Fixture provider disabled by policy." } },
      { status: 401, body: { ok: false, kind: "auth", expired: true, error: "Fixture connection expired." } },
      { status: 503, body: { ok: false, kind: "unavailable", error: "Fixture native engine is not ready." } },
    ]) {
      status = failure.status
      payload = { ...failure.body, privateDiagnostic: "must not escape", token: "synthetic-private-token" }
      const response = await post(input)
      expect(response.status).toBe(failure.status)
      expect(await response.json()).toEqual(failure.body)
    }
    status = 403; payload = { ok: true }
    const contradictory = await post(input)
    expect(contradictory.status).toBe(403)
    expect(await contradictory.json()).toMatchObject({ ok: false, kind: "unavailable" })
    status = 200; payload = { ok: true, privateDiagnostic: "x".repeat(9000) }
    expect((await post(input)).status).toBe(503)
    payload = "malformed managed result"
    expect((await post(input)).status).toBe(503)

    const beforeInvalid = requests.length
    for (const body of [null, { ...input, publisher: "other" }, { ...input, publisher: { toString: "google" } }, { ...input, model: "model#variant" }, { ...input, model: "" }, { ...input, endpoint: "https://outside.invalid" }])
      expect((await post(body)).status).toBe(400)
    for (const body of [{ ...input, project: "unsaved-project" }, { ...input, location: "global" }])
      expect((await post(body)).status).toBe(409)
    expect(requests).toHaveLength(beforeInvalid)
    expect(readFileSync(config, "utf8")).toBe(saved)
    expect(external).toEqual([])
    engine.stop(true)
    expect((await post(input)).status).toBe(503)
  } finally {
    globalThis.fetch = requestFetch
    for (const key of environment) {
      if (previous[key] === undefined) delete process.env[key]
      else process.env[key] = previous[key]
    }
    ingest.stop(true)
    engine.stop(true)
    rmSync(directory, { recursive: true, force: true })
  }
})
