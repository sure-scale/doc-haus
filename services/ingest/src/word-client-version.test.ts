import { expect, test } from "bun:test"
import { Hono } from "hono"
import { tmpdir } from "node:os"
import path from "node:path"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { wordCors } from "./word-documents"

test("stale Word clients cannot reach any token-route mutation, while current clients and preflight work", async () => {
  const keys = ["DOCHAUS_WORD_ORIGIN", "DOCHAUS_WORD_EXTERNAL_HTTPS", "DOCHAUS_RUNTIME_DIR"] as const
  const previous = keys.map((key) => process.env[key])
  const origin = "https://localhost:47612"
  process.env.DOCHAUS_WORD_ORIGIN = origin
  process.env.DOCHAUS_WORD_EXTERNAL_HTTPS = "1"
  process.env.DOCHAUS_RUNTIME_DIR = path.join(tmpdir(), `dochaus-word-version-${crypto.randomUUID()}`)
  let mutations = 0
  const app = new Hono()
  app.use("/word/*", wordCors)
  app.post("/word/:operation", (c) => { mutations++; return c.json({ ok: true }) })
  try {
    for (const operation of ["saves", "heartbeat", "done"]) {
      for (const version of [undefined, "0.0.0.0"]) {
        const response = await app.request(`/word/${operation}`, {
          method: "POST", headers: { Origin: origin, ...(version ? { "x-dochaus-word-version": version } : {}) },
        })
        expect(response.status).toBe(426)
        expect(response.headers.get("Access-Control-Allow-Origin")).toBe(origin)
        expect(await response.json()).toMatchObject({ code: "word_addin_update_required", error: expect.stringContaining("out of date") })
        expect(mutations).toBe(0)
      }
    }
    const preflight = await app.request("/word/saves", { method: "OPTIONS", headers: { Origin: origin } })
    expect(preflight.status).toBe(204)
    expect(preflight.headers.get("Access-Control-Allow-Headers")?.toLowerCase()).toContain("x-dochaus-word-version")
    expect(mutations).toBe(0)
    for (const operation of ["saves", "heartbeat", "done"]) {
      expect((await app.request(`/word/${operation}`, {
        method: "POST", headers: { Origin: origin, "x-dochaus-word-version": WORD_ADDIN_VERSION },
      })).status).toBe(200)
    }
    expect(mutations).toBe(3)
  } finally {
    keys.forEach((key, index) => {
      if (previous[index] === undefined) delete process.env[key]
      else process.env[key] = previous[index]
    })
  }
})
