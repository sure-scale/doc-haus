import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { initLeanState, writeLeanState } from "../../../dochaus/workflows/lean/state"
import { deleteMatterSession, MatterSessionDeletionConflictError, MatterSessionDeletionError } from "./matter-session"

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach(cleanup => cleanup()))

function engine(status = 204) {
  const directory = mkdtempSync(path.join(tmpdir(), "matter-session-"))
  const deleted: string[] = []
  const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch(request) {
    const url = new URL(request.url)
    expect(request.headers.get("x-opencode-directory")).toBe(directory)
    // The engine adapter confirms the session's matter before a destructive call.
    if (request.method !== "DELETE") return Response.json({ data: { id: decodeURIComponent(url.pathname.split("/").at(-1)!), location: { directory } } })
    deleted.push(url.pathname)
    return new Response(null, { status })
  } })
  cleanups.push(() => { server.stop(true); rmSync(directory, { recursive: true, force: true }) })
  return { directory, deleted, options: { engineURL: server.url.origin } }
}

function leanRun(directory: string, status: "running" | "blocked" | "completed" | "failed") {
  const state = initLeanState(directory, { request: "Draft a note.", ownerSessionID: "ses_owner", model: { providerID: "provider", modelID: "model" } })
  writeLeanState(directory, { ...state, status })
}

test("the owner of an unfinished lean run cannot be deleted; other conversations can", async () => {
  for (const status of ["running", "blocked"] as const) {
    const f = engine()
    leanRun(f.directory, status)
    await expect(deleteMatterSession(f.directory, "ses_owner", f.options)).rejects.toBeInstanceOf(MatterSessionDeletionConflictError)
    expect(f.deleted).toEqual([])
    expect(await deleteMatterSession(f.directory, "ses_other", f.options)).toEqual({ ok: true, deleted: true })
    expect(f.deleted).toEqual(["/api/session/ses_other"])
  }
})

test("the owner of a finished lean run, or of no run, is deleted through the engine", async () => {
  for (const status of ["completed", "failed", undefined] as const) {
    const f = engine()
    if (status) leanRun(f.directory, status)
    expect(await deleteMatterSession(f.directory, "ses_owner", f.options)).toEqual({ ok: true, deleted: true })
    expect(f.deleted).toEqual(["/api/session/ses_owner"])
  }
})

test("an absent session is not an error; an engine refusal is", async () => {
  const absent = engine(404)
  expect(await deleteMatterSession(absent.directory, "ses_gone", absent.options)).toEqual({ ok: true, deleted: false })
  const refused = engine(500)
  await expect(deleteMatterSession(refused.directory, "ses_owner", refused.options)).rejects.toBeInstanceOf(MatterSessionDeletionError)
})
