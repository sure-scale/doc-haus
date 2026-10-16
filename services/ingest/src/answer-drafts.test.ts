import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { answerDraftRoutes, deleteSessionAnswerDrafts } from "./answer-drafts"
import { AnswerDraftStore, type DraftAnswers } from "../../../apps/web/src/answer-draft-store"

const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))

function fixture() {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-answer-drafts-"))
  roots.push(root)
  for (const id of ["first", "second"]) {
    mkdirSync(path.join(root, id))
    writeFileSync(path.join(root, id, "matter.json"), JSON.stringify({ id }))
  }
  return answerDraftRoutes((id) => path.join(root, id))
}

test("HTTP drafts survive database reopen and another browser; isolate matters and sessions", async () => {
  const app = fixture()
  const server = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: app.fetch })
  const url = `http://127.0.0.1:${server.port}/matters/first/answer-drafts/session.one`
  try {
    const put = await fetch(url, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ revision: 0, answers: { form1: { name: "Unfinished answer", choices: ["one"] } } }),
    })
    expect(put.status).toBe(200)
    expect(await (await fetch(url)).json()).toEqual({
      revision: 1,
      answers: { form1: { name: "Unfinished answer", choices: ["one"] } },
    })
    for (const endpoint of ["/matters/first/answer-drafts/session.two", "/matters/second/answer-drafts/session.one"]) {
      expect(await (await fetch(`http://127.0.0.1:${server.port}${endpoint}`)).json()).toEqual({
        answers: {},
        revision: 0,
      })
    }
  } finally {
    server.stop(true)
  }
})

test("stale browser cannot replace answers or recreate a cleared draft", async () => {
  const app = fixture()
  const url = "/matters/first/answer-drafts/lean.owner"
  const put = (revision: number, answers: DraftAnswers) =>
    app.request(url, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ revision, answers }),
    })
  expect((await put(0, { q1: "First browser" })).status).toBe(200)
  expect((await put(0, { q1: "Stale browser" })).status).toBe(409)
  expect(await (await app.request(url)).json()).toEqual({ revision: 1, answers: { q1: "First browser" } })
  expect((await put(1, {})).status).toBe(200)
  expect((await put(1, { q1: "Delayed save" })).status).toBe(409)
  expect(await (await app.request(url)).json()).toEqual({ revision: 2, answers: {} })
})

test("reject malformed input, missing matters, invalid ids and scopes", async () => {
  const app = fixture()
  expect((await app.request("/matters/missing/answer-drafts/lean.owner")).status).toBe(404)
  expect((await app.request("/matters/bad.id/answer-drafts/lean.owner")).status).toBe(400)
  expect((await app.request("/matters/first/answer-drafts/invalid")).status).toBe(400)
  for (const body of [
    "broken",
    JSON.stringify({ revision: -1, answers: {} }),
    JSON.stringify({ revision: 0, answers: { q: { text: "wrong kind" } } }),
  ]) {
    expect(
      (
        await app.request("/matters/first/answer-drafts/lean.owner", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body,
        })
      ).status,
    ).toBe(400)
  }
})

test("deleting a conversation removes its unfinished answers without clearing other sessions", async () => {
  const app = fixture()
  for (const scope of ["session.one", "session.two"]) {
    await app.request(`/matters/first/answer-drafts/${scope}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ revision: 0, answers: { form1: { text: "Unsent answer" } } }),
    })
  }
  deleteSessionAnswerDrafts(path.join(roots.at(-1)!, "first"), "one")
  expect(await (await app.request("/matters/first/answer-drafts/session.one")).json()).toEqual({
    answers: {},
    revision: 0,
  })
  expect((await (await app.request("/matters/first/answer-drafts/session.two")).json()).answers.form1.text).toBe(
    "Unsent answer",
  )
})

test("browser migration keeps local answers on failure and gives server entries precedence", async () => {
  let saved = { answers: { shared: "Server answer" } as DraftAnswers, revision: 1 }
  let fail = true
  let removed = false
  const store = new AnswerDraftStore(
    async (input) => {
      if (!input) return saved
      if (fail) throw new Error("Offline: keep this page open")
      saved = { answers: input.answers, revision: input.revision + 1 }
      return saved
    },
    () => ({ shared: "Old browser answer", local: "Unsent work" }),
    () => {
      removed = true
    },
  )
  await store.load()
  expect(removed).toBe(false)
  expect(store.getSnapshot().answers).toEqual({ shared: "Server answer", local: "Unsent work" })
  expect(store.getSnapshot().error).toContain("Offline")
  fail = false
  await store.retry()
  expect(removed).toBe(true)
  expect(saved.answers.local).toBe("Unsent work")
})

test("queued edits and submitted-answer removal finish in order after dialog unmount", async () => {
  const writes: DraftAnswers[] = []
  let release: (() => void) | undefined
  const store = new AnswerDraftStore(async (input) => {
    if (!input) return { answers: {}, revision: 0 }
    writes.push(input.answers)
    if (writes.length === 1)
      await new Promise<void>((resolve) => {
        release = resolve
      })
    return { answers: input.answers, revision: input.revision + 1 }
  })
  await store.load()
  const unsubscribe = store.subscribe(() => {})
  store.set("q1", "First keystroke")
  store.set("q1", "Finished answer")
  store.set("q1", undefined)
  unsubscribe()
  release!()
  for (let i = 0; i < 20 && store.getSnapshot().saving; i++) await new Promise((resolve) => setTimeout(resolve, 1))
  expect(writes).toEqual([{ q1: "First keystroke" }, {}])
  expect(store.getSnapshot().saving).toBe(false)
})
