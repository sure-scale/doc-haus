import { Hono } from "hono"
import { existsSync } from "node:fs"
import path from "node:path"
import { openDb } from "./db"

// User work, separate from submitted answers and from workflow execution.
// A revision prevents one browser from silently replacing another's edits.
export function answerDraftRoutes(resolveMatter: (id: string) => string) {
  const app = new Hono()
  app.use("/matters/:id/answer-drafts/:scope", async (c, next) => {
    if (!/^[a-z0-9][a-z0-9-]*$/.test(c.req.param("id") ?? "")) return c.json({ error: "Invalid matter id." }, 400)
    if (!/^(session|lean)\.[a-zA-Z0-9_-]+$/.test(c.req.param("scope") ?? ""))
      return c.json({ error: "Invalid answer draft scope." }, 400)
    const directory = resolveMatter(c.req.param("id")!)
    if (!existsSync(path.join(directory, "matter.json"))) return c.json({ error: "Matter not found." }, 404)
    await next()
  })
  app.get("/matters/:id/answer-drafts/:scope", (c) => {
    const db = openDb(resolveMatter(c.req.param("id")))
    try {
      return c.json(readDraft(db, c.req.param("scope")))
    } finally {
      db.close()
    }
  })
  app.put("/matters/:id/answer-drafts/:scope", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (
      !input ||
      !Number.isSafeInteger(input.revision) ||
      input.revision < 0 ||
      !validAnswers(input.answers, c.req.param("scope").startsWith("lean."))
    )
      return c.json({ error: "Invalid answer drafts." }, 400)
    const db = openDb(resolveMatter(c.req.param("id")))
    try {
      const result = db
        .transaction(() => {
          const current = readDraft(db, c.req.param("scope"))
          if (current.revision !== input.revision) return null
          const revision = current.revision + 1
          db.run(
            "INSERT INTO answer_drafts(scope, answers, revision) VALUES (?, ?, ?) ON CONFLICT(scope) DO UPDATE SET answers = excluded.answers, revision = excluded.revision",
            [c.req.param("scope"), JSON.stringify(input.answers), revision],
          )
          return { answers: input.answers, revision }
        })
        .immediate()
      if (!result)
        return c.json(
          { error: "These answers changed in another browser. Copy your entries, then reload to reconcile them." },
          409,
        )
      return c.json(result)
    } finally {
      db.close()
    }
  })
  return app
}

export function deleteSessionAnswerDrafts(directory: string, sessionID: string) {
  if (!existsSync(path.join(directory, ".dochaus", "legal.db"))) return
  const db = openDb(directory)
  try {
    db.run("DELETE FROM answer_drafts WHERE scope = ?", [`session.${sessionID}`])
  } finally {
    db.close()
  }
}

function readDraft(db: ReturnType<typeof openDb>, scope: string) {
  const row = db.query("SELECT answers, revision FROM answer_drafts WHERE scope = ?").get(scope) as {
    answers: string
    revision: number
  } | null
  return row ? { answers: JSON.parse(row.answers), revision: row.revision } : { answers: {}, revision: 0 }
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

function validAnswers(value: unknown, lean: boolean): boolean {
  if (!record(value) || JSON.stringify(value).length > 1_000_000) return false
  return Object.entries(value).every(
    ([id, answer]) =>
      id.length > 0 &&
      id.length <= 256 &&
      (lean
        ? typeof answer === "string"
        : record(answer) &&
          Object.values(answer).every(
            (field) =>
              typeof field === "string" ||
              typeof field === "number" ||
              typeof field === "boolean" ||
              (Array.isArray(field) && field.every((item) => typeof item === "string")),
          )),
  )
}
