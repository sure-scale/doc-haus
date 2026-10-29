import { Hono } from "hono"
import { getBunServer } from "hono/bun"
import { existsSync } from "node:fs"
import path from "node:path"
import type { ChronologyDecision } from "../../../dochaus/lib/chronology"
import {
  actOnChronology,
  batchChronologyProposals,
  ChronologyError,
  chronologyProposalDrafts,
  decideChronologyProposal,
  projectChronologySources,
  readChronology,
  stageChronologyProposals,
  validChronologyAction,
  verifyChronologyProposals,
} from "./chronology-store"
import { assessChronology, chronologyPreflight, startChronologyRun, stopChronologyRun } from "./chronology-generation"

const revision = (x: unknown): x is number =>
  Number.isSafeInteger(x) && (x as number) >= 0 && (x as number) < Number.MAX_SAFE_INTEGER
const record = (x: unknown): x is Record<string, unknown> => !!x && typeof x === "object" && !Array.isArray(x)

export function chronologyRoutes(resolveMatter: (id: string) => string) {
  const app = new Hono()
  app.onError((error, c) => {
    if (error instanceof ChronologyError) return c.json({ error: error.message, code: error.code }, error.status)
    throw error
  })
  app.use("/matters/:id/chronology/*", async (c, next) => {
    const id = c.req.param("id") ?? ""
    if (!/^[a-z0-9][a-z0-9-]*$/.test(id)) return c.json({ error: "Invalid matter id." }, 400)
    const dir = resolveMatter(id)
    if (!existsSync(path.join(dir, "matter.json"))) return c.json({ error: "Matter not found." }, 404)
    await next()
  })
  app.use("/matters/:id/chronology", async (c, next) => {
    const id = c.req.param("id") ?? ""
    if (!/^[a-z0-9][a-z0-9-]*$/.test(id)) return c.json({ error: "Invalid matter id." }, 400)
    const dir = resolveMatter(id)
    if (!existsSync(path.join(dir, "matter.json"))) return c.json({ error: "Matter not found." }, 404)
    await next()
  })
  app.get("/matters/:id/chronology", (c) => c.json(projectChronologySources(resolveMatter(c.req.param("id")))))
  app.get("/matters/:id/chronology/preflight", async (c) =>
    c.json(await chronologyPreflight(resolveMatter(c.req.param("id")))),
  )
  app.post("/matters/:id/chronology/assess", async (c) => {
    // A bounded model job can outlast Bun's HTTP idle timeout. Keep only
    // this assessment response open; ordinary routes retain their limits.
    const server = c.env ? getBunServer<Bun.Server<unknown>>(c) : undefined
    server?.timeout?.(c.req.raw, 0)
    const dir = resolveMatter(c.req.param("id"))
    return c.json(projectChronologySources(dir, await assessChronology(dir)))
  })
  app.post("/matters/:id/chronology/actions", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (!record(input) || !revision(input.expectedRevision) || !validChronologyAction(input.action))
      return c.json({ error: "Supply a current revision and valid chronology action." }, 400)
    const dir = resolveMatter(c.req.param("id"))
    const result = actOnChronology(dir, input.expectedRevision, input.action)
    return c.json(
      "preview" in result
        ? {
            preview: projectChronologySources(dir, { ...readChronology(dir), events: result.preview, proposals: [] })
              .events,
          }
        : projectChronologySources(dir, result),
    )
  })
  app.post("/matters/:id/chronology/proposals", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (!record(input) || !revision(input.expectedRevision))
      return c.json({ error: "Supply the current chronology revision." }, 400)
    const dir = resolveMatter(c.req.param("id"))
    const proposals = chronologyProposalDrafts(input.proposals, readChronology(dir))
    await verifyChronologyProposals(dir, proposals)
    return c.json(projectChronologySources(dir, stageChronologyProposals(dir, input.expectedRevision, proposals)))
  })
  app.post("/matters/:id/chronology/proposals/batch", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (
      !record(input) ||
      !revision(input.expectedRevision) ||
      !Array.isArray(input.proposalIds) ||
      !input.proposalIds.every((id) => typeof id === "string")
    )
      return c.json({ error: "Choose straightforward additions and supply the current revision." }, 400)
    const dir = resolveMatter(c.req.param("id"))
    return c.json(
      projectChronologySources(dir, batchChronologyProposals(dir, input.expectedRevision, input.proposalIds)),
    )
  })
  app.post("/matters/:id/chronology/proposals/:proposalId/decision", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (
      !record(input) ||
      !revision(input.expectedRevision) ||
      !["accept", "keep", "defer"].includes(String(input.decision))
    )
      return c.json({ error: "Choose Accept, Keep current or Skip for later and supply the current revision." }, 400)
    const dir = resolveMatter(c.req.param("id"))
    return c.json(
      projectChronologySources(
        dir,
        decideChronologyProposal(
          dir,
          input.expectedRevision,
          c.req.param("proposalId"),
          input.decision as ChronologyDecision,
        ),
      ),
    )
  })
  app.post("/matters/:id/chronology/run", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (
      !record(input) ||
      !revision(input.expectedRevision) ||
      !["all-cleared", "mixed", "stop"].includes(String(input.privilege)) ||
      !["initial", "refresh", "resume"].includes(String(input.mode)) ||
      (input.requestId !== undefined && typeof input.requestId !== "string")
    )
      return c.json({ error: "Choose a fresh privilege posture and run mode with the current revision." }, 400)
    const dir = resolveMatter(c.req.param("id"))
    return c.json(
      projectChronologySources(
        dir,
        await startChronologyRun(dir, {
          expectedRevision: input.expectedRevision,
          privilege: input.privilege as "all-cleared" | "mixed" | "stop",
          mode: input.mode as "initial" | "refresh" | "resume",
          requestId: input.requestId as string | undefined,
        }),
      ),
    )
  })
  app.post("/matters/:id/chronology/run/stop", async (c) => {
    const input = await c.req.json().catch(() => null)
    if (!record(input) || !revision(input.expectedRevision))
      return c.json({ error: "Supply the current revision to stop this run." }, 400)
    const dir = resolveMatter(c.req.param("id"))
    return c.json(projectChronologySources(dir, await stopChronologyRun(dir, input.expectedRevision)))
  })
  return app
}
