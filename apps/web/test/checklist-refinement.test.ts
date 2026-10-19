import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { MemoryRouter } from "react-router-dom"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import { refineLeanChecklist, type LeanDraftingStatus } from "../src/api/ingest/lean"
import RunHeader from "../src/components/drafting/RunHeader"
import type { RunActions } from "../src/components/drafting/useRunActions"

const originalFetch = globalThis.fetch

// Bun's fetch type also carries preconnect; keep it so a stub is a whole fetch.
const stubFetch = (respond: () => Promise<Response>): typeof fetch => Object.assign(respond, { preconnect: originalFetch.preconnect })
afterEach(() => { globalThis.fetch = originalFetch })

function refinementStatus(): LeanDraftingStatus {
  return {
    ...structuredClone(blockedGate) as LeanDraftingStatus,
    stages: { ...blockedGate.stages as LeanDraftingStatus["stages"], "checklist-check": { status: "blocked", attempts: 1, started_at: null, finished_at: null } },
    checklist_refinement: { available: true, parents: [{
      parent_id: "private-requirement-id", check: "Address each requested amendment in the draft.",
      target: { document: "d1", role: "body", treatment: "operative" },
      finding: { status: "indeterminate", reason: "The earlier inspection did not cover the whole draft." }, scopes: ["target_body"],
    }] },
  }
}

const actions: RunActions = { busy: false, pausing: false, stopping: false, error: "", resuming: false, resume: async () => true, answer: async () => true, update: async () => true, refine: async () => true, pause: async () => true, stop: async () => true }
function render(status: LeanDraftingStatus, own = true) {
  return renderToStaticMarkup(createElement(MemoryRouter, null, createElement(RunHeader, { status, actions, own, onOpen: () => {}, onShowDrafts: () => {} })))
}

test("full draft check requires explicit selection in the settled owner run", () => {
  const status = refinementStatus()
  const html = render(status)
  expect(html).toContain("Address each requested amendment in the draft.")
  expect(html).toContain("Could not confirm: The earlier inspection did not cover the whole draft.")
  expect(html).toContain('type="radio"')
  expect(html).not.toContain('checked=""')
  expect(html).toMatch(/<button class="primary" type="submit" disabled="">Check the full draft<\/button>/)
  expect(html).not.toContain("private-requirement-id")
  expect(html).not.toContain("target_body")
  for (const other of [
    { ...status, in_flight: true }, { ...status, pending_requests: 1 }, { ...status, status: "completed" as const },
    { ...status, checklist_refinement: { ...status.checklist_refinement!, available: false } },
    { ...status, stages: { ...status.stages, "checklist-check": { ...status.stages["checklist-check"]!, status: "done" as const } } },
  ]) expect(render(other)).not.toContain("Check the full draft")
  expect(render(status, false)).not.toContain("Check the full draft")
})

test("full draft check posts the selected original requirement with owner identity and preserves refusals", async () => {
  const calls: { url: string; init?: RequestInit }[] = []
  globalThis.fetch = (async (url, init) => {
    calls.push({ url: String(url), init })
    return Response.json({ profile: "lean", mode: "queued", request_id: "queued-refinement", status: { ...refinementStatus(), pending_requests: 1 } }, { status: 202 })
  }) as typeof fetch
  const result = await refineLeanChecklist("matter-check", "owner-conversation", "private-requirement-id")
  expect(result.status.pending_requests).toBe(1)
  expect(calls[0]!.url).toEndWith("/matters/matter-check/lean/checklist/refine")
  expect(calls[0]!.init?.method).toBe("POST")
  expect(JSON.parse(String(calls[0]!.init?.body))).toEqual({ ownerSessionID: "owner-conversation", parents: [{ parent_id: "private-requirement-id", scope: "target_body" }] })
  globalThis.fetch = stubFetch(async () => Response.json({ error: "The active drafting run changed. Reload before refining its checklist." }, { status: 409 }))
  await expect(refineLeanChecklist("matter-check", "old-owner", "private-requirement-id")).rejects.toThrow("The active drafting run changed. Reload before refining its checklist.")
})
