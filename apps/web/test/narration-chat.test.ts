import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { MemoryRouter } from "react-router-dom"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import { isLeanDraftingStatus, type LeanDraftingStatus } from "../src/api/ingest/lean"
import type { Part } from "../src/api/session-view"
import { stepLabel } from "../src/components/chat/step-labels"
import { partActivity, stageMilestones, type ActivityLine } from "../src/components/narration/activity"
import DraftingNarration from "../src/components/narration/DraftingNarration"
import ResultCard from "../src/components/narration/ResultCard"
import Thinking from "../src/components/narration/Thinking"
import TimelineEntry from "../src/components/narration/TimelineEntry"
import { draftedIn, narrationTimeline, stepRows } from "../src/components/narration/timeline"

// The chat side of a drafting run: folded tool rows, clipped searches, stage
// rows with their roll-ups, the Thinking pill and the result, rendered from the
// recorded lean status fixtures.

function lean(value: unknown): LeanDraftingStatus {
  if (!isLeanDraftingStatus(value)) throw new Error("fixture is not a lean drafting status")
  return value
}

const running = lean(start.status)
const done = lean(completed)

function tool(id: string, name: string, state: Extract<Part, { type: "tool" }>["state"]): Part {
  return { type: "tool", id, sessionID: "ses_child2", messageID: "msg", tool: name, state }
}

function activity(id: string, name: string, input: Record<string, unknown>, output = "FOUND") {
  const line = partActivity(tool(id, name, { status: "completed", input, output, title: "" }), 1, "matter")
  if (!line) throw new Error(`no activity for ${id}`)
  return line
}

test("a read names its file from the native path input, in the chat and the steps panel", () => {
  expect(activity("p1", "read", { path: "/m/sources/Supply Terms.docx" })).toMatchObject({
    text: "Read Supply Terms.docx",
    group: "read",
    target: "Supply Terms.docx",
  })
  expect(activity("p2", "read", { path: "/m/.dochaus/runs/r1/plan.json" }).text).toBe("Read the drafting plan")
  const part = tool("p3", "read", {
    status: "completed",
    input: { path: "/m/sources/Supply Terms.docx" },
    output: "",
    title: "",
  })
  if (part.type !== "tool") throw new Error("expected a tool part")
  expect(stepLabel(part, "/m")).toContain("Supply Terms.docx")
})

test("a long search shows a clipped query and keeps the full one for the tooltip", () => {
  const query = "delivery obligations and the remedies available when goods arrive late or damaged in transit"
  const line = activity("s1", "search-document", { document: "a/terms.docx", query })
  expect(line.text.length).toBeLessThan(query.length + 30)
  expect(line.text).toEndWith('…"')
  expect(line.detail).toBe(`"${query}"`)
  expect(activity("s2", "search-document", { document: "a/terms.docx", query: "delivery" }).detail).toBeUndefined()
})

test("consecutive uses of one tool fold into a counted row with the latest target and flags", () => {
  const lines = [
    activity("r1", "read-document", { document: "a/one.docx" }),
    activity("r2", "read-document", { document: "a/two.docx" }),
    activity("q1", "verify-quote", { document: "a/one.docx" }, "NOT FOUND: x"),
    activity("q2", "verify-quote", { document: "a/two.docx" }),
    activity("q3", "verify-quote", { document: "a/two.docx" }, "NOT FOUND: y"),
  ]
  const rows = stepRows(lines)
  expect(rows.map((row) => [row.id, row.count, row.text, row.target])).toEqual([
    ["r1", 2, "Read source material · 2", "two.docx"],
    ["q1", 3, "Checked quotations · 3, 2 not found word for word", "two.docx"],
  ])
  expect(rows[1]).toMatchObject({ flag: true, flagged: 2 })

  const failed = partActivity(
    tool("e1", "read-document", { status: "error", input: { document: "a/one.docx" }, error: "x", title: "" }),
    1,
    "matter",
  )
  expect(failed?.group).toBeUndefined()
  expect(stepRows([lines[0]!, failed!, lines[1]!])).toHaveLength(3)

  const live = narrationTimeline(lines, [], true).at(-1)
  if (live?.kind !== "steps") throw new Error("expected a steps block")
  expect(live.shown).toHaveLength(2)
  const html = renderToStaticMarkup(createElement(TimelineEntry, { block: live }))
  expect(html).toContain('aria-expanded="false"')
  expect(html).toContain("Show 5 steps")
  expect(html).toContain('class="narration-line tool flagged"')
  expect(html).toContain('<span class="narration-suffix"> — two.docx</span>')
})

test("the Thinking pill alone is the live region and names what the run is doing, with its elapsed time", () => {
  const html = renderToStaticMarkup(
    createElement(Thinking, { label: "Reading Supply Terms.docx", since: Date.now() - 65_000, summarising: false }),
  )
  expect(html).toStartWith(
    '<div class="narration-thinking" role="status" aria-live="polite" title="Reading Supply Terms.docx">',
  )
  expect(html.match(/aria-live/g)).toHaveLength(1)
  expect(html).toContain('<span class="narration-thinking-text">Reading Supply Terms.docx</span>')
  expect(html).toMatch(/1m [56]s/)
  const summarising = renderToStaticMarkup(
    createElement(Thinking, { label: "Reading Supply Terms.docx", since: Date.now(), summarising: true }),
  )
  expect(summarising).toContain('<span class="narration-thinking-text">Summarising progress</span>')
})

test("a running run shows its live stage and current activity; a finished one shows its result", () => {
  const narration = {
    status: running,
    visible: true,
    lines: stageMilestones(running),
    notes: [] as { id: string; at: number; text: string; keys: string[] }[],
    narrating: false,
    failed: false,
    onForeignPart: () => {},
  }
  const render = (value: typeof narration) =>
    renderToStaticMarkup(
      createElement(
        MemoryRouter,
        null,
        createElement(DraftingNarration, { narration: value, matterId: "m1", showRequest: false, onViewDocument: () => {} }),
      ),
    )
  const live = render(narration)
  expect(live).not.toContain("narration-intro")
  expect(live).toContain('class="narration-stage is-running"')
  expect(live).toContain('<span class="narration-thinking-text">Drafting</span>')
  expect(live.match(/aria-live/g)).toHaveLength(1)

  const stage = narration.lines[0]!
  const reading: ActivityLine = {
    id: "p1",
    at: stage.at + 1,
    kind: "tool",
    state: "running",
    text: "Reading Supply Terms.docx",
    step: "d1.draft",
  }
  expect(render({ ...narration, lines: [...narration.lines, reading] })).toContain(
    '<span class="narration-thinking-text">Reading Supply Terms.docx</span>',
  )

  expect(draftedIn(running)).toBeUndefined()
  expect(draftedIn(done)).toMatch(/^(under a minute|\d+m|\d+h \d+m)$/)
  const finished = render({ ...narration, status: done, lines: stageMilestones(done) })
  expect(finished).toContain("Ready for review")
  expect(finished.match(/class="narration-stage is-done"/g)).toHaveLength(6)
  expect(finished).toContain('class="narration-duration"')
  expect(finished).not.toContain("narration-thinking")
  // The run header above the chat says how long the run took; the result does not repeat it.
  expect(finished).not.toContain("Drafted in")
})

test("a finished stage folds its activity into a one-line roll-up beside its duration", () => {
  const line = (id: string, extra: Partial<ActivityLine>): ActivityLine => ({
    id,
    at: 2,
    kind: "tool",
    state: "done",
    text: id,
    step: "d1.draft",
    ...extra,
  })
  const lines = [
    line("d1.draft:stage:a", { kind: "milestone", at: 1, until: 61_000, text: "Draft written" }),
    line("r1", { group: "read", target: "one.docx" }),
    line("r2", { group: "read", target: "two.docx" }),
    line("q1", { group: "quote", flag: true }),
    line("q2", { group: "quote" }),
    line("q3", { group: "quote", flag: true }),
  ]
  const block = narrationTimeline(lines, [], false)[0]
  if (block?.kind !== "stage") throw new Error("expected a stage row")
  const html = renderToStaticMarkup(createElement(TimelineEntry, { block }))
  expect(html).toContain('class="narration-icon is-done"')
  expect(html).toContain('<span class="narration-duration">1m 0s</span>')
  expect(html).toContain('class="narration-toggle is-rollup"')
  expect(html).toContain("2 reads of source material")
  expect(html).toContain("Checked 3 quotes, ")
  expect(html).toContain('<span class="narration-flagged">2 not found word for word</span>')
})


test("completed successful chronology is available behind a closed activity disclosure", () => {
  const html = renderFinishedNarration(stageMilestones(done))
  const disclosure = html.match(/<details><summary>Drafting activity[\s\S]*?<\/details>/)?.[0]
  expect(disclosure).toBeDefined()
  expect(disclosure).toContain("6 steps")
  expect(disclosure?.match(/class="narration-stage is-done"/g)).toHaveLength(6)
  expect(html.replace(disclosure!, "")).toContain("Review draft")
})

test("a flagged completed stage and narrator prose remain outside closed activity", () => {
  const stages = stageMilestones(done)
  const stage = stages[0]!
  const quote: ActivityLine = {
    ...activity("flagged-quote", "verify-quote", { document: "terms.docx" }, "NOT FOUND: quotation"),
    at: stage.at + 1,
    step: stage.step,
  }
  const html = renderFinishedNarration([...stages, quote], [{
    id: "note-review",
    at: stage.at + 2,
    text: "Check the quotation against the source before relying on it.",
    keys: [],
  }])
  const disclosures = html.match(/<details><summary>Drafting activity[\s\S]*?<\/details>/g) ?? []
  expect(disclosures.length).toBeGreaterThan(0)
  const outside = disclosures.reduce((rest, disclosure) => rest.replace(disclosure, ""), html)
  expect(outside).toContain('class="narration-flagged">1 not found word for word')
  expect(outside).toContain("Check the quotation against the source before relying on it.")
  expect(outside).toContain(stage.text)
  for (const disclosure of disclosures) {
    expect(disclosure).not.toContain("not found word for word")
    expect(disclosure).not.toContain("Check the quotation")
  }
})

function renderResult(status: LeanDraftingStatus) {
  return renderToStaticMarkup(createElement(ResultCard, { status, matterId: "m1", onViewDocument: () => {} }))
}

/** The completed fixture with more planned documents, each delivered with its
 * redline, as a run asked for several documents delivers them. */
function withDrafts(titles: string[]): LeanDraftingStatus {
  const rows = done.documents!.documents
  const body = rows.find((document) => document.kind === "body")!
  const redline = rows.find((document) => document.kind === "redline")!
  const added = titles.flatMap((title, index) => {
    const id = `d${index + 2}`
    return [
      { ...body, document: id, title, path: `.dochaus/lean/${id}/body-applied.docx`, deliverable: `${title}.docx` },
      { ...redline, document: id, title, path: `.dochaus/lean/${id}/redline.docx`, deliverable: `${title} - redline.docx` },
    ]
  })
  const notes = rows.findIndex((document) => document.kind === "drafting_notes")
  return {
    ...done,
    documents: { ...done.documents!, documents: [...rows.slice(0, notes), ...added, ...rows.slice(notes)] },
    plan: { documents: [...done.plan!.documents, ...titles.map((title, index) => ({ id: `d${index + 2}`, title, purpose: "" }))] },
  }
}

test("a lone draft keeps the page and its drafting notes are featured for review beside it", () => {
  const html = renderResult(done)
  const stack = html.slice(html.indexOf('<ul class="result-stack"'), html.indexOf('<details class="result-files"'))
  expect(html).toContain('<div class="result-document">')
  expect(html.match(/class="primary result-review"/g)).toHaveLength(1)
  expect(html).toContain("Review draft")
  expect(stack).toContain('aria-label="Also for review"')
  expect(stack).toContain('class="result-row is-notes"')
  expect(stack).toContain('aria-label="Review Supply note - drafting notes.docx"')
  // The draft stays the page and the redline stays with all the files.
  expect(stack).not.toContain("Supply note.docx")
  expect(stack).not.toContain("redline")
  expect(html).toContain('<details class="result-files">')
  expect(html).toContain('All generated files <span class="result-file-count">3</span>')
})

test("every draft of a several-document run is featured for review, in plan order, before the notes", () => {
  const html = renderResult(withDrafts(["Cover note", "Board resolution", "Disclosure letter"]))
  const stack = html.slice(html.indexOf('<ul class="result-stack"'), html.indexOf('<details class="result-files"'))
  expect(html).not.toContain('<div class="result-document">')
  expect(html).not.toContain("primary result-review")
  expect(stack).toContain('aria-label="Files for review"')
  const reviewed = [...stack.matchAll(/aria-label="Review ([^"]+)"/g)].map((match) => match[1])
  expect(reviewed).toEqual([
    "Supply note.docx",
    "Cover note.docx",
    "Board resolution.docx",
    "Disclosure letter.docx",
    "Supply note - drafting notes.docx",
  ])
  expect(stack.match(/class="result-row is-draft"/g)).toHaveLength(4)
  expect(stack).not.toContain("redline")
  // Everything to review is featured, so the full list stays folded.
  expect(html).toContain('<details class="result-files">')
  expect(html).toContain('All generated files <span class="result-file-count">9</span>')
  const files = html.slice(html.indexOf('<details class="result-files"'))
  expect(files).toContain("Disclosure letter - redline.docx")
})

function renderFinishedNarration(lines: ActivityLine[], notes: { id: string; at: number; text: string; keys: string[] }[] = []) {
  return renderToStaticMarkup(createElement(MemoryRouter, null, createElement(DraftingNarration, {
    narration: {
      status: done, visible: true, lines, notes, narrating: false, failed: false, onForeignPart: () => {},
    },
    matterId: "m1", showRequest: false, onViewDocument: () => {},
  })))
}
