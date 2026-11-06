import { afterEach, expect, test } from "bun:test"
import { TOOL_SYNTAX_ISSUE } from "./checks"
import { readLeanEvents } from "./events"
import { runLean } from "./runner"
import {
  SPINE_ABSENCE_PERMISSIONS, absenceClaimsPrompt, absenceIssue, absenceVerifyPrompt, checkAbsenceClaims, parseAbsenceClaims, parseAbsenceVerdict, type AbsenceEntry,
} from "./spine-absence"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS } from "./stages"
import { readLeanJson, readLeanState, readOpenItems } from "./state"
import { AGENT, SOURCE, SOURCE_TEXT, cleanupMatters, happyPath, matter, scriptedModel, type Reply } from "./test-support"

// The spine step's check of its parent entries' absence claims, driven through
// the real coordinator with a scripted model: a classifier job per entry, a
// verifier job per claim, a contradicted claim sent back to the spine with the
// source's words, recorded with an open item on the final attempt, and every
// settled job saved by content. All text is synthetic.

afterEach(cleanupMatters)

const DIRECTORS = "The supplier's current directors are Alex Doe and Sam Roe."
const SOURCE_WITH_DIRECTORS = `${SOURCE_TEXT} ${DIRECTORS}`
const CLAIM = "the supplier's current directors"
const ONLY = { kind: "fact", detail: "The supply terms are the only source of the note.", evidence: [{ source_id: "request", quote: "Draft a short supply note" }], limitations: [] }
const SILENT = { kind: "open-item", detail: "The note names who signs for the supplier.", evidence: [], limitations: ["No supplied document identifies the supplier's current directors."] }
const CITED = { kind: "fact", detail: "The supplier's current directors are Alex Doe and Sam Roe.", evidence: [{ source_id: SOURCE, quote: DIRECTORS }], limitations: [] }
const FOUND = { source_id: SOURCE, quote: DIRECTORS }
const ISSUE = absenceIssue(1, CLAIM, FOUND)

const spine = (...entries: object[]) => JSON.stringify({ entries, limitations: [] })
/** The classifier finds the claim in the silent entry alone. */
const classifier: Reply = prompt => JSON.stringify({ claims: prompt.includes(SILENT.limitations[0]!) ? [CLAIM] : [] })
const contradicting: Reply = () => JSON.stringify({ verdict: "contradicted", ...FOUND })
const prompts = (model: ReturnType<typeof scriptedModel>, agent: string) => model.prompts.filter(prompt => prompt.agent === agent)
const spineChecks = (directory: string) => readLeanEvents(directory).filter(event => event.stage === "spine" && event.event === "checker" && event.step === undefined)
const absenceCounts = (directory: string) => readLeanEvents(directory).filter(event => event.event === "spine_absence").map(event => event.counts)

test("the classifier reads one entry's kind, detail and limitations; the verifier the sources, entry and claim, each in tagged blocks", () => {
  const classify = absenceClaimsPrompt(SILENT)
  expect(classify).toContain(`<entry>\nkind: open-item\ndetail: ${SILENT.detail}\nlimitations:\n- ${SILENT.limitations[0]}\n</entry>`)
  expect(classify).toContain("<question>")
  expect(classify).toContain("no lookup is needed")
  expect(absenceClaimsPrompt({ ...SILENT, limitations: [] })).toContain("limitations:\nNone.")
  const verify = absenceVerifyPrompt([SOURCE], SILENT, CLAIM)
  for (const tag of ["sources", "entry", "claim", "question"]) expect(verify).toContain(`<${tag}>`)
  expect(verify).toContain(`- ${SOURCE}\n`)
  expect(verify).toContain(`<claim>\nThe entry says the sources do not record: ${CLAIM}\n</claim>`)
  // The sources come first, so the prefix every claim shares caches across the jobs.
  expect(verify.indexOf("<sources>")).toBeLessThan(verify.indexOf("<entry>"))
})

test("the classifier's reply is a list of claims, each trimmed and kept once; any other reply is retried", () => {
  expect(parseAbsenceClaims('{"claims":[]}')).toEqual({ issues: [], value: [] })
  expect(parseAbsenceClaims(`Here it is:\n{"claims":[" ${CLAIM} ", "${CLAIM}", "a price"]}`)).toEqual({ issues: [], value: [CLAIM, "a price"] })
  for (const reply of ['{"claims":"none"}', '{"claims":[""]}', '{"claims":[3]}', '{"absent":[]}', "None."])
    expect(parseAbsenceClaims(reply).issues).toEqual([expect.stringContaining('Reply with the JSON object only: {"claims":')])
  expect(parseAbsenceClaims('<tool_call>{"name":"read"}</tool_call>').issues).toEqual([TOOL_SYNTAX_ISSUE])
})

test("a verdict stands, or is contradicted by words the named source holds; a fake quote fails quoteFound", () => {
  const directory = matter(SOURCE_WITH_DIRECTORS)
  const parse = (reply: object | string) => parseAbsenceVerdict(typeof reply === "string" ? reply : JSON.stringify(reply), directory, [SOURCE])
  expect(parse({ verdict: "stands" })).toEqual({ issues: [], value: { verdict: "stands" } })
  expect(parse({ verdict: "contradicted", ...FOUND })).toEqual({ issues: [], value: { verdict: "contradicted", ...FOUND } })
  // Invented words are refused with the quotation check's own feedback.
  const [fake] = parse({ verdict: "contradicted", source_id: SOURCE, quote: "The directors are Jo Bloggs and Kim Lee." }).issues
  expect(fake).toContain(`The reply quotes "The directors are Jo Bloggs and Kim Lee." as from "${SOURCE}", but those words are not in that document.`)
  // A near miss (here a change of case) is shown the source's own words.
  expect(parse({ verdict: "contradicted", source_id: SOURCE, quote: DIRECTORS.toLowerCase() }).issues[0]).toContain(`The document reads: "${DIRECTORS.slice(0, -1)}". Copy those exact characters`)
  expect(parse({ verdict: "contradicted", source_id: "board-minutes.docx", quote: DIRECTORS }).issues).toEqual([`Unknown source_id "board-minutes.docx". Use one of ${JSON.stringify([SOURCE])}.`])
  expect(parse({ verdict: "contradicted", source_id: SOURCE }).issues[0]).toContain("names the source_id and quotes its exact words")
  expect(parse({ verdict: "unclear" }).issues).toEqual([expect.stringContaining('{"verdict":"stands"}')])
  expect(parse("I could not tell.").issues[0]).toContain('{"verdict":"stands"}')
})

test("a contradicted claim sends the spine back with the source's words; the restated spine passes, and an unchanged entry is not asked again", async () => {
  const directory = matter(SOURCE_WITH_DIRECTORS)
  const model = scriptedModel({ ...happyPath,
    [AGENT.spine]: prompt => prompt.includes("did not pass") ? spine(ONLY, CITED) : spine(ONLY, SILENT),
    [AGENT.absence]: classifier, [AGENT["absence-verify"]]: contradicting })
  const state = await runLean(directory, { sessions: model.port })

  expect(state.status).toBe("completed")
  const asked = prompts(model, AGENT.spine)
  expect(asked).toHaveLength(2)
  expect(asked[1]!.text).toContain(`- ${ISSUE}`)
  expect(ISSUE).toBe(`entries[1] says the sources do not record "${CLAIM}"; "${SOURCE}" states "${DIRECTORS}". Cite it and restate the entry, or drop the claim.`)
  expect(spineChecks(directory).map(event => [event.attempt, event.issues])).toEqual([[0, [ISSUE]], [1, []]])
  expect((readLeanJson(directory, "spine.json") as { entries: object[] }).entries).toEqual([ONLY, CITED])
  // Two entries classified on the first reply; on the retry only the restated one.
  expect(prompts(model, AGENT.absence)).toHaveLength(3)
  expect(prompts(model, AGENT["absence-verify"])).toHaveLength(1)
  expect(absenceCounts(directory)).toEqual([
    { entries: 2, claims: 1, reused: 0, unchecked: 0, stands: 0, contradicted: 1 },
    { entries: 2, claims: 0, reused: 1, unchecked: 0, stands: 0, contradicted: 0 },
  ])
  // The classifier runs tool-less, the verifier with the read tools, each under the owner session.
  const sessions = (agent: string) => model.sessions.filter(session => session.agent === agent)
  for (const session of sessions(AGENT.absence)) expect(session).toMatchObject({ parentID: "ses_owner", permissions: SPINE_ABSENCE_PERMISSIONS })
  expect(sessions(AGENT["absence-verify"])).toEqual([expect.objectContaining({ parentID: "ses_owner", permissions: LEAN_CHILD_PERMISSIONS })])
  expect(LEAN_STAGE_AGENTS.absence).toBe("lean-spine-absence")
  expect(LEAN_STAGE_AGENTS["absence-verify"]).toBe("lean-spine-absence-verify")
  // The step lists every job's session beside its own: 3 classifiers, 1 verifier and 2 splits.
  expect(state.stages.spine?.session_ids).toHaveLength(6)
  expect(readOpenItems(directory).filter(item => item.stage === "spine")).toEqual([])
})

test("a claim that stands passes the spine at once", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: () => spine(ONLY, SILENT), [AGENT.absence]: classifier })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(prompts(model, AGENT.spine)).toHaveLength(1)
  expect(prompts(model, AGENT["absence-verify"])).toHaveLength(1)
  expect(spineChecks(directory).map(event => event.issues)).toEqual([[]])
  expect(absenceCounts(directory)).toEqual([{ entries: 2, claims: 1, reused: 0, unchecked: 0, stands: 1, contradicted: 0 }])
  expect((readLeanJson(directory, "spine.json") as { entries: object[] }).entries).toEqual([ONLY, SILENT])
})

test("a spine that keeps a contradicted claim is accepted on the final attempt with the issue recorded and an open item for the notes", async () => {
  const directory = matter(SOURCE_WITH_DIRECTORS)
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: () => spine(ONLY, SILENT), [AGENT.absence]: classifier, [AGENT["absence-verify"]]: contradicting })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

  expect(prompts(model, AGENT.spine)).toHaveLength(LEAN_MAX_RETRIES + 1)
  // Every attempt carries the issue: as retry feedback, then recorded on the final one.
  expect(spineChecks(directory).map(event => [event.attempt, event.issues])).toEqual([0, 1, 2, 3].map(attempt => [attempt, [ISSUE]]))
  expect(readLeanState(directory)?.stages.spine?.feedback).toEqual([])
  // The entry is checked once; every later attempt reuses its saved claim and verdict.
  expect(prompts(model, AGENT.absence)).toHaveLength(2)
  expect(prompts(model, AGENT["absence-verify"])).toHaveLength(1)
  expect((readLeanJson(directory, "spine.json") as { entries: object[] }).entries).toEqual([ONLY, SILENT])
  const items = readOpenItems(directory).filter(item => item.stage === "spine")
  expect(items).toEqual([expect.objectContaining({
    stage: "spine", kind: "omission", ref: { ...FOUND, repair_target: "memo" },
    detail: `Drafting spine: this open-item says the sources do not record ${CLAIM}, but "${SOURCE}" states: "${DIRECTORS}". The lawyer should confirm the drafts and notes rely on what the source states.`,
  })])
  // The drafting notes are shown it and name it, so the gate passes with it disclosed, never silent.
  expect(prompts(model, AGENT.finalize).some(prompt => prompt.text.includes(`- ${items[0]!.id}:`))).toBe(true)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass" })
})

test("saved claims and verdicts are reused for unchanged entries; a changed entry asks again", async () => {
  const directory = matter(SOURCE_WITH_DIRECTORS)
  const model = scriptedModel({ [AGENT.absence]: classifier, [AGENT["absence-verify"]]: contradicting })
  const state = readLeanState(directory)!
  const check = (entries: AbsenceEntry[], final = false) => checkAbsenceClaims(directory, state, { sessions: model.port }, { entries, limitations: [] }, { final }, () => {})

  const first = await check([ONLY, SILENT])
  expect(first).toEqual({ issues: [ISSUE], recorded: [], openItems: [] })
  expect(model.prompts).toHaveLength(3)
  // Unchanged entries: no job is asked again, and the verdict is the saved one.
  expect(await check([ONLY, SILENT])).toEqual(first)
  expect(model.prompts).toHaveLength(3)
  // Final: the same verdict is recorded, with its open item, instead of retried.
  const final = await check([ONLY, SILENT], true)
  expect(final.issues).toEqual([])
  expect(final.recorded).toEqual([ISSUE])
  expect(final.openItems.map(item => [item.stage, item.kind, item.ref.repair_target])).toEqual([["spine", "omission", "memo"]])
  expect(model.prompts).toHaveLength(3)
  // A changed entry is classified and verified again.
  await check([ONLY, { ...SILENT, detail: "The note names who signs for the supplier, by role." }])
  expect(model.prompts).toHaveLength(5)
  expect(absenceCounts(directory).map(counts => counts?.reused)).toEqual([0, 3, 3, 1])
})

test("a verifier that never quotes the source leaves its claim standing, unverified, and is not saved", async () => {
  const directory = matter(SOURCE_WITH_DIRECTORS)
  const fake: Reply = () => JSON.stringify({ verdict: "contradicted", source_id: SOURCE, quote: "The directors are Jo Bloggs and Kim Lee." })
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: () => spine(ONLY, SILENT), [AGENT.absence]: classifier, [AGENT["absence-verify"]]: fake })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const verify = prompts(model, AGENT["absence-verify"])
  expect(verify).toHaveLength(LEAN_MAX_RETRIES + 1)
  expect(new Set(verify.map(prompt => prompt.sessionID)).size).toBe(1)
  for (const retry of verify.slice(1)) expect(retry.text).toContain("but those words are not in that document")
  expect(prompts(model, AGENT.spine)).toHaveLength(1)
  expect(absenceCounts(directory)).toEqual([{ entries: 2, claims: 1, reused: 0, unchecked: 1, stands: 0, contradicted: 0 }])
  expect(readLeanJson(directory, "spine-absence.json")).toMatchObject({ verdicts: {} })
})
