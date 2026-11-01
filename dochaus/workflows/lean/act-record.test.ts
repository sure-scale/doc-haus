import { Database } from "bun:sqlite"
import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  ACT_RECORD_AGENT, ACT_RECORD_SAMPLES, LEAN_ACTS_SCHEMA, actRecordBlock, actRecordPack, parseActHow, parseActList, readActRecord, runActRecordStage, signerFlags,
  type ActHow, type ActRecord, type ActWho, type ListedAct, type RecordedAct,
} from "./act-record"
import { jsonBreakNote } from "./checks"
import { evidencePrompt } from "./evidence"
import { researchHandoff } from "./research"
import { stageInputSha } from "./runner"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { quoteFound } from "./quotes"
import { initLeanState, writeLeanText, writeQuestions, type LeanState, type Question } from "./state"

// The act record's jobs driven through runActRecordStage with a scripted model,
// against a synthetic matter invented for these tests: a boatyard licensing a
// rowing club's boat storage. Only the model is scripted; every check is real.

const SOURCE = "storage-proposal.docx"
const SOURCE_TEXT = "Fenwick Boatyard Limited will grant the Harbourlight Rowing Club a licence to store its boats in the north shed. " +
  "The licence takes effect only once the Club Committee approves the licence fee. " +
  "Under the Club rules, only a general meeting of the members may approve a fee the Club pays."
const REQUEST = "Prepare a storage licence under which Fenwick Boatyard Ltd grants the Harbourlight Rowing Club a licence to store its boats. " +
  "The Club Committee is to approve the licence fee before the licence is granted. Also prepare a short cover note for the Club."

const GRANT: ListedAct = { act: "grant the storage licence", source: "request", quote: "Fenwick Boatyard Ltd grants the Harbourlight Rowing Club a licence to store its boats" }
const APPROVE: ListedAct = { act: "approve the licence fee", source: "request", quote: "The Club Committee is to approve the licence fee before the licence is granted" }
const STORE: ListedAct = { act: "store the boats in the north shed", source: SOURCE, quote: "store its boats in the north shed" }
const NOTE = { deliverable: "a short cover note for the Club", source: "request", quote: "Also prepare a short cover note for the Club" }

const GRANT_WHO: ActWho = { authorised_by: ["Fenwick Boatyard Ltd"], performed_by: ["Fenwick Boatyard Ltd"], basis: [{ source: "request", quote: GRANT.quote }], reason: null }
const APPROVE_WHO: ActWho = { authorised_by: ["the Club Committee"], performed_by: ["the Club Committee"], basis: [{ source: "request", quote: APPROVE.quote }], reason: null }
const GRANT_HOW: ActHow = {
  instrument: "the storage licence, its grant clause", executed_by: ["Fenwick Boatyard Limited", "Harbourlight Rowing Club"],
  requires: [{ act: 1, need: "the Club Committee's approval of the licence fee", basis: [{ source: SOURCE, quote: "The licence takes effect only once the Club Committee approves the licence fee." }] }],
  basis: [{ source: "request", quote: "Prepare a storage licence" }], reason: null, conflict: null,
}
const APPROVE_HOW: ActHow = { instrument: null, executed_by: [], requires: [], basis: [], reason: "The record does not say how the Club Committee records its approval.", conflict: null }

type Job = "what" | "verify" | "deliverables" | "who" | "how"
type Reply = (job: Job, attempt: number, sample: number) => unknown

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

function matter(request = REQUEST): { directory: string; state: LeanState } {
  const directory = mkdtempSync(path.join(tmpdir(), "act-record-"))
  directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"), { recursive: true, mode: 0o700 })
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(SOURCE, "1", SOURCE_TEXT, 0, SOURCE_TEXT.length)
  db.close()
  const state = initLeanState(directory, { request, ownerSessionID: "ses_owner", model: { providerID: "test", modelID: "scripted" } })
  return { directory, state }
}

/** A scripted model that answers each session's job as the prompt's reply shape asks for it. */
function scripted(reply: Reply) {
  const sessions: Parameters<StageSessionPort["createSession"]>[0][] = []
  const prompts: { agent: string; sessionID: string; text: string }[] = []
  const jobs = new Map<string, { job: Job; sample: number }>()
  const port: StageSessionPort = {
    async createSession(input) {
      sessions.push(input)
      return `ses_child${sessions.length}`
    },
    async prompt(input) {
      prompts.push(input)
      if (!jobs.has(input.sessionID)) {
        const job: Job = input.text.includes('{"index":0,"how":{') ? "how" : input.text.includes('{"index":0,"who":{') ? "who"
          : input.text.includes('{"index":0,"verify":{') ? "verify"
          : input.text.includes('{"deliverables":[') ? "deliverables" : "what"
        jobs.set(input.sessionID, { job, sample: [...jobs.values()].filter(row => row.job === job).length })
      }
      const { job, sample } = jobs.get(input.sessionID)!
      const value = reply(job, prompts.filter(prompt => prompt.sessionID === input.sessionID).length - 1, sample)
      return typeof value === "string" ? value : JSON.stringify(value)
    },
  }
  const asked = (job: Job) => prompts.filter(prompt => jobs.get(prompt.sessionID)?.job === job)
  return { port, sessions, prompts, asked }
}

const indexed = <T>(field: "verify" | "who" | "how", values: T[]) => ({ acts: values.map((value, index) => ({ index, [field]: value })) })
const KEEP = { decision: "act", same_as: null, kind: null, reason: null }
/** Every job answered validly: both acts listed and kept, the note asked for, WHO and HOW settled. */
const valid: Reply = job => job === "what" ? { acts: [GRANT, APPROVE] } : job === "deliverables" ? { deliverables: [NOTE] }
  : job === "verify" ? indexed("verify", [KEEP, KEEP])
  : job === "who" ? indexed("who", [GRANT_WHO, APPROVE_WHO]) : indexed("how", [GRANT_HOW, APPROVE_HOW])
const listedOnly = (record: ActRecord | undefined) => record?.acts.map(({ act, source, quote }) => ({ act, source, quote }))

async function run(reply: Reply, setup: (directory: string) => void = () => {}, request = REQUEST) {
  const { directory, state } = matter(request)
  setup(directory)
  const model = scripted(reply)
  const result = await runActRecordStage(directory, state, { sessions: model.port }, () => {})
  return { directory, model, result, record: readActRecord(directory) }
}

describe("WHAT samples", () => {
  test("the union over samples that fail, list nothing or list overlapping acts keeps each act once", async () => {
    const { model, result, record } = await run((job, _attempt, sample) => job !== "what" ? valid(job, 0, 0)
      : [
          { acts: [{ act: "grant the licence", source: "request", quote: "words the lawyer never wrote" }] },
          { acts: [] },
          { acts: [GRANT] },
          // A reworded act on part of the same words, and the other act.
          { acts: [{ act: "license the boat storage", source: "request", quote: "grants the Harbourlight Rowing Club a licence" }, APPROVE] },
        ][sample])
    expect(result).toEqual({ status: "done", outputs: ["acts.json"], openItems: [] })
    expect(record?.status).toBe("listed")
    expect(listedOnly(record)).toEqual([GRANT, APPROVE])
    const perSample = model.asked("what").reduce((counts, prompt) => counts.set(prompt.sessionID, (counts.get(prompt.sessionID) ?? 0) + 1), new Map<string, number>())
    // The failing sample is retried to the limit; the empty one is asked once.
    expect([...perSample.values()]).toEqual([LEAN_MAX_RETRIES + 1, 1, 1, 1])
    expect(record?.issues).toEqual([])
  })

  test("samples that all list nothing record none, and WHO and HOW are not asked", async () => {
    const { model, record } = await run(job => job === "what" ? { acts: [] } : { deliverables: [] })
    expect(record).toEqual({ schema: LEAN_ACTS_SCHEMA, status: "none", acts: [], verify: null, deliverables: [], issues: [] })
    expect(model.asked("what")).toHaveLength(ACT_RECORD_SAMPLES)
    expect([model.asked("who"), model.asked("how")]).toEqual([[], []])
  })

  test("samples that never pass record unchecked with their issues, never none", async () => {
    const { record } = await run(job => job === "what" ? { acts: [{ act: "grant the licence", source: "request", quote: "words the lawyer never wrote" }] } : { deliverables: [NOTE] })
    expect(record?.status).toBe("unchecked")
    expect(record?.acts).toEqual([])
    expect(record?.issues.map(row => row.job)).toEqual(["what"])
    expect(record?.issues[0].issues.join(" ")).toContain("words the lawyer never wrote")
    expect(record?.deliverables).toEqual([NOTE])
  })
})

describe("VERIFY", () => {
  // The same grant, directed again by the proposal's words: the union keeps both, as their sources differ.
  const GRANT_AGAIN: ListedAct = { act: "grant the boat storage licence", source: SOURCE, quote: "Fenwick Boatyard Limited will grant the Harbourlight Rowing Club a licence" }
  const COVER: ListedAct = { act: "prepare a cover note for the Club", source: "request", quote: NOTE.quote }
  const candidates = [GRANT, COVER, APPROVE, GRANT_AGAIN]
  const NOT_ACT = { decision: "not_act", same_as: null, kind: "document", reason: `The request asks for a note to be written: "${NOTE.quote}" [source: request]. Nothing is authorised or carried by it.` }
  const decided = indexed("verify", [KEEP, NOT_ACT, KEEP, { ...KEEP, decision: "same", same_as: 0 }])
  const listing: Reply = (job, attempt) => job === "what" ? { acts: candidates } : job === "verify" ? decided : valid(job, attempt, 0)

  test("a candidate that is not an act is recorded as excluded with its kind and reason, a repeat is merged, and neither reaches WHO or HOW", async () => {
    const { model, record } = await run(listing)
    expect(record?.status).toBe("listed")
    expect(listedOnly(record)).toEqual([GRANT, APPROVE])
    expect(record?.verify).toEqual({ excluded: [{ ...COVER, kind: "document", reason: NOT_ACT.reason }], merged: [{ ...GRANT_AGAIN, into: 0 }] })
    expect(record?.acts.map(act => [act.who, act.how])).toEqual([[GRANT_WHO, GRANT_HOW], [APPROVE_WHO, APPROVE_HOW]])
    // VERIFY sees every candidate; WHO and HOW see only the kept acts.
    expect(model.asked("verify")).toHaveLength(1)
    expect(model.asked("verify")[0]!.text).toContain(JSON.stringify(COVER.act))
    for (const job of ["who", "how"] as const) {
      const shown = model.asked(job)[0]!.text
      expect([shown.includes(JSON.stringify(COVER.act)), shown.includes(JSON.stringify(GRANT_AGAIN.act)), shown.includes(JSON.stringify(APPROVE.act))]).toEqual([false, false, true])
    }
  })

  test("every candidate excluded records none, keeping the exclusions, and WHO and HOW are not asked", async () => {
    const { model, record } = await run((job, attempt) => job === "what" ? { acts: [COVER] } : job === "verify" ? indexed("verify", [NOT_ACT]) : valid(job, attempt, 0))
    expect(record?.status).toBe("none")
    expect(record?.acts).toEqual([])
    expect(record?.verify?.excluded).toEqual([{ ...COVER, kind: "document", reason: NOT_ACT.reason }])
    expect([model.asked("who"), model.asked("how")]).toEqual([[], []])
  })

  test("an unknown candidate index, a repeat of a candidate that is not kept or an unverifiable quote in a reason is retried naming it", async () => {
    const wrong = [
      { acts: [...decided.acts, { index: 7, verify: KEEP }] },
      indexed("verify", [KEEP, NOT_ACT, KEEP, { ...KEEP, decision: "same", same_as: 1 }]),
      indexed("verify", [KEEP, { ...NOT_ACT, reason: 'The request says "write the Club a friendly letter" [source: request].' }, KEEP, { ...KEEP, decision: "same", same_as: 0 }]),
    ]
    const { model, record } = await run((job, attempt) => job === "verify" ? wrong[attempt] ?? decided : listing(job, attempt, 0))
    const retries = model.asked("verify").slice(1).map(prompt => prompt.text)
    expect(retries).toHaveLength(3)
    expect(retries[0]).toContain('gives "index": 7, which is not the index of a listed act')
    expect(retries[1]).toContain('"verify.same_as" is 1')
    expect(retries[2]).toContain("write the Club a friendly letter")
    expect(record?.verify?.excluded.map(row => row.act)).toEqual([COVER.act])
  })

  test("a kind other than a document or the advisers' own work is retried, and the corrected kind is recorded", async () => {
    const { model, record } = await run((job, attempt) => job === "verify" && attempt === 0
      ? indexed("verify", [KEEP, { ...NOT_ACT, kind: "term" }, KEEP, { ...KEEP, decision: "same", same_as: 0 }])
      : listing(job, attempt, 0))
    expect(model.asked("verify")[1]!.text).toContain('"verify.kind" is "term". For "not_act", give one of "document", "firm_work".')
    expect(record?.verify?.excluded).toEqual([{ ...COVER, kind: "document", reason: NOT_ACT.reason }])
  })

  test("a verify that never passes keeps the whole union for WHO and HOW and records that it did not run", async () => {
    const { model, record } = await run((job, attempt) => job === "verify" ? { acts: [] } : job === "what" ? { acts: [GRANT, APPROVE] } : valid(job, attempt, 0))
    expect(model.asked("verify")).toHaveLength(LEAN_MAX_RETRIES + 1)
    expect(record?.status).toBe("listed")
    expect(listedOnly(record)).toEqual([GRANT, APPROVE])
    expect(record?.verify).toBeNull()
    expect(record?.issues.map(row => row.job)).toEqual(["verify"])
    expect(record?.acts.map(act => [act.who, act.how])).toEqual([[GRANT_WHO, GRANT_HOW], [APPROVE_WHO, APPROVE_HOW]])
  })
})

describe("WHO and HOW", () => {
  test("a WHO party, quote or index error is retried naming its field, and the corrected reply is recorded", async () => {
    const { model, record } = await run((job, attempt) => job === "who" && attempt === 0
      ? { acts: [{ index: 0, who: { ...GRANT_WHO, authorised_by: ["Westmere Holdings"], basis: [{ source: "request", quote: "Fenwick leases the shed" }] } }] }
      : valid(job, attempt, 0))
    const retry = model.asked("who")[1]?.text ?? ""
    expect(retry).toContain('"who.authorised_by" names the party "Westmere Holdings"')
    expect(retry).toContain('"who.basis[0].quote"')
    expect(retry).toContain("No entry for act 1.")
    expect(record?.acts.map(act => act.who)).toEqual([GRANT_WHO, APPROVE_WHO])
    expect(record?.issues).toEqual([])
  })

  test("a HOW requirement on the act itself, then a circle of requirements, is refused and retried", async () => {
    const requiring = (first: number | null, second: number | null) => indexed("how", [
      { ...GRANT_HOW, requires: first === null ? [] : [{ ...GRANT_HOW.requires[0], act: first }] },
      { ...APPROVE_HOW, requires: second === null ? [] : [{ act: second, need: "the licence granted", basis: [{ source: "request", quote: GRANT.quote }] }] },
    ])
    const { model, record } = await run((job, attempt) => job !== "how" ? valid(job, attempt, 0) : [requiring(0, null), requiring(1, 0)][attempt] ?? valid(job, attempt, 0))
    const retries = model.asked("how").slice(1).map(prompt => prompt.text)
    expect(retries[0]).toContain('"how.requires[0].act" is 0, the act itself')
    expect(retries[1]).toContain("act 0 requires act 1, act 1 requires act 0")
    expect(retries).toHaveLength(2)
    expect(record?.acts.map(act => act.how)).toEqual([GRANT_HOW, APPROVE_HOW])
  })

  test("a HOW reply that is not valid JSON is retried naming where it breaks, what was expected there and the words around it", async () => {
    // A requirement's basis list closed as an object: `"}}` where `"}]}` belongs.
    const sound = JSON.stringify(indexed("how", [GRANT_HOW, APPROVE_HOW]))
    const broken = sound.replace('"}]}', '"}}')
    const at = broken.indexOf('"}}') + 2
    const opened = broken.indexOf('"basis":[') + '"basis":'.length
    const { model, record } = await run((job, attempt) => job === "how" && attempt === 0 ? broken : valid(job, attempt, 0))
    const retry = model.asked("how")[1]!.text
    expect(retry).toContain(`The reply is not valid JSON: at character ${at}, "]" was expected to close the "basis" list opened at character ${opened}, but "}" was found.`)
    expect(retry).toContain(`${broken.slice(at - 40, at)}>>>${broken.slice(at, at + 20)}`)
    expect(retry).toContain("Fix that and resend the complete JSON.")
    // The format restatement still follows.
    expect(retry).toContain('Reply with the JSON object only: {"acts":[{"index":0,"how":{')
    expect(record?.acts.map(act => act.how)).toEqual([GRANT_HOW, APPROVE_HOW])
  })

  test("HOW is shown each act's WHO", async () => {
    const { model } = await run(valid)
    // Only the WHO reply names the committee in lower case, as a JSON string.
    expect(model.asked("how")[0]!.text).toContain('"the Club Committee"')
    expect(model.asked("who")[0]!.text).not.toContain('"the Club Committee"')
  })

  test("a job that never passes leaves its field null for every act and keeps its issues; the acts stand", async () => {
    const { record } = await run((job, attempt) => job === "how" ? indexed("how", [{ ...GRANT_HOW, requires: [{ ...GRANT_HOW.requires[0], act: 0 }] }, APPROVE_HOW])
      : job === "deliverables" ? { deliverables: [{ ...NOTE, source: "minutes.docx" }] } : valid(job, attempt, 0))
    expect(listedOnly(record)).toEqual([GRANT, APPROVE])
    expect(record?.acts.map(act => [act.who, act.how])).toEqual([[GRANT_WHO, null], [APPROVE_WHO, null]])
    expect(record?.deliverables).toBeNull()
    expect(record?.issues.map(row => row.job)).toEqual(["deliverables", "how"])
    expect(record?.issues.find(row => row.job === "how")?.issues.join(" ")).toContain('"how.requires[0].act"')
  })

  test("a conflict between the instructions and a source is recorded with both quotes once each is found in its own side's source", async () => {
    const conflict = {
      instruction: { source: "request", quote: APPROVE.quote },
      authority: { source: SOURCE, quote: "only a general meeting of the members may approve a fee the Club pays" },
      difference: "The instructions have the Club Committee approve the fee; the Club rules leave that to a general meeting of the members.",
    }
    const approving = (attempt: number) => [
      // The authority's words are not in the source it names.
      { ...conflict, authority: { ...conflict.authority, quote: "the members decide every fee" } },
      // The authority side quotes the instructions.
      { ...conflict, authority: { source: "request", quote: "Prepare a storage licence" } },
    ][attempt] ?? conflict
    const { directory, model, record } = await run((job, attempt) => job !== "how" ? valid(job, attempt, 0) : indexed("how", [GRANT_HOW, { ...APPROVE_HOW, conflict: approving(attempt) }]))
    const retries = model.asked("how").slice(1).map(prompt => prompt.text)
    expect(retries).toHaveLength(2)
    expect(retries[0]).toContain('"how.conflict.authority[0].quote"')
    expect(retries[1]).toContain('"how.conflict.authority[0].source" is "request"')
    expect(record?.acts.map(act => act.how?.conflict ?? null)).toEqual([null, conflict])
    // The spine is shown both sides.
    expect(actRecordBlock(directory)).toContain(
      `   Conflict, an open item for the lawyer and not resolved here: the instructions say ${JSON.stringify(APPROVE.quote)} (request); ${SOURCE} says "${conflict.authority.quote}". ${conflict.difference}`)
  })

  test("requirements between listed acts are refused in a longer circle and accepted in a chain", () => {
    const { directory } = matter()
    const acts = [GRANT, APPROVE, STORE]
    const basis = [{ source: SOURCE, quote: "The licence takes effect only once the Club Committee approves the licence fee." }]
    const reply = (edges: (number | null)[]) => JSON.stringify(indexed("how", edges.map(act => ({ ...APPROVE_HOW, requires: act === null ? [] : [{ act, need: "an earlier act", basis }] }))))
    expect(parseActHow(reply([1, 2, 0]), directory, acts).issues.join(" ")).toContain("act 0 requires act 1, act 1 requires act 2, act 2 requires act 0")
    expect(parseActHow(reply([1, 2, null]), directory, acts).issues).toEqual([])
  })
})

describe("where a reply that is not JSON breaks", () => {
  const what = (reply: string) => jsonBreakNote(reply)?.match(/^The reply is not valid JSON: (at character \d+, .*?)\. Around it/)?.[1]
  test("a cut-off reply, a comma before a closing bracket and a missing comma are each placed; a fence is counted in the offset; valid JSON gives none", () => {
    expect(what('{"acts":[{"index":0}')).toBe('at character 20, the reply ends before "]" closes the "acts" list opened at character 8')
    expect(what('{"acts":[{"index":0},]}')).toBe('at character 21, a comma stands before "]" with no value after it')
    expect(what('{"acts":[{"index":0} {"index":1}]}')).toBe('at character 21, "," was expected between two values, but "{" was found')
    expect(what('```json\n{"acts":[{"index":0}}\n```')).toBe('at character 28, "]" was expected to close the "acts" list opened at character 16, but "}" was found')
    expect(what('{"acts":[{"why":"one\ntwo"}]}')).toBe("at character 20, a line break or other control character was found inside a string; write it as an escape such as \\n")
    expect(jsonBreakNote('{"acts":[{"why":"a \\"quoted\\" } ] word"}]}')).toBeUndefined()
  })
})

describe("an adopted recommendation", () => {
  // The lawyer delegated how the licence is granted and answered the fee
  // question in their own words; the request itself names no act. The adopted
  // option's words direct an act, but they are the drafter's judgment, so they
  // are quoted under their own name, never as the lawyer's.
  const ADOPTED = "Fenwick Boatyard Limited grants the Club a licence by one licence deed signed by both parties."
  const OWN_WORDS = "The Club Committee approves the licence fee first."
  const answeredAt = "2026-10-21T09:00:00.000Z"
  const QUESTIONS: Question[] = [
    { id: "q-grant", stage: "intake", question: "How is the licence granted?", context: "The proposal does not say.", status: "answered", kind: "judgment",
      advice: { options: [{ id: "o1", text: ADOPTED }, { id: "o2", text: "Grant the licence by letter." }], recommendation: { option: "o1", rationale: "A deed binds both parties.", basis: [] } },
      answer: ADOPTED, answer_kind: "adopted", option_id: "o1", answer_class: "delegated", lawyer_words: "Decide for me.", answered_at: answeredAt },
    { id: "q-fee", stage: "intake", question: "Who approves the fee?", context: "The proposal names the Committee.", status: "answered", kind: "instruction", answer: OWN_WORDS, answered_at: answeredAt },
  ]
  const asked = (directory: string) => writeQuestions(directory, QUESTIONS)
  const NO_ACT = "Prepare the storage paperwork for Fenwick Boatyard Limited and the Harbourlight Rowing Club."
  const DIRECTED: ListedAct = { act: "grant the storage licence by deed", source: "adopted:q-grant", quote: "grants the Club a licence by one licence deed" }
  const OWN: ListedAct = { act: "approve the licence fee", source: "request", quote: "The Club Committee approves the licence fee first" }
  const listing = (acts: ListedAct[]) => JSON.stringify({ acts })

  test("its words direct an act under its own source name on the first reply, never as the request or the lawyer's answer", async () => {
    const who: ActWho[] = [
      { authorised_by: ["Fenwick Boatyard Limited"], performed_by: ["Fenwick Boatyard Limited"], basis: [{ source: "adopted:q-grant", quote: DIRECTED.quote }], reason: null },
      { authorised_by: ["the Club Committee"], performed_by: ["the Club Committee"], basis: [{ source: "request", quote: OWN.quote }], reason: null },
    ]
    const conflict = {
      instruction: { source: "adopted:q-grant", quote: "one licence deed signed by both parties" },
      authority: { source: SOURCE, quote: "The licence takes effect only once the Club Committee approves the licence fee." },
      difference: "The adopted option has the deed take effect on signing; the proposal makes it wait for the fee approval.",
    }
    const how: ActHow[] = [{ ...APPROVE_HOW, conflict }, APPROVE_HOW]
    const { directory, model, record } = await run(job => job === "what" ? { acts: [DIRECTED, OWN] } : job === "deliverables" ? { deliverables: [] }
      : job === "verify" ? indexed("verify", [KEEP, KEEP]) : job === "who" ? indexed("who", who) : indexed("how", how), asked, NO_ACT)
    expect(listedOnly(record)).toEqual([DIRECTED, OWN])
    expect(record?.acts.map(act => [act.who, act.how])).toEqual([[who[0], how[0]], [who[1], how[1]]])
    expect(record?.issues).toEqual([])
    // Every sample passed on its first reply.
    expect(model.asked("what")).toHaveLength(ACT_RECORD_SAMPLES)
    expect(actRecordPack(directory)).toContain('"adopted:q-grant"')

    // The same words under "request" or the question's answer are refused, and the note names the adopted source.
    for (const source of ["request", "answer:q-grant"]) {
      const refused = parseActList(listing([{ ...DIRECTED, source }]), directory)
      expect(refused.value).toBeUndefined()
      expect(refused.issues.join(" ")).toContain('"adopted:q-grant"')
    }
    // The adopted source holds only the adopted option's words: not the lawyer's answer, nor another question's.
    expect(quoteFound(directory, { source_id: "adopted:q-grant", quote: "Decide for me" })).toBe(false)
    expect(quoteFound(directory, { source_id: "adopted:q-fee", quote: OWN.quote })).toBe(false)
    // The lawyer's own answer stays quotable as theirs.
    expect(parseActList(listing([OWN]), directory)).toEqual({ issues: [], value: [OWN] })
    expect(quoteFound(directory, { source_id: "answer:q-fee", quote: OWN.quote })).toBe(true)
    // The authority side of a conflict is never the adopted option.
    const authority = parseActHow(JSON.stringify(indexed("how", [{ ...APPROVE_HOW, conflict: { ...conflict, instruction: { source: "request", quote: OWN.quote }, authority: conflict.instruction } }, APPROVE_HOW])),
      directory, [DIRECTED, OWN])
    expect(authority.issues.join(" ")).toContain('"how.conflict.authority[0].source" is "adopted:q-grant"')
  })
})

describe("signer flags", () => {
  const act = (who: ActWho | null, how: ActHow | null): RecordedAct => ({ ...GRANT, who, how })
  const who = (...authorised_by: string[]): ActWho => ({ ...GRANT_WHO, authorised_by })
  const how = (...executed_by: string[]): ActHow => ({ ...GRANT_HOW, executed_by })
  const flags = (...acts: RecordedAct[]) => signerFlags({ schema: LEAN_ACTS_SCHEMA, status: "listed", acts, verify: null, deliverables: [], issues: [] })

  test("an instrument its authorising parties sign is ok, whatever the spelling of their legal form", () => {
    expect(flags(
      act(who("Fenwick Boatyard Ltd"), how("Fenwick Boatyard Limited", "Harbourlight Rowing Club")),
      act(who("Fenwick Boatyard Ltd."), how("FENWICK BOATYARD LIMITED")),
    ).map(flag => flag.status)).toEqual(["ok", "ok"])
  })

  test("an instrument that leaves out an authorising party is flagged with that party", () => {
    expect(flags(
      act(who("Fenwick Boatyard Ltd"), how("Fenwick Boatyard LLC")),
      act(who("the Club Committee", "Fenwick Boatyard Ltd"), how("Fenwick Boatyard Limited")),
      act(who("the Club Committee"), how()),
    )).toEqual([
      { act: 0, status: "flag", missing: ["Fenwick Boatyard Ltd"] },
      { act: 1, status: "flag", missing: ["the Club Committee"] },
      { act: 2, status: "flag", missing: ["the Club Committee"] },
    ])
  })

  test("an act with no authorising party, no WHO or no HOW is unchecked", () => {
    expect(flags(act(who(), how("Fenwick Boatyard Limited")), act(null, how("Fenwick Boatyard Limited")), act(who("Fenwick Boatyard Ltd"), null)).map(flag => flag.status))
      .toEqual(["unchecked", "unchecked", "unchecked"])
  })
})

test("every act-record prompt starts with the same pack, runs as the same agent under the owner session and names the documents without their text", async () => {
  const { directory, model } = await run(valid)
  const pack = actRecordPack(directory)
  const first = [...new Map(model.prompts.map(prompt => [prompt.sessionID, prompt] as const)).keys()]
    .map(sessionID => model.prompts.find(prompt => prompt.sessionID === sessionID)!.text)
  expect(first).toHaveLength(ACT_RECORD_SAMPLES + 4)
  for (const text of first) expect(text.startsWith(`${pack}\n\n`)).toBe(true)
  // The tails differ by job only: one for every WHAT sample, one each for the others.
  expect(new Set(first.map(text => text.slice(pack.length))).size).toBe(5)
  expect(new Set(model.sessions.map(session => `${session.agent} ${session.parentID}`))).toEqual(new Set([`${ACT_RECORD_AGENT} ses_owner`]))
  expect(new Set(model.prompts.map(prompt => prompt.agent))).toEqual(new Set([ACT_RECORD_AGENT]))
  expect(pack).toContain(SOURCE)
  expect(first.join("\n")).not.toContain("in the north shed")
})

describe("the spine's act record", () => {
  const recordAs = (directory: string, saved: ActRecord) => writeLeanText(directory, "acts.json", `${JSON.stringify(saved, null, 2)}\n`)

  test("a listed record shows each act's WHO, HOW and signer flag, then the deliverables, with the shared evidence in the spine's prompt", async () => {
    const { directory } = await run(valid)
    const block = actRecordBlock(directory)
    expect(block).toBe([
      "Act record:",
      "The act record is a reading of the instructions to weigh, not an instruction or evidence: use it when deciding which documents carry which acts and who executes them; " +
      "where the record and the sources differ, the sources govern and the difference becomes an open item.",
      '1. "grant the storage licence"',
      `   Directed by request: ${JSON.stringify(GRANT.quote)}`,
      '   WHO: authorised by "Fenwick Boatyard Ltd"; performed by "Fenwick Boatyard Ltd".',
      '   HOW: instrument "the storage licence, its grant clause"; executed by "Fenwick Boatyard Limited", "Harbourlight Rowing Club"; requires act 2 ("the Club Committee\'s approval of the licence fee").',
      '2. "approve the licence fee"',
      `   Directed by request: ${JSON.stringify(APPROVE.quote)}`,
      '   WHO: authorised by "the Club Committee"; performed by "the Club Committee".',
      `   HOW: instrument not settled; executed by not settled. Unsettled: ${APPROVE_HOW.reason}`,
      '   Signer flag: "the Club Committee" must authorise the act but is not among its instrument\'s recorded signatories.',
      "Deliverables:",
      `- "a short cover note for the Club", asked for by request: ${JSON.stringify(NOTE.quote)}`,
    ].join("\n"))
    expect(evidencePrompt(directory, "spine")).toContain(`Verified research:\n${researchHandoff(directory)}\n${block}\nBuild a compact drafting spine`)
  })

  test("a record listing no act says so; an unchecked or missing record shows no block", async () => {
    const { directory } = await run(job => job === "what" ? { acts: [] } : { deliverables: [] })
    expect(actRecordBlock(directory)).toEndWith("\nThe act record found no act that the instructions direct.\nDeliverables: none listed.")
    expect(evidencePrompt(directory, "spine")).toContain("The act record found no act that the instructions direct.")
    recordAs(directory, { ...readActRecord(directory)!, status: "unchecked", deliverables: null })
    expect(actRecordBlock(directory)).toBe("")
    expect(evidencePrompt(directory, "spine")).not.toContain("Act record:")
    expect(actRecordBlock(matter().directory)).toBe("")
  })

  test("a changed act record makes the spine stale; the same record rewritten does not", async () => {
    const { directory, record: saved } = await run(valid)
    const spine = stageInputSha(directory, "spine")
    recordAs(directory, saved!)
    expect(stageInputSha(directory, "spine")).toBe(spine)
    recordAs(directory, { ...saved!, acts: saved!.acts.slice(0, 1) })
    expect(stageInputSha(directory, "spine")).not.toBe(spine)
  })
})
