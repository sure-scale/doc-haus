import { afterEach, describe, expect, test } from "bun:test"
import { TOOL_SYNTAX_ISSUE, checkDraftingNotes } from "./checks"
import {
  LEAN_COVERAGE_SCHEMA, coverageCarries, coverageItems, coveragePlanHash, coveragePrompt, coverageScopeQuestion, parseCoverage, planCoverageIssues, type CoverageAct, type LeanCoverage,
} from "./package-coverage"
import { writeLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { EXCERPT, instructionFound } from "./quotes"
import { disclosableItems } from "./stages"
import { writeLeanText, writeQuestions } from "./state"
import { SOURCE, cleanupMatters, matter } from "./test-support"

// Package coverage: the coverage job's reply is verified against the matter,
// and the plan is held to the verified acts. Synthetic parties and acts only.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule. Kestrel Works then releases the deposit."
const APPROVE = "the Orchard Committee approves the revised schedule"
const RELEASE = "Kestrel Works shall release the deposit"

const SHAPE = 'Reply with the JSON object only: {"acts":[{"act":"...","source":"...","quote":"...","parties":["..."],"document":"planned id or null"}]}.'
const rowIssue = (index: number) =>
  `acts[${index}] needs nonempty "act", "source" and "quote" strings, "parties" as a list of names ([] when the record does not settle them), and "document" as a planned id or null. ${SHAPE}`
const DIRECTED = `"approve the revised schedule" ("${APPROVE}" [source: request])`
const move = (verb: string, who = '"the Orchard Committee"') => `${verb} the act to a document executed by ${who}, or add a companion document executed by ${who}.`
const NOT_EXECUTED_BY_KESTREL = `${DIRECTED} must be taken by "the Orchard Committee", but d1 "Release letter" is executed by "Kestrel Works". ${move("Move")}`
const unnamedParty = (party: string) =>
  `acts[0] ("approve the revised schedule") names the party "${party}", but neither the request, the lawyer's updates and answers nor any matter document names that party. ` +
  "Name each party as the sources name it, and list only parties the record names ([] when it does not settle who must take the act)."

const planned = (id: DocID, title: string, executed_by?: string[]): PlannedDocument =>
  ({ id, title, purpose: "", parts: [{ heading: "Terms", scope: "" }], ...(executed_by ? { executed_by } : {}) })
const act = (fields: Partial<CoverageAct> = {}): CoverageAct =>
  ({ act: "approve the revised schedule", source: "request", quote: APPROVE, parties: ["the Orchard Committee"], document: "d1", ...fields })
const reply = (...acts: Partial<CoverageAct>[]) => JSON.stringify({ acts: acts.map(fields => act(fields)) })
/** The id the plan step gives its package-scope question on `acts` alone, through
 * the step itself: keyed on each act's quote and parties, never its document. */
const scopeID = (...acts: CoverageAct[]) => coverageScopeQuestion(acts.map(row => ({ ...row, document: null })), [])!.id

function prepared() {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", REQUEST)
  return directory
}

describe("plan coverage issues", () => {
  test("an act no planned document carries is planned in a document its party executes", () => {
    const directory = prepared()
    expect(planCoverageIssues([act({ document: null })], [planned("d1", "Release letter", ["Kestrel Works"])], directory)).toEqual([
      `The instructions direct ${DIRECTED}, but no planned document carries it. Plan it in a document executed by "the Orchard Committee", adding one when none is planned.`,
    ])
    // Without a settled party the act is still planned, but no party is supplied.
    expect(planCoverageIssues([act({ document: null, parties: [] })], [planned("d1", "Release letter")], directory)).toEqual([
      `The instructions direct ${DIRECTED}, but no planned document carries it. Plan it in the document that should carry it, adding a document when none is planned.`,
    ])
  })

  test("an act in a document its party does not execute names the act, its source, the party and the document", () => {
    const directory = prepared()
    expect(planCoverageIssues([act()], [planned("d1", "Release letter", ["Kestrel Works"])], directory)).toEqual([NOT_EXECUTED_BY_KESTREL])
    expect(planCoverageIssues([act()], [planned("d1", "Advice note", [])], directory)).toEqual([
      `${DIRECTED} must be taken by "the Orchard Committee", but d1 "Advice note" is executed by no party (Executed by: none). ${move("Move")}`,
    ])
    expect(planCoverageIssues([act()], [planned("d1", "Release letter")], directory)).toEqual([
      `${DIRECTED} must be taken by "the Orchard Committee", and d1 "Release letter" carries it but has no "Executed by:" line. ` +
      `Add that line naming the parties who execute it, separated by ";". If they do not include "the Orchard Committee", ${move("move")}`,
    ])
  })

  test("a party run together with another in one executing party is told to list each party separately", () => {
    const directory = prepared()
    writeLeanText(directory, "request.md", `${REQUEST} Kestrel Works and the Orchard Committee sign the release letter.`)
    expect(planCoverageIssues([act()], [planned("d1", "Release letter", ["Kestrel Works and the Orchard Committee"])], directory)).toEqual([
      `${DIRECTED} must be taken by "the Orchard Committee", but d1 "Release letter" is executed by "Kestrel Works and the Orchard Committee". ` +
      'It declares "Kestrel Works and the Orchard Committee" as one executing party: if that names several parties who execute the document, list each separately on the "Executed by:" line, ' +
      `separated by ";", as the request or sources name them; otherwise ${move("move")}`,
    ])
  })

  test("a party matches its document's executing party regardless of case, spacing, curly quotes or a leading the", () => {
    const directory = matter("The Orchard Committee approves the revised schedule and Kestrel's Works releases the deposit.")
    writeLeanText(directory, "request.md", REQUEST)
    const documents = [planned("d1", "Resolution", ["The  ORCHARD committee"]), planned("d2", "Release letter", ["Kestrel’s Works"])]
    expect(planCoverageIssues([act({ parties: ["orchard Committee"] }), act({ parties: ["kestrel's works"], document: "d2" })], documents, directory)).toEqual([])
  })

  test("a party matches its document's executing party whichever common spelling of its legal form each uses, but not another form", () => {
    const directory = matter("Kestrel Works Limited releases the deposit once Wren Atelier, L.L.C. approves the revised schedule.")
    writeLeanText(directory, "request.md", REQUEST)
    const documents = [planned("d1", "Resolution", ["Wren Atelier LLC"]), planned("d2", "Release letter", ["Kestrel Works Ltd"])]
    expect(planCoverageIssues([act({ parties: ["Wren Atelier, L.L.C."] }), act({ parties: ["Kestrel Works Limited"], document: "d2" })], documents, directory)).toEqual([])
    const who = '"Kestrel Works LLC"'
    expect(planCoverageIssues([act({ parties: ["Kestrel Works LLC"], document: "d2" })], documents, directory)).toEqual([
      `${DIRECTED} must be taken by ${who}, but d2 "Release letter" is executed by "Kestrel Works Ltd". ${move("Move", who)}`,
    ])
  })

  test("an act whose parties the record does not settle raises nothing once a document not declared as executed by none carries it, and no acts raise nothing", () => {
    const directory = prepared()
    expect(planCoverageIssues([act({ parties: [] })], [planned("d1", "Release letter", ["Kestrel Works"])], directory)).toEqual([])
    // No "Executed by:" line: no party is supplied for the plan to declare.
    expect(planCoverageIssues([act({ parties: [] })], [planned("d1", "Release letter")], directory)).toEqual([])
    expect(planCoverageIssues([], [planned("d1", "Release letter", ["Kestrel Works"]), planned("d2", "Advice note", [])], directory)).toEqual([])
  })

  test("an act carried by a document executed by no party is planned in a document a party executes", () => {
    expect(planCoverageIssues([act({ parties: [] })], [planned("d1", "Advice note", [])], prepared())).toEqual([
      `${DIRECTED} is carried by d1 "Advice note", which is executed by no party (Executed by: none). ` +
      'Plan the act in a document a party executes, naming that party as the request or sources name it; only when the record does not name who executes the document, ' +
      'leave its "Executed by:" line out instead of writing "none". Never supply a party the record does not name.',
    ])
  })

  test("an act several parties must take is held to a document all of them execute", () => {
    const directory = prepared()
    const both = act({ parties: ["the Orchard Committee", "Kestrel Works"] })
    const who = '"the Orchard Committee" and "Kestrel Works"'
    expect(planCoverageIssues([both], [planned("d1", "Release letter", ["Kestrel Works", "Orchard Committee"])], directory)).toEqual([])
    expect(planCoverageIssues([both], [planned("d1", "Release letter", ["Kestrel Works"])], directory)).toEqual([
      `${DIRECTED} must be taken by ${who}, but d1 "Release letter" is executed by "Kestrel Works", not by "the Orchard Committee". ${move("Move", who)}`,
    ])
    expect(planCoverageIssues([both], [planned("d1", "Advice note", [])], directory)).toEqual([
      `${DIRECTED} must be taken by ${who}, but d1 "Advice note" is executed by no party (Executed by: none). ${move("Move", who)}`,
    ])
    expect(planCoverageIssues([act({ parties: ["the Orchard Committee", "Kestrel Works"], document: null })], [], directory)).toEqual([
      `The instructions direct ${DIRECTED}, but no planned document carries it. Plan it in a document executed by ${who}, adding one when none is planned.`,
    ])
  })

  test("a declared executing party must be named by the instructions, an answer or a matter document", () => {
    const directory = prepared()
    writeQuestions(directory, [{ id: "q-1", stage: "intake", question: "Who countersigns?", context: "Needed to plan.", status: "answered", answer: "Wren Atelier countersigns." }])
    const documents = [planned("d1", "Resolution", ["orchard committee", "Wren Atelier"]), planned("d2", "Release letter", ["Kestrel Works", "Lantern Group"])]
    expect(planCoverageIssues([], documents, directory)).toEqual([
      'd2 "Release letter" declares "Executed by: Lantern Group", but neither the request, the lawyer\'s updates and answers nor any matter document names "Lantern Group". ' +
      'Name each executing party as the request or sources name it, separating several parties with ";". ' +
      'If the record does not name who executes the document, leave the "Executed by:" line out; never supply a party the record does not name.',
    ])
  })
})

describe("package-scope question", () => {
  test("acts left uncovered are asked about in one question, in the plural, each with where the plan has it", () => {
    const release = act({ act: "release the deposit", source: SOURCE, quote: RELEASE, parties: ["Kestrel Works"], document: "d2" })
    expect(coverageScopeQuestion([act(), release], [planned("d1", "Advice note", []), planned("d2", "Release letter")])?.question).toBe(
      "The document plan leaves these instructed acts without a document executed by the parties who must take them: " +
      `"approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; planned in d1 "Advice note", executed by no party. ` +
      `"release the deposit", directed by "${RELEASE}" [source: ${SOURCE}]; to be taken by "Kestrel Works"; planned in d2 "Release letter", executed by parties the plan does not name. ` +
      "For each, should the package add a document, or change a planned one, so that those parties take it, or is it outside this request? " +
      "If the package should carry it, say in which document and who executes it.")
    const one = coverageScopeQuestion([act()], [planned("d1", "Advice note", [])])!.question
    expect(one).toStartWith("The document plan leaves this instructed act without a document executed by the parties who must take it: ")
    expect(one).toContain("executed by no party. Should the package add a document")
  })
})

describe("coverage reply", () => {
  const documents = [planned("d1", "Resolution", ["the Orchard Committee"]), planned("d2", "Release letter", ["Kestrel Works"])]

  test("verified acts are returned as given, from a bare, json-fenced or plain-fenced reply", () => {
    const directory = prepared()
    const parsed = parseCoverage("```json\n" + reply({}, { act: "release the deposit", source: SOURCE, quote: RELEASE, parties: ["Kestrel Works"], document: "d2" }, { parties: [], document: null }) + "\n```", directory, documents)
    expect(parsed).toEqual({ issues: [], value: [act(), act({ act: "release the deposit", source: SOURCE, quote: RELEASE, parties: ["Kestrel Works"], document: "d2" }), act({ parties: [], document: null })] })
    expect(parseCoverage("```\n" + reply({}) + "\n```", directory, documents)).toEqual({ issues: [], value: [act()] })
    expect(parseCoverage('{"acts":[]}', directory, documents)).toEqual({ issues: [], value: [] })
  })

  test("a reply that prints a tool call is retried with the tool-call issue", () => {
    expect(parseCoverage('<tool_call>{"name":"read-document","arguments":{"document":"x"}}</tool_call>', prepared(), documents)).toEqual({ issues: [TOOL_SYNTAX_ISSUE] })
  })

  test("a reply that is not an object with an acts list is retried with the shape", () => {
    const directory = prepared()
    for (const text of ["The acts are listed below.", '{"acts":"x"}', "null", "[]"])
      expect(parseCoverage(text, directory, documents)).toEqual({ issues: [SHAPE] })
  })

  test("each malformed act is retried with the shape", () => {
    const directory = prepared()
    const { parties: _parties, ...unstated } = act()
    const rows: unknown[] = [
      "x", { ...act(), source: "" }, { ...act(), quote: "" }, { ...act(), parties: "Kestrel Works" }, { ...act(), parties: [5] }, { ...act(), parties: [" "] },
      { ...act(), parties: null }, { ...unstated, party: "Kestrel Works" }, { ...act(), document: 3 }, unstated, { ...act(), act: " " },
    ]
    for (const row of rows) expect(parseCoverage(JSON.stringify({ acts: [row] }), directory, documents)).toEqual({ issues: [rowIssue(0)] })
    // Every malformed act is reported, each by its own index.
    expect(parseCoverage(JSON.stringify({ acts: rows }), directory, documents)).toEqual({ issues: rows.map((_, index) => rowIssue(index)) })
  })

  test("an unknown document id is retried with the planned ids", () => {
    const parsed = parseCoverage(reply({ document: "d3" as DocID }), prepared(), documents)
    expect(parsed.value).toBeUndefined()
    expect(parsed.issues).toEqual([`acts[0] ("approve the revised schedule") names the document "d3", which is not planned. Use one of the planned ids d1, d2, or null when no planned document carries it.`])
  })

  test("a party named nowhere in the instructions or sources is retried, even when the plan declares it; one named there is not", () => {
    const directory = prepared()
    expect(parseCoverage(reply({ parties: ["Lantern Group"] }), directory, documents)).toEqual({ issues: [unnamedParty("Lantern Group")] })
    // Named in a matter document only, with different case and without "the".
    expect(parseCoverage(reply({ parties: ["ORCHARD COMMITTEE"] }), directory, documents)).toEqual({ issues: [], value: [act({ parties: ["ORCHARD COMMITTEE"] })] })
    // A planned document's executing party is the plan's own writing and grounds nothing.
    expect(parseCoverage(reply({ parties: ["Lantern Group"] }), directory, [...documents, planned("d3", "Consent", ["Lantern Group"])])).toEqual({ issues: [unnamedParty("Lantern Group")] })
    // Of several parties, each unnamed one is retried on its own.
    expect(parseCoverage(reply({ parties: ["Kestrel Works", "Lantern Group", "Wren Atelier"] }), directory, documents)).toEqual({ issues: [unnamedParty("Lantern Group"), unnamedParty("Wren Atelier")] })
  })

  test("an act listed again, differing only in spacing, case, a leading the or party order, is kept once and raises one issue", () => {
    const directory = prepared()
    const parsed = parseCoverage(reply({}, { act: " Approve  the revised SCHEDULE ", parties: ["Orchard committee"] }), directory, documents)
    expect(parsed).toEqual({ issues: [], value: [act()] })
    expect(planCoverageIssues(parsed.value!, [planned("d1", "Release letter", ["Kestrel Works"])], directory)).toEqual([NOT_EXECUTED_BY_KESTREL])
    const both = parseCoverage(reply({ parties: ["the Orchard Committee", "Kestrel Works"] }, { parties: ["kestrel works", "Orchard Committee"] }), directory, documents)
    expect(both).toEqual({ issues: [], value: [act({ parties: ["the Orchard Committee", "Kestrel Works"] })] })
    // A party listed twice for one act is kept once.
    expect(parseCoverage(reply({ parties: ["the Orchard Committee", "Orchard committee"] }), directory, documents)).toEqual({ issues: [], value: [act()] })
    // The same act carried by another document, or by none, or taken by other parties, is a separate act.
    expect(parseCoverage(reply({}, { document: null }, { parties: [] }), directory, documents)).toEqual({ issues: [], value: [act(), act({ document: null }), act({ parties: [] })] })
  })

  test("a quote not in its source, or a source that is not a matter document, is retried", () => {
    const directory = prepared()
    const missing = parseCoverage(reply({ source: SOURCE, quote: "the Orchard Committee shall approve every schedule" }), directory, documents)
    expect(missing.value).toBeUndefined()
    expect(missing.issues).toEqual([expect.stringContaining(`quotes "the Orchard Committee shall approve every schedule" as from "${SOURCE}", but those words are not in that document.`)])
    const misattributed = parseCoverage(reply({ quote: RELEASE }), directory, documents)
    expect(misattributed.issues).toEqual([expect.stringContaining(`They do appear in "${SOURCE}"`)])
    expect(parseCoverage(reply({ source: "minutes.docx" }), directory, documents)).toEqual({
      issues: [`acts[0] ("approve the revised schedule") gives the source "minutes.docx", which is not a matter document. Use one of "request", "${SOURCE}".`],
    })
  })
})

describe("an act whose document the instructions leave out", () => {
  const documents = [planned("d1", "Release letter", ["Kestrel Works"])]
  const EXCLUDE = "Prepare only documents Kestrel Works executes"
  const ANSWERED = "The committee's approval is outside this package."
  const excluded = act({ document: null, excluded_by: EXCLUDE })
  const WHERE = 'acts[0] ("approve the revised schedule")'
  // The refusal quotes the whole sentences of one lawyer text that contain the words, or says no one text has them.
  const containing = (span: string, several = false) => `In the lawyer's text, the whole sentence${several ? "s" : ""} containing them ${several ? "are" : "is"} "${span}". `
  const NOWHERE = "No one text of the request, the lawyer's updates or answers contains those words exactly as written, capitals included. "
  const notTheLawyers = (words: string, hint: string, where = WHERE) => `${where} gives "excluded_by": "${words}", but those words are not one or more whole sentences of the request, one of the lawyer's updates or one answer, exactly as written. ` +
    hint + "Copy verbatim, with its capitals and from its first word to its last, the whole sentence of the lawyer's own words that leaves this act's document out of this package; " +
    "planned documents and matter documents do not count. Otherwise leave \"excluded_by\" out."
  const directing = (words: string, quote = APPROVE) => `${WHERE} gives "excluded_by": "${words}", which overlaps the words that direct this act ("${quote}"). ` +
    "\"excluded_by\" cannot be the words that direct the act or a sentence containing them: copy verbatim another whole sentence of the request, one update or one answer " +
    "that leaves this act's document out of this package, or leave \"excluded_by\" out."
  const accepted = (directory: string, words: string) =>
    expect(parseCoverage(reply({ document: null, excluded_by: words }), directory, documents)).toEqual({ issues: [], value: [act({ document: null, excluded_by: words })] })
  /** Refused as not whole sentences; the sentences the refusal quotes, given back, pass that check. */
  const refused = (directory: string, words: string) => {
    const parsed = parseCoverage(reply({ document: null, excluded_by: words }), directory, documents)
    const [, several, span] = /In the lawyer's text, the whole sentence(s?) containing them (?:is|are) "(.*)"\. Copy verbatim/.exec(parsed.issues[0] ?? "") ?? []
    // Words in several places are shown the sentences of each.
    const [places, list] = /In the lawyer's text, \d+ places contain them; the whole sentences there are (".*?")(?: and \d+ more not quoted here)?\. (?=Copy verbatim)/.exec(parsed.issues[0] ?? "") ?? []
    expect(parsed).toEqual({ issues: [notTheLawyers(words, span !== undefined ? containing(span, !!several) : places ?? NOWHERE)] })
    for (const whole of span !== undefined ? [span] : list?.slice(1, -1).split(/", "|" and "/) ?? [])
      expect(parseCoverage(reply({ document: null, excluded_by: whole }), directory, documents).issues.filter(issue => issue.includes("not one or more whole sentences"))).toEqual([])
  }

  function excluding() {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    return directory
  }

  test("is accepted with the lawyer's own words that leave its document out, from the request or an answer", () => {
    const directory = excluding()
    expect(parseCoverage(reply({ document: null, excluded_by: ` ${EXCLUDE} ` }), directory, documents)).toEqual({ issues: [], value: [excluded] })
    writeQuestions(directory, [{ id: "q-1", stage: "intake", question: "Should the package carry the approval?", context: "Needed to plan.", status: "answered", answer: ANSWERED }])
    accepted(directory, "The committee's approval is outside this package")
    // Without the field, or with it null or blank, the act is listed as before.
    for (const excluded_by of [undefined, null, " "])
      expect(parseCoverage(JSON.stringify({ acts: [{ ...act({ document: null }), excluded_by }] }), directory, documents)).toEqual({ issues: [], value: [act({ document: null })] })
  })

  test("is retried when its words are not the lawyer's own, are not a string, or come with a planned document", () => {
    const directory = excluding()
    expect(parseCoverage(reply({ document: null, excluded_by: "Prepare only the release letter" }), directory, documents))
      .toEqual({ issues: [notTheLawyers("Prepare only the release letter", NOWHERE)] })
    // Words found only in a matter document are not the lawyer's.
    expect(parseCoverage(reply({ document: null, excluded_by: RELEASE }), directory, documents)).toEqual({ issues: [notTheLawyers(RELEASE, NOWHERE)] })
    // Nor are the plan's own words.
    const written = "Kestrel Works alone signs this letter"
    expect(parseCoverage(reply({ document: null, excluded_by: written }), directory, [{ ...documents[0]!, purpose: written }])).toEqual({ issues: [notTheLawyers(written, NOWHERE)] })
    expect(parseCoverage(JSON.stringify({ acts: [{ ...act({ document: null }), excluded_by: 5 }] }), directory, documents)).toEqual({
      issues: [`${WHERE} gives "excluded_by" that is not a string. Give the whole sentence of the request, one of the lawyer's updates or one answer that leaves this act's document out of this package, verbatim, or leave "excluded_by" out.`],
    })
    expect(parseCoverage(reply({ excluded_by: EXCLUDE }), directory, documents)).toEqual({
      issues: [`${WHERE} gives "excluded_by" with the document "d1". An act whose document the instructions leave out of this package is carried by no planned document: ` +
        'give "document": null with "excluded_by", or the planned document without it.'],
    })
  })

  test("refused as not whole sentences, it is shown the whole sentences of one lawyer text that contain its words, or told no one text has them as written", () => {
    const directory = excluding()
    const shown = (words: string, hint: string) =>
      expect(parseCoverage(reply({ document: null, excluded_by: words }), directory, documents)).toEqual({ issues: [notTheLawyers(words, hint)] })
    shown("Prepare only documents", containing(`${EXCLUDE}.`))
    shown("then releases the deposit. Prepare only documents Kestrel Works executes", containing(`Kestrel Works then releases the deposit. ${EXCLUDE}.`, true))
    // The check is verbatim and case-sensitive, and reads one text at a time.
    shown(EXCLUDE.toLowerCase(), NOWHERE)
    writeQuestions(directory, [{ id: "q-1", stage: "intake", question: "Should the package carry the approval?", context: "Needed to plan.", status: "answered", answer: `No. ${ANSWERED}` }])
    shown("approval is outside this package", containing(ANSWERED))
    shown(`${EXCLUDE}. ${ANSWERED}`, NOWHERE)
    // A hard-wrapped sentence is shown on one line, as the check reads it; a long one is quoted whole, never cut (refusal hint: long sentence).
    const long = `We will not prepare ${"the committee's minutes, ".repeat(12)}or its approval in this package.`
    writeLeanText(directory, "request.md", `${REQUEST}\nWe will not prepare the committee's\napproval in this package.\n${long}`)
    shown("We will not prepare the committee's", "In the lawyer's text, 2 places contain them; the whole sentences there are " +
      "\"We will not prepare the committee's approval in this package.\" and 1 more not quoted here. ")
    shown("or its approval in this package", containing(long))
    accepted(directory, long)
  })

  test("is accepted only as whole sentences of the request, ending at . ! ? … (with any closing quotation mark, bracket or emphasis) or ;, never as a fragment or single words", () => {
    const directory = excluding()
    accepted(directory, EXCLUDE)
    accepted(directory, `${EXCLUDE}.`)
    accepted(directory, `Kestrel Works then releases the deposit. ${EXCLUDE}`)
    for (const words of ["the", "Prepare", "Prepare only documents", "only documents Kestrel Works executes", "then releases the deposit. Prepare only documents Kestrel Works executes"])
      refused(directory, words)
    writeLeanText(directory, "request.md", `${REQUEST}\nLeave the committee's approval out;  Kestrel Works signs alone.\nPrepare what the note says: "Use the release letter only." Keep it short.`)
    for (const words of ["Leave the committee's approval out", "Leave the committee's approval out;", "Kestrel Works signs alone", `Prepare what the note says: "Use the release letter only."`,
      "Prepare what the note says: \"Use the release letter only"])
      accepted(directory, words)
    for (const words of ["Leave the committee's approval", "approval out; Kestrel Works signs alone", "Use the release letter only.", "Prepare what the note says:"])
      refused(directory, words)
  })

  test("is retried when its words change the case of the lawyer's sentence or run across two of the lawyer's texts", () => {
    const directory = excluding()
    refused(directory, EXCLUDE.toUpperCase())
    refused(directory, EXCLUDE.toLowerCase())
    writeQuestions(directory, [{ id: "q-1", stage: "intake", question: "Should the package carry the approval?", context: "Needed to plan.", status: "answered", answer: ANSWERED }])
    refused(directory, "the committee's approval is outside this package")
    refused(directory, `${EXCLUDE}. ${ANSWERED}`)
  })

  test("is accepted as a whole sentence of one included update, not as a fragment of it", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", REQUEST)
    writeLeanText(directory, "requests/done/000000000000001-000000-aaaaaaaa.json",
      `${JSON.stringify({ kind: "update", text: `Keep the letter short. ${EXCLUDE}.`, id: "000000000000001-000000-aaaaaaaa", outcome: { status: "included" } })}\n`)
    accepted(directory, EXCLUDE)
    accepted(directory, `Keep the letter short. ${EXCLUDE}.`)
    refused(directory, "documents Kestrel Works executes")
    refused(directory, "short. Prepare only documents Kestrel Works executes")
  })

  test("is accepted as a whole sentence of one answer, not as a fragment of it", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", REQUEST)
    writeQuestions(directory, [{ id: "q-1", stage: "intake", question: "Should the package carry the approval?", context: "Needed to plan.", status: "answered", answer: `No. ${ANSWERED}` }])
    accepted(directory, ANSWERED)
    accepted(directory, `No. ${ANSWERED}`)
    refused(directory, "approval is outside this package")
    refused(directory, "No. The committee's approval")
  })

  test("is accepted as two whole sentences, not as one and a fragment of the next", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}. ${ANSWERED}`)
    accepted(directory, `${EXCLUDE}. ${ANSWERED}`)
    accepted(directory, `${EXCLUDE}.\n ${ANSWERED.slice(0, -1)}`)
    refused(directory, `${EXCLUDE}. The committee's approval`)
  })

  const LEAVE = "Leave the committee's approval out"

  test("is accepted as the whole sentence of a list item, heading or blockquote line, with or without the line's marker; a marker inside a line starts no sentence", () => {
    const directory = matter(SOURCE_TEXT)
    for (const marker of ["- ", "* ", "• ", "+ ", "· ", "– ", "— ", "◦ ", "▪ ", "1. ", "1) ", "1.1 ", "2.3.1 ", "a. ", "(a) ", "a) ", "(A) ", "(iv) ", "IV) ",
      "## ", "> ", "> - "]) {
      writeLeanText(directory, "request.md", `${REQUEST}\nOut of scope\n${marker}${LEAVE}.\nKeep it short.`)
      for (const words of [LEAVE, `${LEAVE}.`, `${marker}${LEAVE}.`]) accepted(directory, words)
    }
    // The next list item ends the sentence.
    writeLeanText(directory, "request.md", `${REQUEST}\nOut of scope\n- ${LEAVE}\n- Keep it short`)
    accepted(directory, LEAVE)
    // A capital and "." open a name, not a list item: the line carries on the line above.
    writeLeanText(directory, "request.md", `${REQUEST}\nSend the release letter to\nA. Moss for signature.`)
    refused(directory, "Send the release letter to")
    for (const marker of ["- ", "– ", "(a) ", "1) "]) {
      writeLeanText(directory, "request.md", `${REQUEST}\nOut of scope ${marker}${LEAVE}.`)
      refused(directory, LEAVE)
    }
    // A blockquote line starts a sentence only after a sentence end; one that carries on the line above starts none.
    writeLeanText(directory, "request.md", `${REQUEST}\n> Keep it short.\n> ${LEAVE}.\n> Leave out the consent of Orchard Ltd.\n> or its parent.\n> Leave the committee's\n> approval out.`)
    accepted(directory, LEAVE)
    for (const words of ["or its parent.", "Leave the committee's", "approval out."]) refused(directory, words)
  })

  test("reads a line opening with a decimal amount before a lowercase word as a hard wrap, not a multi-level number", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST}\nRelease a deposit of\n3.5 million once the schedule is approved.\nKeep the fee at\n2.25 per cent of the price.\n1.1 ${LEAVE}.`)
    for (const words of ["Release a deposit of 3.5 million once the schedule is approved.", "Keep the fee at 2.25 per cent of the price.", LEAVE]) accepted(directory, words)
    for (const words of ["Release a deposit of", "3.5 million once the schedule is approved.", "Keep the fee at", "2.25 per cent of the price."]) refused(directory, words)
  })

  test("is accepted as a whole sentence without the emphasis, quotation marks, brackets or ellipsis around it, but not as words wrapped inside a sentence", () => {
    const directory = matter(SOURCE_TEXT)
    for (const [open, close] of [["**", "**"], ["_", "_"], ['"', '"'], ["\u201C", "\u201D"], ["'", "'"], ["(", ")"], ["[", "]"], ["*\u201C", "\u201D*"]]) {
      writeLeanText(directory, "request.md", `${REQUEST} ${open}${LEAVE}.${close} Keep it short.`)
      for (const words of [LEAVE, `${LEAVE}.`, `${open}${LEAVE}.${close}`]) accepted(directory, words)
    }
    writeLeanText(directory, "request.md", `${REQUEST} \u2026${LEAVE}\u2026 Keep it short.`)
    for (const words of [LEAVE, `${LEAVE}\u2026`]) accepted(directory, words)
    writeLeanText(directory, "request.md", `${REQUEST} The note says **${LEAVE}** for now. It says "${LEAVE}" too. ${LEAVE}\u2026 and keep it short.`)
    for (const words of [LEAVE, `**${LEAVE}**`, `"${LEAVE}"`, `${LEAVE}\u2026`]) refused(directory, words)
  })

  test("is accepted across a hard wrap, never as a fragment that starts or stops at one, or after an abbreviation or a numbered label", () => {
    for (const lineBreak of ["\n", "\r\n", "\r"]) {
      const directory = matter(SOURCE_TEXT)
      writeLeanText(directory, "request.md", [REQUEST,
        "Please tell me whether we should",
        "leave the committee's approval out of this package.",
        "We will not prepare the committee's",
        "approval in this package.",
        "Leave out the consent of Orchard Ltd. or its parent.",
        "Keep everything, e.g. the committee's approval, in this package.",
        "Keep the rest, i.e. \u201Cthe committee's minutes\u201D, in this package.",
        "Resolution No. 3 stays out.",
        "Out of scope:",
        LEAVE,
        "",
        "Keep it short",
        "## Left out",
        "The committee's approval",
      ].join(lineBreak))
      for (const words of ["We will not prepare the committee's approval in this package.", "Leave out the consent of Orchard Ltd. or its parent.",
        "Keep everything, e.g. the committee's approval, in this package.", "Resolution No. 3 stays out.", LEAVE, "Keep it short", "The committee's approval"])
        accepted(directory, words)
      for (const words of ["leave the committee's approval out of this package.", "Please tell me whether we should", "We will not prepare the committee's",
        "approval in this package.", "or its parent.", "Leave out the consent of Orchard Ltd.", "the committee's approval, in this package.", "Keep everything, e.g.",
        "\"the committee's minutes\", in this package.", "3 stays out.", "Resolution No."])
        refused(directory, words)
    }
  })

  test("is accepted as a whole sentence before or after an inline list label; a number after an abbreviation still carries the sentence on", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} Two points: 1. ${LEAVE}. 2. Keep it short.`)
    for (const words of [LEAVE, `${LEAVE}.`, "Keep it short.", "2. Keep it short."]) accepted(directory, words)
    writeLeanText(directory, "request.md", `${REQUEST} 1) ${LEAVE}. 2) Keep it short.`)
    for (const words of ["Kestrel Works then releases the deposit.", `1) ${LEAVE}.`, "2) Keep it short."]) accepted(directory, words)
    writeLeanText(directory, "request.md", `${REQUEST} Leave out cl. 3 of the deed. Keep the deed (see cl. 3) and the minutes.`)
    accepted(directory, "Leave out cl. 3 of the deed.")
    for (const words of ["Leave out cl.", "3 of the deed.", "Keep the deed (see cl.", "3) and the minutes."]) refused(directory, words)
  })

  test("ends a sentence at ! or ? even before a lowercase word; only a . or … run carries on", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} Is the approval needed? leave the committee's approval out! keep it short.`)
    for (const words of ["Is the approval needed?", "leave the committee's approval out", "leave the committee's approval out!", "keep it short."]) accepted(directory, words)
  })

  test("reads a line's trailing spaces (a Markdown hard line break) as no boundary of their own", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST}  \nPlease tell me whether we should  \nleave it out.  \n> Use the release letter only.  \n> ${LEAVE}.  \n> Keep it short.`)
    for (const words of ["Please tell me whether we should leave it out.", LEAVE, `${LEAVE}.`, `> ${LEAVE}.`, "Keep it short."]) accepted(directory, words)
    for (const words of ["Please tell me whether we should", "leave it out."]) refused(directory, words)
  })

  test("starts a sentence on the line after one ending with \":\" inside closing emphasis, quotation marks or brackets", () => {
    const directory = matter(SOURCE_TEXT)
    for (const lead of ["**Out of scope:**", "_Out of scope:_", "\"Out of scope:\"", "“Out of scope:”", "'Out of scope:'", "‘Out of scope:’", "(Out of scope:)", "[Out of scope:]"]) {
      writeLeanText(directory, "request.md", `${REQUEST}\n${lead}\n${LEAVE}\n\nKeep it short`)
      accepted(directory, LEAVE)
    }
  })

  test("is retried when it is, or contains, the words that direct the act", () => {
    const directory = excluding()
    // The act's own directing words are a fragment of the request.
    refused(directory, APPROVE)
    const first = "Prepare the papers so that the Orchard Committee approves the revised schedule."
    expect(parseCoverage(reply({ document: null, excluded_by: first }), directory, documents)).toEqual({ issues: [directing(first)] })
    expect(parseCoverage(reply({ document: null, quote: first, excluded_by: first }), directory, documents)).toEqual({ issues: [directing(first, first)] })
    // A sentence inside the directing words, and directing words verified without regard to case.
    expect(parseCoverage(reply({ document: null, quote: REQUEST, excluded_by: "Kestrel Works then releases the deposit." }), directory, documents))
      .toEqual({ issues: [directing("Kestrel Works then releases the deposit.", REQUEST)] })
    expect(parseCoverage(reply({ document: null, quote: APPROVE.toUpperCase(), excluded_by: first }), directory, documents)).toEqual({ issues: [directing(first, APPROVE.toUpperCase())] })
  })

  test("listed both with and without excluded_by for the same parties is retried, naming the entries", () => {
    const directory = excluding()
    const conflict = (all: string, withIt: string, without: string) =>
      `${all} describe an act the same way ("approve the revised schedule") for the same parties: with "excluded_by" in ${withIt}, without it in ${without}. ` +
      "If these are different acts, describe each so that the descriptions differ. If they are the same act, keep one entry for it: " +
      "with \"excluded_by\" and \"document\": null only when the lawyer's own words leave its document out of this package, otherwise without \"excluded_by\"."
    expect(parseCoverage(reply({ document: null, excluded_by: EXCLUDE }, {}), directory, documents)).toEqual({ issues: [conflict("acts[0] and acts[1]", "acts[0]", "acts[1]")] })
    // Uncovered, differing only in spacing, case and party naming.
    expect(parseCoverage(reply({ document: null }, { act: " Approve  the revised SCHEDULE ", parties: ["Orchard committee"], document: null, excluded_by: EXCLUDE }), directory, documents))
      .toEqual({ issues: [conflict("acts[0] and acts[1]", "acts[1]", "acts[0]")] })
    expect(parseCoverage(reply({}, { document: null, excluded_by: EXCLUDE }, { document: null }), directory, documents))
      .toEqual({ issues: [conflict("acts[0], acts[1] and acts[2]", "acts[1]", "acts[0] and acts[2]")] })
    // Other parties make another act.
    expect(parseCoverage(reply({ document: null, excluded_by: EXCLUDE }, { parties: ["Kestrel Works"] }), directory, documents))
      .toEqual({ issues: [], value: [excluded, act({ parties: ["Kestrel Works"] })] })
    // Two different acts described the same way are asked to be told apart, not merged.
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}. The Orchard Committee also approves the budget.`)
    expect(parseCoverage(reply({}, { quote: "The Orchard Committee also approves the budget", document: null, excluded_by: EXCLUDE }), directory, documents))
      .toEqual({ issues: [conflict("acts[0] and acts[1]", "acts[1]", "acts[0]")] })
  })

  test("on the last attempt, words still refused are dropped, so the act is held to the plan and asked about like any act; other defects still fail", () => {
    const directory = excluding()
    const last = (text: string) => parseCoverage(text, directory, documents, { final: true })
    const first = "Prepare the papers so that the Orchard Committee approves the revised schedule."
    // Not whole sentences, the words that direct the act, not a string.
    for (const text of [reply({ document: null, excluded_by: "Prepare only documents" }), reply({ document: null, excluded_by: first }),
      JSON.stringify({ acts: [{ ...act({ document: null }), excluded_by: 5 }] })]) {
      expect(parseCoverage(text, directory, documents, { final: false }).issues).toHaveLength(1)
      const parsed = last(text)
      expect(parsed).toEqual({ issues: [], value: [act({ document: null })] })
      expect(planCoverageIssues(parsed.value!, documents, directory)).toEqual([
        `The instructions direct ${DIRECTED}, but no planned document carries it. Plan it in a document executed by "the Orchard Committee", adding one when none is planned.`,
      ])
      expect(coverageScopeQuestion(parsed.value!, documents)?.id).toBe(scopeID(act()))
    }
    // Given with a planned document, the act is held to that document.
    const carried = last(reply({ excluded_by: EXCLUDE }))
    expect(carried).toEqual({ issues: [], value: [act()] })
    expect(planCoverageIssues(carried.value!, documents, directory)).toEqual([NOT_EXECUTED_BY_KESTREL])
    // Words that pass are kept; any other defect still fails the reply.
    expect(last(reply({ document: null, excluded_by: EXCLUDE }))).toEqual({ issues: [], value: [excluded] })
    expect(last(reply({ document: null, excluded_by: "Prepare only documents", parties: ["Lantern Group"] }))).toEqual({ issues: [unnamedParty("Lantern Group")] })
  })

  test("on the last attempt, an act listed both with and without excluded_by for the same parties, in no planned document, is kept from the entries without it", () => {
    const directory = excluding()
    const last = (text: string) => parseCoverage(text, directory, documents, { final: true })
    const open = last(reply({ document: null }, { act: " Approve  the revised SCHEDULE ", parties: ["Orchard committee"], document: null, excluded_by: EXCLUDE }))
    expect(open).toEqual({ issues: [], value: [act({ document: null })] })
    expect(planCoverageIssues(open.value!, documents, directory)).toEqual([
      `The instructions direct ${DIRECTED}, but no planned document carries it. Plan it in a document executed by "the Orchard Committee", adding one when none is planned.`,
    ])
    // Other parties make another act, which keeps its words.
    expect(last(reply({ document: null, excluded_by: EXCLUDE }, { parties: ["Kestrel Works"] }))).toEqual({ issues: [], value: [excluded, act({ parties: ["Kestrel Works"] })] })
  })

  test("listed twice with different words that leave its document out is kept with each", () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}. ${ANSWERED}`)
    expect(parseCoverage(reply({ document: null, excluded_by: EXCLUDE }, { document: null, excluded_by: ANSWERED }), directory, documents))
      .toEqual({ issues: [], value: [excluded, act({ document: null, excluded_by: ANSWERED })] })
  })

  test("raises no revision issue and no package-scope question; an act no planned document carries without it still raises both", () => {
    const directory = excluding()
    expect(planCoverageIssues([excluded], documents, directory)).toEqual([])
    expect(coverageScopeQuestion([excluded], documents)).toBeUndefined()
    const open = act({ act: "release the deposit", source: SOURCE, quote: RELEASE, parties: ["Kestrel Works"], document: null })
    expect(planCoverageIssues([excluded, open], documents, directory)).toEqual([
      `The instructions direct "release the deposit" ("${RELEASE}" [source: ${SOURCE}]), but no planned document carries it. Plan it in a document executed by "Kestrel Works", adding one when none is planned.`,
    ])
    const question = coverageScopeQuestion([excluded, open], documents)!
    expect(question.id).toBe(scopeID(open))
    expect(question.question).toContain('"release the deposit"')
    expect(question.question).not.toContain("approve the revised schedule")
  })

  test("goes back to the outline, never to the lawyer, when a planned document is executed by all its parties; a document only some of them execute raises nothing", () => {
    const directory = excluding()
    const resolution = planned("d2", "Committee resolution", ["the Orchard Committee", "Kestrel Works"])
    expect(planCoverageIssues([excluded], [...documents, resolution], directory)).toEqual([
      `The check read the instructions as leaving out of this package the document for ${DIRECTED}, which must be taken by "the Orchard Committee", ` +
        'but d2 "Committee resolution", executed by "the Orchard Committee" and "Kestrel Works", is planned and carries no instructed act. If the instructions leave that document out, do not plan it.',
    ])
    expect(coverageScopeQuestion([excluded], [...documents, resolution])).toBeUndefined()
    const jointly = act({ document: null, excluded_by: EXCLUDE, parties: ["the Orchard Committee", "Kestrel Works"] })
    expect(planCoverageIssues([jointly], [...documents, planned("d2", "Committee resolution", ["the Orchard Committee"])], directory)).toEqual([])
    expect(planCoverageIssues([act({ document: null, excluded_by: EXCLUDE, parties: [] })], [...documents, resolution], directory)).toEqual([])
  })
})

describe("refusal hint: fix wave 2", () => {
  const documents = [planned("d1", "Release letter", ["Kestrel Works"])]
  const LEAVE = "Leave the committee's approval out"
  const notWhole = (directory: string, words: string) =>
    parseCoverage(reply({ document: null, excluded_by: words }), directory, documents).issues.filter(issue => issue.includes("not one or more whole sentences"))
  /** `words` are refused as not whole sentences with `hint`, and each of `passing` (what the hint quotes), given back, passes that check. */
  const hinted = (directory: string, words: string, hint: string, passing: string[]) => {
    const issues = parseCoverage(reply({ document: null, excluded_by: words }), directory, documents).issues
    expect(issues).toHaveLength(1)
    expect(issues[0]).toContain(`exactly as written. ${hint}Copy verbatim, with its capitals`)
    for (const span of passing) expect(notWhole(directory, span)).toEqual([])
  }
  const one = (span: string, markup = false) => `In the lawyer's text, the whole sentence containing them${markup ? ", with its markup as written," : ""} is "${span}". `
  const places = (spans: string[], more = 0, markup = false) => {
    const listed = [...spans.map(span => `"${span}"`), ...(more ? [`${more} more not quoted here`] : [])]
    return `In the lawyer's text, ${spans.length + more} places contain them; the whole sentences there${markup ? ", with their markup as written," : ""} are ` +
      `${listed.slice(0, -1).join(", ")} and ${listed.at(-1)}. `
  }

  test("words copied without the markup inside the lawyer's sentence are shown that sentence with its markup, which passes; the copy without it stays refused", () => {
    const directory = matter(SOURCE_TEXT)
    const clean = "Leave the committee's approval out of this package."
    for (const [quoted, written] of [["> ", "Leave the committee's approval > out of this package."], ["> > ", "Leave the committee's approval > > out of this package."]]) {
      writeLeanText(directory, "request.md", `${REQUEST}\n${quoted}Leave the committee's approval\n${quoted}out of this package.`)
      hinted(directory, clean, one(written, true), [written])
      // The words are quoted again as the lawyer's (coverageItems), and the quote checks read them verbatim.
      expect([instructionFound(directory, clean), instructionFound(directory, written)]).toEqual([false, true])
    }
    // Markup alone, not in the text as written, is nowhere.
    for (const words of ["**", "_"]) hinted(directory, words, "No one text of the request, the lawyer's updates or answers contains those words exactly as written, capitals included. ", [])
    writeLeanText(directory, "request.md", `${REQUEST} Do *not* prepare the committee's approval. Prepare **only** the release letter. Keep it __short__.`)
    hinted(directory, "Do not prepare the committee's approval.", one("Do *not* prepare the committee's approval.", true), ["Do *not* prepare the committee's approval."])
    hinted(directory, "Prepare only the release letter", one("Prepare **only** the release letter.", true), ["Prepare **only** the release letter."])
    hinted(directory, "Keep it short", one("Keep it __short__.", true), ["Keep it __short__."])
    // A fragment is still a fragment, markup or not.
    hinted(directory, "Prepare only", one("Prepare **only** the release letter.", true), [])
    // Markup the lawyer did not write is shown without it.
    hinted(directory, "Kestrel Works **then** releases the deposit.", one("Kestrel Works then releases the deposit.", true), [])
    // Words as written somewhere are shown only where they are written so.
    writeLeanText(directory, "request.md", `${REQUEST} Leave the approval out of the minutes. Leave the **approval** out of this package.`)
    hinted(directory, "Leave the approval out", one("Leave the approval out of the minutes."), [])
    hinted(directory, "Leave the **approval** out", one("Leave the **approval** out of this package."), ["Leave the **approval** out of this package."])
  })

  test("words found in several sentences are shown each of them, not the first; one sentence that holds them twice is one place", () => {
    const directory = matter(SOURCE_TEXT)
    const MINUTES = "Leave the approval out of the minutes.", PACKAGE = "Leave the approval out of this package."
    writeLeanText(directory, "request.md", `${REQUEST}\n${MINUTES}\n${PACKAGE}`)
    hinted(directory, "Leave the approval out", 'In the lawyer\'s text, 2 places contain them; the whole sentences there are ' +
      `"${MINUTES}" and "${PACKAGE}". `, [MINUTES, PACKAGE])
    hinted(directory, "Leave the approval out of this", one(PACKAGE), [PACKAGE])
    // Words that run one letter into the next line's sentence are shown both sentences.
    hinted(directory, "out of the minutes.\nL", `In the lawyer's text, the whole sentences containing them are "${MINUTES} ${PACKAGE}". `, [`${MINUTES} ${PACKAGE}`])
    // An answer is another place; the same sentence in two texts is one.
    writeQuestions(directory, [{ id: "q-1", stage: "intake", question: "Should the package carry the approval?", context: "Needed to plan.", status: "answered", answer: `${PACKAGE} Leave the approval out of the deed.` }])
    hinted(directory, "Leave the approval out", places([MINUTES, PACKAGE, "Leave the approval out of the deed."]), [MINUTES, PACKAGE, "Leave the approval out of the deed."])
    writeQuestions(directory, [])
    writeLeanText(directory, "request.md", `${REQUEST}\nKeep the approval and the approval minutes out.`)
    hinted(directory, "the approval", one("Keep the approval and the approval minutes out."), [])
    // Past markup too.
    writeLeanText(directory, "request.md", `${REQUEST}\nLeave the *approval* out of the minutes. Leave the _approval_ out of this package.`)
    hinted(directory, "Leave the approval out", places(["Leave the *approval* out of the minutes.", "Leave the _approval_ out of this package."], 0, true),
      ["Leave the *approval* out of the minutes.", "Leave the _approval_ out of this package."])
  })

  test("several places are quoted whole within the excerpt length, in order, and the rest counted", () => {
    const directory = matter(SOURCE_TEXT)
    const numbered = [1, 2, 3, 4, 5, 6].map(n => `Leave the approval out of the schedule numbered ${n} here.`)
    writeLeanText(directory, "request.md", `${REQUEST}\n${numbered.join(" ")}`)
    hinted(directory, "Leave the approval out", places(numbered.slice(0, 4), 2), numbered.slice(0, 4))
    // A place too long to quote whole is counted; a later one that fits is still quoted.
    const long = `We will not prepare ${"the committee's minutes, ".repeat(12)}or its approval in this package.`
    const short = "We will not prepare the committee's approval in this package."
    writeLeanText(directory, "request.md", `${REQUEST}\n${long}\n${short}`)
    hinted(directory, "We will not prepare the committee's", places([short], 1), [short])
    // When none fits, the first is quoted whole, never cut (refusal hint: long sentence).
    writeLeanText(directory, "request.md", `${REQUEST}\n${long}\n${long.replace("We will", "They will")}`)
    hinted(directory, "will not prepare the committee's", places([long], 1), [long])
  })

  test("on a list line, the sentence shown leaves off the marker the words leave off and keeps the one they include; a marker alone is no sentence", () => {
    const directory = matter(SOURCE_TEXT)
    for (const marker of ["- ", "1. ", "(a) ", "> 1. "]) {
      writeLeanText(directory, "request.md", `${REQUEST}\nOut of scope\n${marker}${LEAVE}.\nKeep it short.`)
      hinted(directory, "Leave the committee's", one(`${LEAVE}.`), [`${LEAVE}.`])
      hinted(directory, `${marker}Leave the committee's`, one(`${marker}${LEAVE}.`), [`${marker}${LEAVE}.`])
    }
  })
})

describe("coverage prompt", () => {
  const documents = [planned("d1", "Resolution", ["the Orchard Committee"]), planned("d2", "Release letter", ["Kestrel Works"])]
  const TERMS = [{ heading: "Terms", scope: "" }]
  const plannedView = (prompt: string): unknown => JSON.parse(/<planned_documents>\n[^\n]*\n([\s\S]*)\n<\/planned_documents>/.exec(prompt)![1]!)

  test("the prompt carries the request and the source names", () => {
    const prompt = coveragePrompt(prepared(), documents)
    expect(prompt).toContain(`<request>\n${REQUEST}\n</request>`)
    expect(prompt).toContain("<sources>\nMatter documents (use these exact names with search-document and read-document, and as source names):\n" +
      `- ${SOURCE}\nTo quote the request, the lawyer's updates or answers, use the source name "request".\n</sources>`)
    expect(prompt).toContain(`- ${SHAPE} Reply {"acts":[]} when the instructions direct no legal act.`)
    expect(prompt).not.toContain("<user_updates>")
    expect(prompt).not.toContain("<answers>")
  })

  test("acts come from the lawyer's instructions, and an authorisation is the act of the party who gives it", () => {
    const prompt = coveragePrompt(prepared(), documents)
    expect(prompt).toContain("- List every legal act that the request, the lawyer's updates or answers direct to be taken or documented in this package. " +
      "A matter document directs an act only where the request adopts it as instructions; an obligation a matter document merely describes is not an instructed act. ")
    expect(prompt).toContain("\"parties\": every party who must take the act for it to be effective (for an authorisation, those who give it), each named as the sources name it, or [] when the record does not settle who must take it;")
    expect(prompt).not.toContain("authorise a later act")
    expect(prompt).not.toContain('"party"')
  })

  test("an act whose document the instructions leave out is listed with the lawyer's words that leave it out", () => {
    expect(coveragePrompt(prepared(), documents)).toContain("\n- When the request, the lawyer's updates or answers leave out of this package the document that the parties who must take an act would execute, " +
      'still list the act, with "document": null and "excluded_by": the whole sentence that leaves that document out, copied verbatim from the request, updates or answers, ' +
      'never from a planned document or a matter document. Leave "excluded_by" out of every other act.\n')
  })

  test("the lawyer's later instructions and answered questions appear under their own tags", () => {
    const directory = prepared()
    writeLeanText(directory, "updated-notes.md", "Kestrel Works also returns the keys.\n")
    writeQuestions(directory, [
      { id: "q-1", stage: "intake", question: "Who countersigns?", context: "Needed to plan.", status: "answered", answer: "Wren Atelier countersigns." },
      { id: "q-2", stage: "intake", question: "Which schedule applies?", context: "Needed to plan.", status: "open" },
    ])
    const prompt = coveragePrompt(directory, documents)
    expect(prompt).toContain("<user_updates>\nLater instructions from the lawyer, in arrival order. An explicit later change supersedes the earlier instruction on that point.\n\n" +
      "Kestrel Works also returns the keys.\n</user_updates>")
    expect(prompt).toContain("<answers>\nThe lawyer's answers; they remain controlling:\n- [q-1] Q: Who countersigns?\n  A: Wren Atelier countersigns.\n</answers>")
    expect(prompt).not.toContain("Which schedule applies?")
  })

  test("the plan is shown by id with its executing parties, and with left-out parts only when it has some", () => {
    const withLeftOut: PlannedDocument = { ...planned("d3", "Advice note", []), left_out: [{ part: "Schedule", reason: "Not instructed." }] }
    const emptyLeftOut: PlannedDocument = { ...planned("d4", "Schedule"), left_out: [] }
    expect(plannedView(coveragePrompt(prepared(), [...documents, withLeftOut, emptyLeftOut]))).toStrictEqual([
      { id: "d1", title: "Resolution", purpose: "", executed_by: ["the Orchard Committee"], parts: TERMS },
      { id: "d2", title: "Release letter", purpose: "", executed_by: ["Kestrel Works"], parts: TERMS },
      { id: "d3", title: "Advice note", purpose: "", executed_by: [], parts: TERMS, left_out: [{ part: "Schedule", reason: "Not instructed." }] },
      { id: "d4", title: "Schedule", purpose: "", executed_by: null, parts: TERMS },
    ])
  })
})

describe("final attempt: fix wave 2", () => {
  // On the coverage job's last attempt, an act listed both with and without
  // words that pass the excluded_by check never loses those words: beside an
  // entry in a planned document they are kept and disclosed with it; only when
  // no entry names a planned document does the act take the uncovered route.
  const EXCLUDE = "Prepare only documents Kestrel Works executes"
  const RESOLUTION = planned("d1", "Committee resolution", ["the Orchard Committee"])
  const LETTER = planned("d2", "Release letter", ["Kestrel Works"])
  const documents = [RESOLUTION, LETTER]
  const excluded = act({ document: null, excluded_by: EXCLUDE })
  const UNCOVERED = `The instructions direct ${DIRECTED}, but no planned document carries it. Plan it in a document executed by "the Orchard Committee", adding one when none is planned.`
  const last = (directory: string, text: string) => parseCoverage(text, directory, documents, { final: true })
  function excluding() {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    return directory
  }
  /** The saved plan and its coverage record as the outline step writes them; the record's open items. */
  function recorded(directory: string, acts: CoverageAct[]) {
    const plan = writeLeanPlan(directory, documents)
    const issues = planCoverageIssues(acts, plan.documents, directory)
    const record: LeanCoverage = { schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: issues.length ? "uncovered" : "covered", acts, issues }
    writeLeanText(directory, "coverage.json", `${JSON.stringify(record, null, 2)}\n`)
    const items = coverageItems(directory)
    const notes = ["# Drafting notes", "", "The committee resolution records the approval.", "", ...items.map(item => `- ${item.detail.replace(/\bd\d+ (?=")/g, "")} [ref: ${item.id}]`)].join("\n")
    expect(checkDraftingNotes(directory, notes, disclosableItems(directory)).issues).toEqual([])
    return items.map(item => item.detail)
  }
  const disclosed = (carried: string) => `Plan coverage: the check of the document plan against the instructions read the lawyer's words "${EXCLUDE}" [source: request] ` +
    "as leaving out the document for this instructed act, yet also read the plan as carrying it: " +
    `"approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; ${carried}. ` +
    "The lawyer should confirm whether this act is to be dealt with in this package as planned or outside it."

  test("checked words contradicted by an entry in a planned document are kept beside it and disclosed with that document, never dropped", () => {
    const directory = excluding()
    const text = reply({ document: null, excluded_by: EXCLUDE }, {})
    expect(parseCoverage(text, directory, documents).issues).toHaveLength(1)
    const parsed = last(directory, text)
    expect(parsed).toEqual({ issues: [], value: [excluded, act()] })
    // The plan carries the act in a document its party executes: no revision, no question.
    expect(planCoverageIssues(parsed.value!, documents, directory)).toEqual([])
    expect(coverageScopeQuestion(parsed.value!, documents)).toBeUndefined()
    // The lawyer's words reach the lawyer with the document the plan has the act in, not as an act no planned document carries.
    expect(recorded(directory, parsed.value!)).toEqual([disclosed('planned in d1 "Committee resolution", executed by "the Orchard Committee"')])
  })

  test("beside entries both in a planned document and in none, every reading is kept: the uncovered one still goes to the outline and the lawyer", () => {
    const directory = excluding()
    const parsed = last(directory, reply({}, { document: null, excluded_by: EXCLUDE }, { document: null }, { document: "d2" }))
    expect(parsed).toEqual({ issues: [], value: [act(), excluded, act({ document: null }), act({ document: "d2" })] })
    const issues = planCoverageIssues(parsed.value!, documents, directory)
    expect(issues).toEqual([UNCOVERED, `${DIRECTED} must be taken by "the Orchard Committee", but d2 "Release letter" is executed by "Kestrel Works". ${move("Move")}`])
    expect(coverageScopeQuestion(parsed.value!, documents)?.id).toBe(scopeID(act()))
    const items = recorded(directory, parsed.value!)
    expect(items[0]).toBe(disclosed('planned in d1 "Committee resolution", executed by "the Orchard Committee", and planned in d2 "Release letter", executed by "Kestrel Works"'))
    expect(items).toHaveLength(3)
  })

  test("contradicted only by entries in no planned document, checked words are dropped: the act goes to the outline and the lawyer", () => {
    const directory = excluding()
    for (const text of [reply({ document: null, excluded_by: EXCLUDE }, { document: null }), reply({ document: null }, { document: null, excluded_by: EXCLUDE }, { document: null })]) {
      const parsed = last(directory, text)
      expect(parsed).toEqual({ issues: [], value: [act({ document: null })] })
      expect(planCoverageIssues(parsed.value!, documents, directory)).toEqual([UNCOVERED])
      expect(coverageScopeQuestion(parsed.value!, documents)?.id).toBe(scopeID(act()))
    }
  })

  test("refused words are still dropped on the last attempt, beside an entry in a planned document or alone", () => {
    const directory = excluding()
    expect(last(directory, reply({ document: null, excluded_by: "Prepare only documents" }, {}))).toEqual({ issues: [], value: [act({ document: null }), act()] })
    expect(last(directory, reply({ document: null, excluded_by: "Prepare only documents" }))).toEqual({ issues: [], value: [act({ document: null })] })
    expect(last(directory, reply({ excluded_by: EXCLUDE }))).toEqual({ issues: [], value: [act()] })
  })
})

describe("sentence bounds: fix wave 2", () => {
  const documents = [planned("d1", "Release letter", ["Kestrel Works"])]
  const LEAVE = "Leave the committee's approval out"
  const checked = (directory: string, words: string) => parseCoverage(reply({ document: null, excluded_by: words }), directory, documents)
  const accepted = (directory: string, ...sentences: string[]) => {
    for (const words of sentences) expect([words, checked(directory, words)]).toEqual([words, { issues: [], value: [act({ document: null, excluded_by: words })] }])
  }
  /** Refused as not whole sentences of the lawyer's text, and for nothing else. */
  const refused = (directory: string, ...fragments: string[]) => {
    for (const words of fragments) {
      const { issues } = checked(directory, words)
      expect([words, issues.length, issues[0]]).toEqual([words, 1, expect.stringContaining("those words are not one or more whole sentences")])
    }
  }
  const request = (...lines: string[]) => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", [REQUEST, ...lines].join("\n"))
    return directory
  }

  test("a number after an abbreviation carries the sentence on unless it is a label of a numbered run in its paragraph", () => {
    const directory = request(
      "Do not prepare the guarantee for Facility No. 2. Facility No. 1 needs one.",
      "Leave out the deed under cl. 4. The Seller will sign it separately.",
      "Leave the notice under s. 146. Please keep the rest.",
      "Leave out the consent (para. 12) The rest stays.",
      "See Sch. 1) The approval stays out.",
    )
    accepted(directory, "Do not prepare the guarantee for Facility No. 2.", "Facility No. 1 needs one.", "Leave out the deed under cl. 4.", "The Seller will sign it separately.",
      "Leave the notice under s. 146.", "Please keep the rest.", "Leave out the consent (para. 12) The rest stays.", "See Sch. 1) The approval stays out.")
    refused(directory, "Do not prepare the guarantee for Facility No.", "2.", "2. Facility No. 1 needs one.", "Leave out the deed under cl.", "4. The Seller will sign it separately.",
      "Leave the notice under s.", "146. Please keep the rest.", "Leave out the consent (para.", "12) The rest stays.", "See Sch.", "1) The approval stays out.")
    // A run in another paragraph labels nothing here.
    const apart = request(`Two points: 1. ${LEAVE}. 2. Keep it short. 3. Sign it.`, "", "Do not prepare the guarantee for Facility No. 4. Facility No. 3 needs one.")
    accepted(apart, LEAVE, "Keep it short.", "2. Keep it short.", "Sign it.", "Do not prepare the guarantee for Facility No. 4.")
    refused(apart, "Do not prepare the guarantee for Facility No.", "4. Facility No. 3 needs one.")
  })

  test("numbers that close sentences after abbreviations are no run of labels, even in sequence; a run opens after a line break or a colon, or as n)", () => {
    const directory = request("Leave out Facility No. 1. Keep Facility No. 2. The rest stays.")
    accepted(directory, "Leave out Facility No. 1.", "Keep Facility No. 2.", "The rest stays.", "Leave out Facility No. 1. Keep Facility No. 2.")
    refused(directory, "Leave out Facility No.", "Keep Facility No.", "1. Keep Facility No. 2.", "2. The rest stays.")
    const runs = request("Out of scope", `1. ${LEAVE}. 2. Keep it short.`, "Kestrel Works signs. (1) Send it. (2) File it.", "Kestrel Works dates it. 1) Post it. 2) Note it.")
    accepted(runs, LEAVE, `${LEAVE}.`, "Keep it short.", "Kestrel Works signs.", "(1) Send it.", "(2) File it.", "Kestrel Works dates it.", "1) Post it.", "2) Note it.")
    // A number before a lowercase word labels nothing, even in a run.
    const lower = request("Keep the deed (see cl. 1) and the minutes (see cl. 2) and the plan.")
    accepted(lower, "Keep the deed (see cl. 1) and the minutes (see cl. 2) and the plan.")
    refused(lower, "Keep the deed (see cl.", "1) and the minutes (see cl.")
    // Kept refused: after a "." the first label could equally close the sentence before it, so the run opens nowhere.
    const after = request(`Note two things. 1. ${LEAVE}. 2. Keep it short.`)
    accepted(after, `Note two things. 1. ${LEAVE}. 2. Keep it short.`, "Keep it short.")
    refused(after, "Note two things.", LEAVE, `${LEAVE}.`)
  })

  test("a letter with \".\" opens a list item only before a word that is not a number or lowercase, as a list's first letter or the next one", () => {
    const directory = request(
      "Leave out the notice under", "s. 146 of the Act for now.",
      "Leave out the consent noted at", "p. 5 of the minutes.",
      "Leave out the settlement in Moss", "v. Wren for now.",
    )
    accepted(directory, "Leave out the notice under s. 146 of the Act for now.", "Leave out the consent noted at p. 5 of the minutes.", "Leave out the settlement in Moss v. Wren for now.")
    refused(directory, "Leave out the notice under", "s. 146 of the Act for now.", "146 of the Act for now.", "Leave out the consent noted at", "5 of the minutes.",
      "Leave out the settlement in Moss")
    const lists = request("Out of scope", `a. Leave the committee's`, "approval out", "b. Keep it short", "c. Sign it.", "Also", `i. ${LEAVE}.`)
    accepted(lists, LEAVE, "Keep it short", "b. Keep it short", "Sign it.", `${LEAVE}.`, `i. ${LEAVE}.`)
    refused(lists, "Leave the committee's", "approval out")
    // Kept refused: an item that opens with a number is read as a wrap, with its letter, carrying on the item above.
    const opening = request("Out of scope:", "a. leave the deed out.", "b. 5 copies go to the Lender.")
    accepted(opening, "a. leave the deed out. b. 5 copies go to the Lender.")
    refused(opening, "leave the deed out.", "5 copies go to the Lender.")
  })

  test("a multi-level number without a closing \".\" opens a line only before a capitalised word", () => {
    const directory = request(
      "Leave out the loan agreement at a rate of", "4.25 % above base for now.",
      "Leave out the deed referred to in clause", "4.2 (Conditions) of the SPA.",
      "Do not prepare the deed dated", "1.10.2026 (the Completion Date) yet.",
      "Do not draft the escrow for", "2.5 EUR million until the bank agrees.",
      `1.1 ${LEAVE}.`, `2.3.1 I will sign it.`,
    )
    accepted(directory, "Leave out the loan agreement at a rate of 4.25 % above base for now.", "Leave out the deed referred to in clause 4.2 (Conditions) of the SPA.",
      "Do not prepare the deed dated 1.10.2026 (the Completion Date) yet.", "Do not draft the escrow for 2.5 EUR million until the bank agrees.", LEAVE, "I will sign it.")
    refused(directory, "Leave out the loan agreement at a rate of", "4.25 % above base for now.", "Leave out the deed referred to in clause", "(Conditions) of the SPA.",
      "Do not prepare the deed dated", "(the Completion Date) yet.", "Do not draft the escrow for", "EUR million until the bank agrees.")
  })

  test("a dash at the start of a line opens a list item after a blank line, a sentence end, a \":\" or a heading, after an item with the same dash, or before a word that is not lowercase", () => {
    for (const dash of ["-", "–", "—"]) {
      const directory = request(
        "Leave the guarantee out of this package", `${dash} unless the Lender insists on it.`, "",
        "Out of scope", `${dash} ${LEAVE}`, `${dash} keep it short`, "",
        "Out of scope:", `${dash} sign it`, "",
        "Keep the letter", `${dash} send it`,
      )
      accepted(directory, `Leave the guarantee out of this package ${dash} unless the Lender insists on it.`, LEAVE, "keep it short", `${dash} keep it short`, "sign it",
        "Keep the letter send it".replace(" send", ` ${dash} send`))
      refused(directory, "Leave the guarantee out of this package", "unless the Lender insists on it.", "Keep the letter", "send it")
    }
    // Another dash is no item of the list: the wrapped words carry on the item above.
    const mixed = request("Out of scope", `- ${LEAVE}`, "– unless the Lender insists on it.")
    accepted(mixed, `${LEAVE} – unless the Lender insists on it.`)
    refused(mixed, LEAVE, "unless the Lender insists on it.")
  })
})

describe("dedup key", () => {
  // Entries for the same act and parties are separate instructions when other
  // words direct it (another source or quote) or other lawyer words leave its
  // document out: each is kept, so a checklist ref to either quote finds its act
  // and each instruction's words reach the lawyer. Only a true duplicate (same
  // act, parties, document, source, quote and excluded_by words, apart from the
  // quote checks' spacing) is kept once.
  const SIGN_OFF = "The Orchard Committee must sign off the revised schedule"
  const EXCLUDE = "Prepare only documents Kestrel Works executes"
  const OUTSIDE = "The committee's approval is outside this package"
  const documents = [planned("d1", "Committee resolution", ["the Orchard Committee"]), planned("d2", "Release letter", ["Kestrel Works"])]
  function instructed() {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} ${SIGN_OFF}. ${EXCLUDE}. ${OUTSIDE}.`)
    return directory
  }
  /** The saved plan and its coverage record as the outline step writes them. */
  function recorded(directory: string, acts: CoverageAct[]) {
    const plan = writeLeanPlan(directory, documents)
    const issues = planCoverageIssues(acts, plan.documents, directory)
    const record: LeanCoverage = { schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: issues.length ? "uncovered" : "covered", acts, issues }
    writeLeanText(directory, "coverage.json", `${JSON.stringify(record, null, 2)}\n`)
  }

  test("the same act for the same parties directed from another source or by another quote is kept once per instruction, and a ref to either quote finds its act", () => {
    const directory = instructed()
    expect(parseCoverage(reply({}, { source: SOURCE }), directory, documents)).toEqual({ issues: [], value: [act(), act({ source: SOURCE })] })
    const parsed = parseCoverage(reply({}, { quote: SIGN_OFF }), directory, documents)
    expect(parsed).toEqual({ issues: [], value: [act(), act({ quote: SIGN_OFF })] })
    recorded(directory, parsed.value!)
    expect(coverageCarries(directory, [{ document: "request", quote: SIGN_OFF }], "d1")).toBe(true)
    expect(coverageCarries(directory, [{ document: "request", quote: APPROVE }, { document: "request", quote: SIGN_OFF }], "d1")).toBe(true)
  })

  test("an act no planned document carries, directed by two instructions, is disclosed with each instruction's words under its own id", () => {
    const directory = instructed()
    const parsed = parseCoverage(reply({ document: null }, { document: null, quote: SIGN_OFF }), directory, documents)
    expect(parsed).toEqual({ issues: [], value: [act({ document: null }), act({ document: null, quote: SIGN_OFF })] })
    recorded(directory, parsed.value!)
    const items = coverageItems(directory)
    expect(items.map(item => item.detail)).toEqual([expect.stringContaining(`directed by "${APPROVE}"`), expect.stringContaining(`directed by "${SIGN_OFF}"`)])
    expect(new Set(items.map(item => item.id)).size).toBe(2)
  })

  test("an act whose document two different sentences of the lawyer's leave out is kept with each, and each reaches the lawyer", () => {
    const directory = instructed()
    const parsed = parseCoverage(reply({ document: null, excluded_by: EXCLUDE }, { document: null, excluded_by: OUTSIDE }), directory, documents)
    expect(parsed).toEqual({ issues: [], value: [act({ document: null, excluded_by: EXCLUDE }), act({ document: null, excluded_by: OUTSIDE })] })
    recorded(directory, parsed.value!)
    expect(coverageItems(directory).map(item => item.detail)).toEqual([expect.stringContaining(`words "${EXCLUDE}"`), expect.stringContaining(`words "${OUTSIDE}"`)])
  })

  test("two acts whose issue reads the same give the outline that line once, in first-occurrence order beside the other issues", () => {
    const directory = instructed()
    const release = { act: "release the deposit", source: SOURCE, quote: RELEASE, parties: ["Kestrel Works"], document: null }
    const parsed = parseCoverage(reply({ document: null, excluded_by: EXCLUDE }, release, { document: null, excluded_by: OUTSIDE }), directory, documents)
    expect(parsed.value).toHaveLength(3)
    expect(planCoverageIssues(parsed.value!, documents, directory)).toEqual([
      `The check read the instructions as leaving out of this package the document for ${DIRECTED}, which must be taken by "the Orchard Committee", ` +
        'but d1 "Committee resolution", executed by "the Orchard Committee", is planned and carries no instructed act. If the instructions leave that document out, do not plan it.',
      `The instructions direct "release the deposit" ("${RELEASE}" [source: ${SOURCE}]), but no planned document carries it. Plan it in a document executed by "Kestrel Works", adding one when none is planned.`,
    ])
    // Each act still reaches the lawyer in its own item.
    recorded(directory, parsed.value!)
    expect(coverageItems(directory).map(item => item.detail)).toEqual([
      expect.stringContaining(`words "${EXCLUDE}"`), expect.stringContaining(`words "${OUTSIDE}"`), expect.stringContaining('"release the deposit", directed by'),
    ])
  })

  test("a true duplicate is kept once, whatever its spacing", () => {
    const directory = instructed()
    expect(parseCoverage(reply({}, {}), directory, documents)).toEqual({ issues: [], value: [act()] })
    expect(parseCoverage(reply({ quote: SIGN_OFF }, { quote: SIGN_OFF.replaceAll(" ", "  ") }), directory, documents)).toEqual({ issues: [], value: [act({ quote: SIGN_OFF })] })
    const left = act({ document: null, excluded_by: EXCLUDE })
    expect(parseCoverage(reply(left, { ...left, excluded_by: EXCLUDE.replace(" Kestrel", "\n Kestrel") }), directory, documents)).toEqual({ issues: [], value: [left] })
  })
})

describe("refusal hint: long sentence", () => {
  // A refusal hint quotes each sentence it offers whole, however long, since a
  // sentence cut to the excerpt length is refused again when given back. The
  // hint stays bounded by how many places it quotes; a sentence past the longest
  // it quotes (2000 characters) is pointed to in the lawyer's text, never cut.
  const documents = [planned("d1", "Release letter", ["Kestrel Works"])]
  const WORDS = "please leave those out of this package"
  const LONG = "The approval of the revised schedule, the minutes that record it, the notice that goes to the members, the certificate from the secretary " +
    "and any consent the lenders may require for it are matters the committee will see to itself through its own advisers, so please leave those out of this package."
  const FIND = "find it where those words stand in the request, the lawyer's updates or answers. "
  /** One sentence of exactly `length` characters that holds WORDS. */
  const sized = (length: number) => {
    const start = "The committee will see to ", end = ", so please leave those out of this package."
    const pad = length - start.length - end.length, words = Math.floor((pad - 1) / 5)
    return `${start}${"word ".repeat(words)}${"w".repeat(pad - 5 * words)}${end}`
  }
  const refusal = (directory: string, words: string) => {
    const issues = parseCoverage(reply({ document: null, excluded_by: words }), directory, documents).issues
    expect(issues).toHaveLength(1)
    return issues[0]!
  }
  /** The hint between the refusal's opening and its closing instruction. */
  const hint = (issue: string) => /exactly as written\. (.*?)Copy verbatim, with its capitals/s.exec(issue)?.[1]
  /** What the hint offers to copy: each quoted string in it. */
  const offered = (issue: string) => [...(hint(issue) ?? "").matchAll(/"([^"]*)"/g)].map(match => match[1]!)
  const accepted = (directory: string, words: string) =>
    expect(parseCoverage(reply({ document: null, excluded_by: words }), directory, documents)).toEqual({ issues: [], value: [act({ document: null, excluded_by: words })] })

  test("one sentence longer than the excerpt length is quoted whole, and given back it is accepted", () => {
    const directory = matter(SOURCE_TEXT)
    expect(LONG.length).toBeGreaterThan(EXCERPT)
    writeLeanText(directory, "request.md", `${REQUEST} Prepare the release letter. ${LONG} Keep the deed.`)
    const issue = refusal(directory, WORDS)
    // What the hint offers, given back, is accepted.
    for (const whole of offered(issue)) accepted(directory, whole)
    expect(hint(issue)).toBe(`In the lawyer's text, the whole sentence containing them is "${LONG}". `)
    // Words running into the next sentence are shown both, whole.
    expect(hint(refusal(directory, "out of this package. Keep"))).toBe(`In the lawyer's text, the whole sentences containing them are "${LONG} Keep the deed.". `)
    accepted(directory, `${LONG} Keep the deed.`)
    // Past markup, the sentence is quoted whole with its markup as written.
    const marked = LONG.replace("please leave", "please *leave*")
    writeLeanText(directory, "request.md", `${REQUEST} Prepare the release letter. ${marked} Keep the deed.`)
    expect(hint(refusal(directory, WORDS))).toBe(`In the lawyer's text, the whole sentence containing them, with its markup as written, is "${marked}". `)
    accepted(directory, marked)
  })

  test("of several places, when none fits the excerpt length, the first short enough is quoted whole and the rest counted", () => {
    const directory = matter(SOURCE_TEXT)
    const other = LONG.replace("The approval", "The consent")
    writeLeanText(directory, "request.md", `${REQUEST}\n${LONG}\n${other}`)
    let issue = refusal(directory, WORDS)
    expect(hint(issue)).toBe(`In the lawyer's text, 2 places contain them; the whole sentences there are "${LONG}" and 1 more not quoted here. `)
    for (const whole of offered(issue)) accepted(directory, whole)
    // A first place too long to quote is counted; the next is quoted whole.
    const huge = sized(2001)
    writeLeanText(directory, "request.md", `${REQUEST}\n${huge}\n${LONG}`)
    issue = refusal(directory, WORDS)
    expect(hint(issue)).toBe(`In the lawyer's text, 2 places contain them; the whole sentences there are "${LONG}" and 1 more not quoted here. `)
    for (const whole of offered(issue)) accepted(directory, whole)
    // When each is too long to quote, none is cut.
    writeLeanText(directory, "request.md", `${REQUEST}\n${huge}\n${huge.replace("The committee", "Our committee")}`)
    issue = refusal(directory, WORDS)
    expect(hint(issue)).toBe("In the lawyer's text, 2 places contain them; the whole sentences there are each too long to quote here: " +
      "find them where those words stand in the request, the lawyer's updates or answers. ")
    expect(offered(issue)).toEqual([])
  })

  test("a sentence too long to quote is pointed to, never cut, and copied from the lawyer's text it is accepted", () => {
    const directory = matter(SOURCE_TEXT)
    for (const length of [EXCERPT - 1, EXCERPT, EXCERPT + 1, 1000, 2000]) {
      const whole = sized(length)
      writeLeanText(directory, "request.md", `${REQUEST} Prepare the release letter. ${whole} Keep the deed.`)
      expect([whole.length, hint(refusal(directory, WORDS))]).toEqual([length, `In the lawyer's text, the whole sentence containing them is "${whole}". `])
      accepted(directory, whole)
    }
    for (const length of [2001, 6000]) {
      const whole = sized(length)
      writeLeanText(directory, "request.md", `${REQUEST} Prepare the release letter. ${whole} Keep the deed.`)
      const issue = refusal(directory, WORDS)
      expect(hint(issue)).toBe(`In the lawyer's text, the whole sentence containing them is too long to quote here: ${FIND}`)
      expect([offered(issue), issue.includes("…")]).toEqual([[], false])
      accepted(directory, whole)
    }
    // Words over two sentences that together are too long are pointed to the same way.
    const whole = sized(2001)
    writeLeanText(directory, "request.md", `${REQUEST} Prepare the release letter. ${whole} Keep the deed.`)
    expect(hint(refusal(directory, "out of this package. Keep"))).toBe("In the lawyer's text, the whole sentences containing them are too long to quote here: " +
      "find them where those words stand in the request, the lawyer's updates or answers. ")
  })
})

describe("sentence bounds: line labels in sequence", () => {
  // A lettered or roman label at the start of a line that comes next after a
  // label of an open list (b. after a., ii. after i., v. after iv.), or a list's
  // first label (a. i.) on a line that starts a sentence or follows a sentence
  // end, opens a list item before a word of either case, and the bound falls
  // before the label. A paragraph that opens with no list marker closes every list.
  const documents = [planned("d1", "Release letter", ["Kestrel Works"])]
  const LEAVE = "Leave the guarantee out of this package."
  const LEFT = "the guarantee is to be left out of this package."
  const SIGNS = "the guarantor signs the guarantee for Kestrel Works"
  // A paragraph of its own that directs the act, for texts whose list does not.
  const NOTE = `Note that ${SIGNS}.`
  // A sentence whose hard wrap puts "v." at the start of a line.
  const WHOLE = "We will not, following R v. the Secretary of State, leave the guarantee out of this package."
  const signing = (fields: Partial<CoverageAct> = {}) => act({ act: "sign the guarantee", quote: SIGNS, parties: [], document: null, ...fields })
  const checked = (directory: string, words: string, final = false) =>
    parseCoverage(JSON.stringify({ acts: [signing({ excluded_by: words })] }), directory, documents, { final })
  const accepted = (directory: string, ...sentences: string[]) => {
    for (const words of sentences) expect([words, checked(directory, words)]).toEqual([words, { issues: [], value: [signing({ excluded_by: words })] }])
  }
  /** Refused as not whole sentences of the lawyer's text, and for nothing else. */
  const refused = (directory: string, ...fragments: string[]) => {
    for (const words of fragments) {
      const { issues } = checked(directory, words)
      expect([words, issues.length, issues[0]]).toEqual([words, 1, expect.stringContaining("those words are not one or more whole sentences")])
    }
  }
  /** Kept on the final attempt, with the act's own item overlapping its quote. */
  const kept = (directory: string, words: string) => {
    expect([words, checked(directory, words, true)]).toEqual([words, { issues: [], value: [signing({ excluded_by: words })] }])
    const own = checked(directory, `${SIGNS}.`).issues
    expect([own.length, own[0]]).toEqual([1, expect.stringContaining(`which overlaps the words that direct this act ("${SIGNS}")`)])
  }
  const request = (...lines: string[]) => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", lines.join("\n"))
    return directory
  }

  test("the next label opens an item before a lowercase word, so the sentence above leaves the act's document out and the act's own item overlaps its quote", () => {
    const roman = request("The papers:", "i. Keep the deed.", "ii. Keep the letter.", "iii. Keep the minutes.", `iv. ${LEAVE}`, `v. ${SIGNS}.`)
    for (const directory of [request("The papers:", `a. ${LEAVE}`, `b. ${SIGNS}.`), request("The papers:", `i. ${LEAVE}`, `ii. ${SIGNS}.`), roman]) {
      // The final attempt keeps the exclusion instead of dropping it.
      kept(directory, LEAVE)
      accepted(directory, LEAVE, LEAVE.slice(0, -1))
    }
    accepted(roman, "Keep the deed.", "Keep the letter.", "Keep the minutes.")
  })

  test("the bound before the next line's label leaves the label out of the sentence above, as accepted and as the hint quotes it", () => {
    const directory = request("The papers:", `i. ${LEAVE}`, "ii. Keep the deed.", "", NOTE)
    accepted(directory, LEAVE, "Keep the deed.")
    const { issues } = checked(directory, "Leave the guarantee out")
    expect([issues.length, issues[0]]).toEqual([1, expect.stringContaining(`the whole sentence containing them is "${LEAVE}". Copy verbatim`)])
  })

  test("guards: a label opens nothing mid-line, out of sequence, before a number, or after a first label read as a wrap; (b) and b) open as before", () => {
    // Mid-line labels and "e.g." carry the sentence on.
    const inline = request(`The papers: a. ${LEAVE} b. ${SIGNS}.`, "a. Keep the deed.", "b. Leave out the papers, e.g. the deed of release.")
    refused(inline, LEAVE, "Leave out the papers, e.g.")
    accepted(inline, "Leave out the papers, e.g. the deed of release.")
    // A wrap that starts with a letter other than the next one, or with the next before a number, carries the item on.
    const wraps = request("The papers:", "a. Leave out the settlement in Moss", "v. the Crown for now.", "b. Leave out the notice under", "s. 146 of the Act for now.",
      "c. Leave out the deed signed", "d. 1990 by the Seller.", "", NOTE)
    accepted(wraps, "Leave out the settlement in Moss v. the Crown for now.", "Leave out the notice under s. 146 of the Act for now.", "Leave out the deed signed d. 1990 by the Seller.")
    refused(wraps, "Leave out the settlement in Moss", "Leave out the notice under", "Leave out the deed signed", "1990 by the Seller.")
    const roman = request("The papers:", "i. Leave out the settlement in Moss", "v. the Crown for now.", "", NOTE)
    accepted(roman, "Leave out the settlement in Moss v. the Crown for now.")
    refused(roman, "Leave out the settlement in Moss")
    // A first label before a lowercase word opens an item on a line that starts a sentence, so the label after it follows it.
    const opening = request("Out of scope:", "a. leave the deed out.", "b. the minutes stay.", "", NOTE)
    accepted(opening, "leave the deed out.", "the minutes stay.", "a. leave the deed out. b. the minutes stay.")
    // On a line that carries on the one above, it is read as a wrap, so the label after it follows none.
    const wrapped = request("Leave out the papers listed in", "a. the schedule and", "b. the annex for now.", "", NOTE)
    accepted(wrapped, "Leave out the papers listed in a. the schedule and b. the annex for now.")
    refused(wrapped, "Leave out the papers listed in", "the schedule and", "the annex for now.")
    // Bracketed labels already open an item before any word.
    for (const [first, second] of [["(a)", "(b)"], ["a)", "b)"], ["(i)", "(ii)"]])
      accepted(request("The papers:", `${first} ${LEAVE}`, `${second} ${SIGNS}.`), LEAVE)
  })

  test("a list's first label opens an item before a lowercase word after a colon, a sentence end or at the start, so the next label follows it", () => {
    for (const [first, second] of [["a.", "b."], ["i.", "ii."]]) {
      for (const directory of [request("Notes:", `${first} ${LEFT}`, `${second} ${SIGNS}.`), request(`${first} ${LEFT}`, `${second} ${SIGNS}.`)]) {
        kept(directory, LEFT)
        accepted(directory, LEFT, `${first} ${LEFT}`)
      }
      // After a sentence that ends before the label's lowercase word, both sentences stand alone.
      const after = request("Prepare the release.", `${first} ${LEFT}`, `${second} ${SIGNS}.`)
      kept(after, LEFT)
      accepted(after, "Prepare the release.", LEFT)
    }
  })

  test("the hint quotes a lowercase item's sentence without its label or the next item, and given back it is accepted", () => {
    const directory = request("Notes:", `a. ${LEFT}`, `b. ${SIGNS}.`)
    const { issues } = checked(directory, "the guarantee is to be left out")
    expect([issues.length, issues[0]]).toEqual([1, expect.stringContaining(`the whole sentence containing them is "${LEFT}". Copy verbatim`)])
    accepted(directory, LEFT)
  })

  test("an outer label follows its list across a nested list, which it closes, so a wrap in the outer item carries on", () => {
    const nested = request("Instructions:", "a. Prepare the release.", "   i. Kestrel Works signs it.", `   ii. ${LEAVE}`, `b. ${SIGNS}.`)
    const romans = request("Instructions:", "i. Prepare the release.", "   a. Kestrel Works signs it.", `   b. ${LEAVE}`, `ii. ${SIGNS}.`)
    for (const directory of [nested, romans]) {
      kept(directory, LEAVE)
      accepted(directory, LEAVE, "Prepare the release.", "Kestrel Works signs it.")
    }
    // Guard: "v." after a closed i.–iv. list is no item of it.
    const closed = request("Instructions:", "a. Prepare the release:", "   i. the deed", "   ii. the note", "   iii. the minutes", "   iv. the side letter",
      "b. We will not, following R", "v. the Secretary of State, leave the guarantee out of this package.", "", NOTE)
    accepted(closed, WHOLE, "the side letter")
    refused(closed, "We will not, following R", "the Secretary of State, leave the guarantee out of this package.")
  })

  test("a paragraph that opens with no list marker closes the list, so a wrap after it follows no label; a blank line between items does not", () => {
    const directory = request("Papers:", "i. The deed.", "ii. The note.", "iii. The release.", "iv. The side letter.", "", "We will not, following R",
      "v. the Secretary of State, leave the guarantee out of this package.", "", NOTE)
    accepted(directory, WHOLE)
    refused(directory, "We will not, following R", "the Secretary of State, leave the guarantee out of this package.")
    // The hint offers the whole sentence, never the words after the wrap.
    const { issues } = checked(directory, "leave the guarantee out of this package")
    expect([issues.length, issues[0]]).toEqual([1, expect.stringContaining(`the whole sentence containing them is "${WHOLE}". Copy verbatim`)])
    const split = request("The papers:", `i. ${LEAVE}`, "", `ii. ${SIGNS}.`)
    kept(split, LEAVE)
    accepted(split, LEAVE)
  })
})
