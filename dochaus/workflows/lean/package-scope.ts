import { normalize, SHORT_QUOTE } from "../../lib/quote-verify"
import { shortHash, type Question, type StageID } from "./state"

// The one question on the package's scope: whether the package carries an
// instruction, and in which document, or whether it is outside this request.
// The plan step asks it for instructed acts its plan leaves without an
// instrument their parties execute (draft-sections.ts); the gate asks it for
// instructions the checklist check left open (gate.ts). Each takes its id from
// scopeAnswers, under one prefix, keyed on what it asks about: the gate on the
// instructions' verbatim quotes, the plan step on each act's quote with the
// parties who must take it, so another party's act in the same words is another
// question. An answered question settles only the instructions it showed
// (scopeQuestionShows), read from the instruction quotes the asker wrote at
// fixed places in its text, never from other quoted words, and across askers
// act by act, for the parties who must take each (ScopeParties), or, by its
// id, asked about (scopeAnswers): neither asker asks about those again, and
// neither asks a question an answered one already is. The gate's rows name no
// parties, so the plan step's act lines settle a gate row only when they
// showed every act the plan's coverage record has its quotes directing that
// the plan step asks about (gateScopeParties), each for its own parties, and a
// gate quote settles the plan step's act only when every act its row, read
// whole, directs is for that act's parties.
// Each asks once more about instructions no answered question settled, and no
// more once two of its own are answered. While one asker's question is open
// the other asks none: the open one stays the question to answer, and the run
// discloses what is left. It imports only the run state, so both can import it
// without an import cycle.

/** The id prefix of the one question on the package's scope. */
export const PACKAGE_SCOPE_QUESTION = "package-scope-"

/** A question's id keyed on what it asks about: `prefix` and a short hash of
 * the unique `keys`, sorted, never of item ids or order. The gate's checklist
 * conflict question uses it too. */
export function keyedQuestionID(prefix: string, keys: string[]) {
  return `${prefix}${shortHash(JSON.stringify([...new Set(keys)].toSorted()))}`
}

/** A package-scope question's id: the prefix and the keys of what it asks about. */
export function packageScopeQuestionID(keys: string[]) {
  return keyedQuestionID(PACKAGE_SCOPE_QUESTION, keys)
}

/** The keys of verbatim quotes: each normalized, so spacing and quotation
 * marks do not change the id. The gate's keys, and the default. */
export const quoteKeys = (quotes: string[]) => quotes.map(normalize)

export function answeredScopeQuestion(question: Question) {
  return question.status === "answered" && question.id.startsWith(PACKAGE_SCOPE_QUESTION)
}

/** How many package-scope questions each asker (the plan step, the gate) asks
 * in a run: one, then one more. The coverage job may quote an answer as a new
 * act, and no question shows an answer's words, so without this bound each
 * answer could bring another question. */
export const SCOPE_QUESTIONS_PER_ASKER = 2

/** Which of the `answered` package-scope questions settle each of `rows` (acts
 * or instructions, each with its verbatim quotes), the rows the asker `stage`
 * asks about, and the `id` it asks them under. A row is settled by each one
 * that shows it: by default one that shows one of its quotes
 * (scopeQuestionShows); a caller that knows more about a row passes its own
 * `shows`. The rows none shows would be asked in one question whose id hashes
 * their `keys`: by default their quotes (quoteKeys); a caller whose rows the
 * quotes alone do not tell apart passes its own. An answered question with that
 * id asked about those very rows, whatever its text shows, so it settles them
 * and none is asked (asked again, it would be dropped as answered, leaving no
 * open question to answer). Nor is one asked once SCOPE_QUESTIONS_PER_ASKER of
 * the stage's own package-scope questions are answered. Rows left unasked are
 * disclosed by the caller; the asker uses `id`, so it never asks an answered id. */
export function scopeAnswers<T>(rows: T[], quotes: (row: T) => string[], answered: Question[], stage: StageID,
  shows = (question: Question, row: T) => scopeQuestionShows(question, quotes(row)),
  keys = (row: T) => quoteKeys(quotes(row))) {
  const settled = new Map(rows.map(row => [row, answered.filter(question => shows(question, row))]))
  const unshown = rows.filter(row => !settled.get(row)!.length)
  const id = packageScopeQuestionID(unshown.flatMap(keys))
  const same = unshown.length ? answered.find(question => question.id === id) : undefined
  if (same) for (const row of unshown) settled.set(row, [same])
  const spent = answered.filter(question => question.stage === stage).length >= SCOPE_QUESTIONS_PER_ASKER
  return { by: (row: T) => settled.get(row) ?? [], ask: same || spent ? [] : unshown, id }
}

/** The open package-scope question another stage asked, if any. While it is
 * open no stage asks another; a stage may ask its own again, since a stage's
 * questions replace the ones it asked before. */
export function openScopeQuestion(questions: Question[], stage: StageID) {
  return questions.find(question => question.status === "open" && question.id.startsWith(PACKAGE_SCOPE_QUESTION) && question.stage !== stage)
}

/** Whether a package-scope question showed an instruction. Only the
 * instruction quotes its asker wrote count, read at the fixed words around
 * them (questionQuotes): a party name, act description, document title, check
 * or open decision the question quotes is not one, even when it quotes an
 * instruction. After the usual normalization of quote marks and whitespace,
 * and in any case, a quote is shown by an instruction quote equal to it, one
 * the asker cut to an excerpt (ending in an ellipsis) that it starts with, one
 * that holds it or one it holds (a re-run may quote the whole sentence where
 * the question quoted a clause, or the reverse). A quote and an instruction
 * quote that differ must both be at least SHORT_QUOTE long to match by
 * containment or excerpt, so a short phrase does not show an instruction. A
 * caller that knows who must take its row passes `parties`, the acts the row
 * stands for: the plan step's act itself, or the acts a gate row's quotes
 * direct that the plan step asks about (gateScopeParties), none when unknown.
 * The plan step's act lines, whichever asker's the row is, then show it only
 * when one shows its quote and the question showed each of its acts: an act
 * line named that act's parties, in either reading of the names it wrote
 * (partyReadings), and its quote shows the act's own quote, never only a
 * shorter quote the act's holds (lineShows). So
 * another party's act in the same words is still asked, as is an act quoting
 * the whole sentence when the line quoted one clause of it; an instruction
 * directing two parties' acts is settled by a question that listed both; an
 * act whose parties no one names is matched by its own quote, never by another
 * such act's; and a gate row directing no act is shown by no act line. A gate
 * row's quote shows another gate row as before. It shows the plan step's act
 * only when it is the act's quote or holds it, never when the act's longer
 * quote holds it, and only when the gate row it is read from, judged whole as
 * the gate judges its own rows, directs acts of the act's parties alone: every
 * act any reading of the row directs (`directed`) is for them, and the gate
 * cut none of its quotes to an excerpt, since the words cut away may direct
 * another party's act. A row is read at each ", " it holds, so a quote holding
 * a quoted term or several quotes is judged whole whichever part of it holds
 * the act's quote. */
export function scopeQuestionShows(question: Question, quotes: string[], parties?: ScopeParties) {
  const shown = questionQuotes(question)
  const everyAct = !!parties?.acts.length && parties.acts.every(act => lineShows(shown, act, parties.key))
  return quotes.map(quote => normalize(quote).toLowerCase()).some(quote => quote && shown.some(field => {
    if (!parties || (!field.parties && !parties.act)) return holds(field.quote, quote, true)
    if (!holds(field.quote, quote, !!field.parties)) return false
    if (field.parties) return everyAct
    const row = field.row!
    return !row.some(reading => reading.endsWith("…")) && sameParties(parties.directed(row), parties)
  }))
}

/** Who must take a row, for scopeQuestionShows: `acts`, the acts it stands
 * for, each with its quote and parties: the plan step's act itself, or the
 * acts a gate row's quotes direct that the plan step asks about
 * (gateScopeParties), none when unknown; `key`, the key of a list of parties,
 * as an act gives them or an act line names them; `directed`, the acts the
 * quotes direct; `act`, whether the row is the plan step's act. */
export type ScopeParties = { acts: ScopeAct[]; key: (named: string[]) => string; directed: (quotes: string[]) => ScopeAct[]; act: boolean }

type ScopeAct = { quote: string; parties: string[] }

/** Whether one of the act lines among `shown` showed `act`: it named the
 * act's parties, in either reading (partyReadings), and its quote is the
 * act's, an excerpt the act's starts with, or holds it. A line whose shorter
 * quote the act's merely holds showed another act, perhaps one clause of the
 * act's sentence, not this one. */
function lineShows(shown: ShownQuote[], act: ScopeAct, key: (named: string[]) => string) {
  const quote = normalize(act.quote).toLowerCase()
  return !!quote && shown.some(field => field.parties?.some(named => key(named) === key(act.parties)) && holds(field.quote, quote, false))
}

/** Whether there are `acts` and each is for the parties of the row's act. */
const sameParties = (acts: ScopeAct[], { acts: row, key }: ScopeParties) =>
  acts.length > 0 && new Set([...acts, ...row].map(act => key(act.parties))).size === 1

/** An instruction quote a question shows: an act line's with the readings of
 * the parties it named, a gate row's with every reading of that row. */
type ShownQuote = { quote: string; parties?: string[][]; row?: string[] }

// The plan step's act line (package-coverage.ts actLine): `"<act>", directed by
// "<quote>" [source: <id>]; to be taken by <parties>; <where the plan has it>.`
// The quote ends at the first `" [source: `, and the source id at the first
// "]; to be taken by " followed by the parties as actLine writes them, quoted
// or UNNAMED, so an id holding "]", or those words, is read whole.
const UNNAMED = "parties the instructions and sources do not name"
const ACT_LINE = new RegExp(`, directed by "(.*?)" \\[source: .*?\\]; to be taken by (".*?"|${UNNAMED}); (?:no planned document carries it|planned in )`, "g")
// The gate's row (gate.ts packageScopeQuestion), one per line: `<check>
// Sources: <document> "<quote>", ...`, each quote longer than GATE_EXCERPT cut
// to an excerpt that long, and then its open decisions or the lack of one.
// Rows it wrote earlier, run together and each led by its item id, read alike.
const SOURCES = / Sources: (.*?)" (?:Open decisions: |No open decision recorded)/g
export const GATE_EXCERPT = 160
// Either line is read from the words that open it, `, directed by "` or `
// Sources: `, which its asker writes once per line and never in a field
// written from a model's words (scopeField). So each line read starts where
// the asker wrote one, and reads up to the asker's own words after its quote
// or sources: the quotes read are the ones the asker wrote, never a field's
// words. Instruction quotes and matter document names are written as they
// are, so one holding those words is still read at them; a party name holding
// the words after the parties cuts that act line's parties there.

/** A field a package-scope question writes from a model's words (an act's
 * description, the act line's party list and where the plan has it, a check,
 * an open decision), with the space dropped after the words that open an act
 * line, before a quotation mark or the field's end, or a gate row, so it
 * opens neither (the questions' own words around the field hold neither). A
 * field holding neither is written as it is. */
export const scopeField = (text: string) => text.replace(/(,\s+directed\s+by)\s+(?=["“”]|$)|(Sources:)\s+/g, "$1$2")

const parsed = new WeakMap<Question, { text: string; quotes: ShownQuote[] }>()

/** The instruction quotes a package-scope question shows, lower-cased, each
 * act line's with the readings of the parties it named (partyReadings), each
 * gate row's with the row's readings: all its quotes as sourceQuotes reads
 * them, and its sources as written, so a quote whose start is not found still
 * counts. The lazy match up to the fixed words after a quote reads it whole
 * whatever quotation marks it holds or its excerpt cut. Each question's text is
 * parsed once. */
function questionQuotes(question: Question) {
  const saved = parsed.get(question)
  if (saved?.text === question.question) return saved.quotes
  const text = normalize(question.question)
  const quotes: ShownQuote[] = [
    ...[...text.matchAll(ACT_LINE)].map(([, quote, named]) => ({ quote: quote!.toLowerCase(), parties: partyReadings(named!) })),
    ...[...text.matchAll(SOURCES)].flatMap(([, sources]) => {
      const read = sourceQuotes(sources!).map(quote => quote.toLowerCase())
      const row = [...read, sources!.toLowerCase()]
      return read.map(quote => ({ quote, row }))
    }),
  ]
  parsed.set(question, { text: question.question, quotes })
  return quotes
}

/** The names quotedList wrote for an act line's parties (`"A"`, `"A" and
 * "B"` or `"A", "B" and "C"`). A name may itself hold `", "` or `" and "`, so
 * they are read two ways: split at each, and split only where the name before
 * holds an even number of quotation marks, its defined terms paired, so `Acme
 * Ltd (the "Seller", "we")` stays one name. The second reading is given only
 * when it differs; UNNAMED names none. */
function partyReadings(named: string) {
  if (named === UNNAMED) return [[]]
  const [first = "", ...rest] = named.slice(1, -1).split(/(", "|" and ")/)
  const every = [first], paired = [first]
  for (let index = 0; index < rest.length; index += 2) {
    every.push(rest[index + 1]!)
    if ((paired.at(-1)!.match(/"/g)?.length ?? 0) % 2) paired[paired.length - 1] += rest[index]! + rest[index + 1]!
    else paired.push(rest[index + 1]!)
  }
  return paired.length < every.length ? [every, paired] : [every]
}

/** The quotes of `<document> "<quote>"` sources joined by ", ", the last
 * without its closing mark. A quote runs from the mark after a document name
 * to a following ", " or the end; one holding ", " of its own is read at each,
 * so its whole text is among them. A quote the gate cut to an excerpt is also
 * read back from its end, since it is GATE_EXCERPT long, so quotation marks in
 * its document's name do not move its start. */
function sourceQuotes(sources: string) {
  const starts = [...sources.matchAll(/(?:^|", )[^"]*? "/g)].map(match => match.index + match[0].length)
  const ends = [...[...sources.matchAll(/", /g)].map(match => match.index), sources.length]
  const cut = ends.map(end => end - GATE_EXCERPT).filter(start => start >= 2 && sources.slice(start - 2, start) === ' "' && sources[start + GATE_EXCERPT - 1] === "…")
  return [...starts.flatMap(start => ends.filter(end => end > start).map(end => sources.slice(start, end))),
    ...cut.map(start => sources.slice(start, start + GATE_EXCERPT))]
}

/** Whether an instruction quote a question showed shows `quote`: equal to
 * it, cut to an excerpt that `quote` starts with, holding `quote`, or, when a
 * `longer` quote may be the same instruction, held by it. */
function holds(shown: string, quote: string, longer: boolean) {
  const cut = shown.endsWith("…") ? shown.slice(0, -1) : undefined
  const text = cut ?? shown
  return shown === quote || (cut !== undefined && cut.length >= SHORT_QUOTE && quote.startsWith(cut)) ||
    (quote.length >= SHORT_QUOTE && text.includes(quote)) || (longer && text.length >= SHORT_QUOTE && quote.includes(text))
}
