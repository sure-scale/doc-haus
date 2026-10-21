import type { Client } from "../../api/opencode/client"
import type { DraftScope } from "../../api/opencode/routing"
import { readBrowserValue, writeBrowserValue } from "../../browser-storage"
import { QUICK_DRAFT } from "../../agents"
import type { Turn } from "./turn-model"

// The quick-draft offer (#282): when Auto routes a short drafting request that
// a quick draft can serve, the lawyer chooses in a native session form. The
// form only records the choice; this app dispatches it once the reply settles.

type FormInput = Parameters<Client["session"]["form"]["create"]>[0]
type PendingForm = Awaited<ReturnType<Client["session"]["form"]["list"]>>[number]
export type DraftChoice = "quick" | "full"

/** Form metadata kind for a route-only choice: it carries the request, not a
 * question for a model. */
export const ROUTE_FORM = "route"
const FIELD = "route"

export const QUICK_BANNER = "Quick draft — full drafting checks not run."

/** The route form for one request, with the lawyer's last choice on this
 * matter (else Quick) pre-selected. */
export function routeForm(sessionID: string, request: string, preselect: DraftChoice): FormInput {
  return {
    sessionID,
    title: "Quick draft or full drafting run?",
    metadata: { kind: ROUTE_FORM, route: "draft", request },
    fields: [
      {
        key: FIELD,
        type: "string",
        required: true,
        title: "How should doc.haus draft this?",
        default: preselect,
        options: [
          {
            value: "quick",
            label: "Quick draft here in chat",
            description: "From a few matter sources, quotes verified. No review stages, checklist or memo.",
          },
          {
            value: "full",
            label: "Full drafting run",
            description: "Plan, draft, review and memo. Usually 10–20 minutes.",
          },
        ],
      },
    ],
  }
}

/** The request a pending route form carries, or undefined for any other form. */
export function routeRequest(form: Pick<PendingForm, "metadata">) {
  const metadata = form.metadata
  return metadata?.kind === ROUTE_FORM && metadata.route === "draft" && typeof metadata.request === "string"
    ? metadata.request
    : undefined
}

/** The choice in a route form's answer; anything else is no choice. */
export function routeChoice(answer: Record<string, unknown>): DraftChoice | undefined {
  const value = answer[FIELD]
  return value === "quick" || value === "full" ? value : undefined
}

const choiceKey = (matterId: string) => `dh.quickDraft.choice.${matterId}`

/** The lawyer's last choice on this matter: a pre-selection only, never consent. */
export function rememberedChoice(matterId: string): DraftChoice {
  return readBrowserValue(choiceKey(matterId)) === "full" ? "full" : "quick"
}

export function rememberChoice(matterId: string, choice: DraftChoice) {
  writeBrowserValue(choiceKey(matterId), choice)
}

/** Why a quick draft cannot cover this request, in the lawyer's terms, or
 * undefined when it can. A missing scope is a limit too. */
export function quickDraftLimit(scope: DraftScope | undefined) {
  if (!scope) return "doc.haus could not confirm this request fits a quick draft."
  if (scope.documents !== "one") return "This request asks for more than one document."
  if (scope.legal_effect !== "no") return "This document would have legal effect once sent, signed or filed."
  if (scope.outside_law !== "no") return "This document needs law beyond the matter's documents."
  if (scope.firm_material !== "no") return "This document should follow a firm template or playbook."
  if (scope.sources !== "few") return "This document draws on many of the matter's documents."
}

/** The router judges "sources" from the message alone; "few" is up to three
 * of the matter's documents (dochaus/agent/router.md). A matter with at most
 * three source documents can only need few, so its count decides. */
export function withMatterSources(scope: DraftScope | undefined, sources: number | undefined): DraftScope | undefined {
  return scope && sources !== undefined && sources <= 3 ? { ...scope, sources: "few" } : scope
}

// The quick-draft answer contract (dochaus/agent/quick-draft.md): a "Form:"
// line, the draft, then "Sources" and "Open items".
const FORM_LINE = /^[ \t]*[*_]*form[ \t]*:[ \t]*[*_]*[ \t]*(advice letter|advice|letter|agreement|other)[ \t]*[*_]*[ \t]*\.?[ \t]*$/im
const SOURCES_LINE = /^[ \t]*(?:#{1,6}[ \t]+)?[*_]*sources[*_]*[ \t]*(?::|$)/im
// Square brackets in a draft mark something still to fill or confirm; a
// citation ("[<Document> § <section>]") and a Markdown link are not.
const BRACKETED = /\[([^[\]\n]+)\](?![([:])/g

/** A quick answer read by its contract: its form (which sets the firm's
 * legend, applied by code), its text without the form line, and the bracketed
 * placeholders left in the draft before its sources, each an open item. */
export function quickDraftAnswer(answer: string) {
  const match = FORM_LINE.exec(answer)
  const text = match
    ? `${answer.slice(0, match.index)}${answer.slice(match.index + match[0].length).replace(/^\s*\n/, "")}`.trim()
    : answer.trim()
  const draft = text.slice(0, SOURCES_LINE.exec(text)?.index ?? text.length)
  const placeholders = [...draft.matchAll(BRACKETED)]
    .filter((bracket) => /[\p{L}_●•]/u.test(bracket[1]!) && !bracket[1]!.includes("§"))
    .map((bracket) => bracket[0])
  return { form: match?.[1]?.toLowerCase(), text, placeholders: [...new Set(placeholders)] }
}

/** What a full run started from the quick draft at `index` carries: the request
 * that began the quick exchange, with the lawyer's later messages in it, as the
 * request; and the latest quick text as an unverified working draft, never as
 * request text. */
export function upgradeInput(turns: Turn[], index: number) {
  const quick = turns[index]
  if (!quick || quick.role !== "assistant" || quick.agent !== QUICK_DRAFT) return undefined
  // The contiguous quick-draft exchange that ends at this answer.
  let start = index
  while (start > 0 && (turns[start - 1].role === "user" || turns[start - 1].agent === QUICK_DRAFT)) start -= 1
  const asked = turns.slice(start, index).filter((turn) => turn.role === "user" && turn.text.trim())
  if (!asked.length || !quick.text.trim()) return undefined
  const [first, ...later] = asked.map((turn) => turn.text.trim())
  const request = later.length
    ? `${first}\n\nLater messages from the lawyer, in order:\n${later.map((text) => `- ${text}`).join("\n")}`
    : first
  // A quick answer that declined the request carries no draft. The form line
  // is not draft text: the full run plans its own form.
  const text = quickDraftAnswer(quick.text).text
  return { request, ...(DECLINED.test(text) ? {} : { workingDraft: text }) }
}

// The first line quick-draft.md gives when a request is outside its scope.
const DECLINED = /^quick draft can['’]t cover this request/i
