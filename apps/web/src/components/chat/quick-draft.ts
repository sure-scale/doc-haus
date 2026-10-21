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
  // A quick answer that declined the request carries no draft.
  const text = quick.text.trim()
  return { request, ...(DECLINED.test(text) ? {} : { workingDraft: text }) }
}

// The first line quick-draft.md gives when a request is outside its scope.
const DECLINED = /^quick draft can['’]t cover this request/i
