import type { DraftingPreferences } from "./drafting-preferences"

// Page furniture for documents doc.haus renders: page numbers, the draft line
// and the firm's chosen legend. The planner picks only each document's form
// from a closed set; what the form carries follows here from the firm's
// settings. Legends have legal effect, so they are the firm's own text and are
// never written by a model. Uploaded, counterparty and template files are
// never given furniture, and a document that already has a header or footer
// keeps it.

/** A planned document's form, from its outline "Form:" line. */
export type DocumentForm = "agreement" | "advice" | "letter" | "other"
export const DOCUMENT_FORMS: readonly DocumentForm[] = ["agreement", "advice", "letter", "other"]

/** What a rendered file is: a planned document's form, or the run's internal drafting notes. */
export type FurnitureKind = DocumentForm | "notes"

/** `header` lines are written in order at the top of every page; `pageNumber`
 * goes in the footer; `firstPage: "blank"` leaves the first page without either
 * (a letter is numbered from its second page). */
export type PageFurniture = {
  header: string[]
  pageNumber?: { format: "plain" | "page"; align: "center" | "right" }
  firstPage?: "blank"
}

export const NOTES_MARKING = "Internal drafting notes — not for circulation"

type FurniturePreferences = Pick<DraftingPreferences, "firm" | "dateFormat" | "pageNumbers" | "pageNumberPosition" | "draftLine" | "agreementLegend" | "adviceLegend" | "notesMarking">

/** The furniture a kind of file carries under the firm's settings, or
 * undefined when it carries none. `draftedAt` is the time the draft was made
 * (an ISO timestamp kept in run state, never the render clock); without it
 * there is no draft line. */
export function pageFurniture(kind: FurnitureKind, prefs: FurniturePreferences, draftedAt?: string): PageFurniture | undefined {
  const legend = kind === "agreement" ? prefs.agreementLegend : kind === "advice" ? prefs.adviceLegend : ""
  const date = draftedAt && prefs.draftLine && (kind === "agreement" || kind === "advice") ? formatDraftDate(draftedAt, prefs.dateFormat) : undefined
  const header = [
    ...(legend ? [legend] : []),
    ...(date ? [`${prefs.firm.trim() ? `${prefs.firm.trim()} draft` : "Draft"}: ${date}`] : []),
    ...(kind === "notes" && prefs.notesMarking ? [NOTES_MARKING] : []),
  ]
  const pageNumber = prefs.pageNumbers === "off" ? undefined : { format: prefs.pageNumbers, align: prefs.pageNumberPosition }
  if (!header.length && !pageNumber) return undefined
  return { header, ...(pageNumber ? { pageNumber } : {}), ...(kind === "letter" ? { firstPage: "blank" as const } : {}) }
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

/** The draft date in the firm's date style, read in the host's local time. */
export function formatDraftDate(at: string, style: DraftingPreferences["dateFormat"]) {
  const date = new Date(at)
  if (Number.isNaN(date.getTime())) return undefined
  const [year, month, day] = [date.getFullYear(), date.getMonth(), date.getDate()]
  if (style === "iso") return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`
  if (style === "day-month-year") return `${day} ${MONTHS[month]} ${year}`
  return `${MONTHS[month]} ${day}, ${year}`
}
