import type { DraftingPreferences } from "../../api/ingest/preferences"

// Page layout choices shared by Settings → Drafting and onboarding. The engine
// renders the furniture itself (dochaus/lib/page-furniture.ts); this mirrors its
// text for the preview only, so the web bundle never imports the renderer.

/** Mirrors the server's LEGEND_MAX_LENGTH (dochaus/lib/drafting-preferences.ts). */
export const LEGEND_MAX_LENGTH = 120

// Legends have legal effect, so these are offered, never applied by default.
// "Without prejudice" is deliberately absent: it is a per-letter legal call.
export const AGREEMENT_LEGENDS = ["Subject to contract", "Draft — for discussion purposes only"] as const
export const ADVICE_LEGENDS = ["Privileged & confidential", "Privileged & confidential — attorney-client communication"] as const

export const NOTES_MARKING = "Internal drafting notes — not for circulation"

// General practice by jurisdiction family, keyed on the pack code before any
// "-" (US-NY → US, AU-NSW → AU). Only a visible suggestion: nothing is chosen.
const US = { agreement: AGREEMENT_LEGENDS[1], advice: ADVICE_LEGENDS[1] }
const COMMONWEALTH = { agreement: AGREEMENT_LEGENDS[0], advice: ADVICE_LEGENDS[0] }
const SUGGESTIONS: Record<string, { agreement: string; advice: string }> = {
  US,
  EW: COMMONWEALTH, SCT: COMMONWEALTH, NIR: COMMONWEALTH, IE: COMMONWEALTH,
  AU: COMMONWEALTH, NZ: COMMONWEALTH, SG: COMMONWEALTH, HK: COMMONWEALTH,
}

/** The presets suggested for a set of jurisdiction pack codes; none for others. */
export function suggestedLegends(codes: string[]) {
  const found = codes.flatMap(code => SUGGESTIONS[code.split("-")[0]!.toUpperCase()] ?? [])
  return {
    agreement: [...new Set(found.map(item => item.agreement))],
    advice: [...new Set(found.map(item => item.advice))],
  }
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

/** A date in the firm's date style, as the draft line prints it. */
export function formatDraftDate(date: Date, style: DraftingPreferences["dateFormat"]) {
  const [year, month, day] = [date.getFullYear(), date.getMonth(), date.getDate()]
  if (style === "iso") return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`
  if (style === "day-month-year") return `${day} ${MONTHS[month]} ${year}`
  return `${MONTHS[month]} ${day}, ${year}`
}

/** "Doe & Partners LLP draft: 20 October 2026", or "Draft: …" with no firm. */
export function draftLineText(firm: string, style: DraftingPreferences["dateFormat"], date: Date) {
  return `${firm.trim() ? `${firm.trim()} draft` : "Draft"}: ${formatDraftDate(date, style)}`
}

type LayoutPreferences = Pick<DraftingPreferences, "firm" | "dateFormat" | "pageNumbers" | "pageNumberPosition" | "draftLine" | "agreementLegend">

/** What an agreement's first page carries: header lines in order, then the
 * footer's page number, as the renderer writes them. */
export function agreementFurniture(prefs: LayoutPreferences, date: Date) {
  const legend = prefs.agreementLegend.trim()
  return {
    header: [...(legend ? [legend] : []), ...(prefs.draftLine ? [draftLineText(prefs.firm, prefs.dateFormat, date)] : [])],
    pageNumber: prefs.pageNumbers === "off" ? undefined : { text: prefs.pageNumbers === "page" ? "Page 1" : "1", align: prefs.pageNumberPosition },
  }
}
