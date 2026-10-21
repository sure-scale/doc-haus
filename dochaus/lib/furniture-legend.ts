// The firm's legend a document carries, by its form (#203). Pure, so the web's
// quick draft applies the same legend the renderer writes (page-furniture.ts)
// without importing the renderer. Legends have legal effect: they are the
// firm's own text, empty (none) by default, and never written by a model.

/** The firm's legend for a document of `kind` (a planned form, or "notes"):
 * the agreement legend on an agreement, the advice legend on advice and an
 * advice letter, and none on anything else. */
export function furnitureLegend(kind: string, prefs: { agreementLegend: string; adviceLegend: string }) {
  if (kind === "agreement") return prefs.agreementLegend
  return kind === "advice" || kind === "advice letter" ? prefs.adviceLegend : ""
}
