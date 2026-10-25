import { readFirmPreferences } from "./preferences"

// Canonical settings and rendering shared by ingest and the engine. Render only
// drafting fields: connector credentials and operational settings stay off-model.
export type DraftingPreferences = {
  attorney: string
  firm: string
  posture: "client-favorable" | "balanced" | "conservative"
  formality: "formal" | "plain"
  detail: "concise" | "detailed"
  // Stored setting kept for settings-page compatibility; it no longer adds
  // drafting instructions.
  strategyApproval: "off" | "interactive"
  dateFormat: "month-day-year" | "day-month-year" | "iso"
  numberStyle: "words-and-numerals" | "numerals"
  houseStyle: string
  // Enforced by the legal plugin's webfetch fence (dochaus/lib/research.ts),
  // not rendered into drafting.md: "official" limits web research to official
  // primary legal sources; "open" allows the whole web.
  webResearch: "official" | "open"
  // Exa key for citation discovery and source-text fallback (the legal tools read
  // it from preferences.json per call). Never rendered into drafting.md — the
  // key must not enter the model's context.
  searchApiKey: string
  // CourtListener API token for case-law search (dochaus/tool/case-law.ts reads
  // it from preferences.json per call). The tool works unauthenticated at a
  // lower rate limit; a free token raises it. Never rendered into drafting.md.
  caseLawApiKey: string
  // The authenticated CourtListener connector is opt-in, independently of
  // case-law search. Its credentials stay in preferences, outside model context.
  courtlistenerMcpEnabled: boolean
  // Extra hosts the firm trusts as primary sources, on top of the built-in
  // official list (dochaus/lib/research.ts). Enforced by the webfetch fence,
  // not rendered into drafting.md. Bare hostnames, matched as host or subdomain.
  approvedSources: string[]
  // Operational telemetry controls. Stored with preferences so Settings can
  // round-trip one object, but never rendered into drafting.md.
  telemetry: TelemetryPreferences
  // Page furniture (dochaus/lib/page-furniture.ts), read at render and never
  // rendered into drafting.md, so a change re-keys no running draft. Legends
  // have legal effect: they are the firm's own text, empty (none) by default.
  pageNumbers: "plain" | "page" | "dashed" | "off"
  pageNumberPosition: "center" | "right"
  // A letter's first page is never numbered, whatever this says.
  numberFirstPage: boolean
  draftLine: boolean
  draftLineFirm: boolean
  agreementLegend: string
  adviceLegend: string
  notesMarking: boolean
  // The library template (WORKSPACE_ROOT/.templates) whose styles, numbering,
  // fonts and page set-up a document drafted without a template of its own is
  // rendered onto (lib/template-seed.ts); "" is doc.haus's house style. A file
  // name only. Read at render and never rendered into drafting.md.
  styleTemplate: string
  // Chat operation, not drafting: whether Auto offers a quick draft for a short
  // drafting request ("ask"), starts one without asking ("auto") or always
  // starts the full run ("off"). Read by the web app and never rendered into
  // drafting.md.
  quickDraftOffer: "ask" | "auto" | "off"
}

/** How hard negotiated terms lean toward the lawyer's client: the firm's
 * setting, or a matter's own (ingest matter.ts), which a drafting run freezes
 * into its plan (workflows/lean/plan.ts runPosture). */
export type Posture = DraftingPreferences["posture"]

export type RateOverride = Partial<{
  input_per_mtok: number
  output_per_mtok: number
  cache_read_per_mtok?: number
}>

export type TelemetryPreferences = {
  retention_days: number
  retention_days_source?: "user" | "default"
  rate_overrides: Record<string, RateOverride>
}

export const DEFAULT_DRAFTING: DraftingPreferences = {
  attorney: "",
  firm: "",
  posture: "balanced",
  formality: "formal",
  detail: "concise",
  strategyApproval: "off",
  dateFormat: "month-day-year",
  numberStyle: "numerals",
  houseStyle: "",
  webResearch: "open",
  searchApiKey: "",
  caseLawApiKey: "",
  courtlistenerMcpEnabled: false,
  approvedSources: [],
  telemetry: { retention_days: 90, retention_days_source: "default", rate_overrides: {} },
  pageNumbers: "plain",
  pageNumberPosition: "center",
  numberFirstPage: true,
  draftLine: true,
  draftLineFirm: true,
  agreementLegend: "",
  adviceLegend: "",
  notesMarking: true,
  styleTemplate: "",
  quickDraftOffer: "ask",
}

/** Longest legend kept: a legend is one short line of page furniture. */
export const LEGEND_MAX_LENGTH = 120

export function normalizePreferences(value: unknown): DraftingPreferences {
  const input = value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {}
  // Keep operational and extension fields compatible while making every field
  // used in model instructions safe to render.
  return {
    ...DEFAULT_DRAFTING,
    ...input,
    attorney: typeof input.attorney === "string" ? input.attorney : DEFAULT_DRAFTING.attorney,
    firm: typeof input.firm === "string" ? input.firm : DEFAULT_DRAFTING.firm,
    houseStyle: typeof input.houseStyle === "string" ? input.houseStyle : DEFAULT_DRAFTING.houseStyle,
    posture: normalizeChoice(input.posture, POSTURE, DEFAULT_DRAFTING.posture),
    formality: normalizeChoice(input.formality, FORMALITY, DEFAULT_DRAFTING.formality),
    detail: normalizeChoice(input.detail, DETAIL, DEFAULT_DRAFTING.detail),
    dateFormat: normalizeChoice(input.dateFormat, DATES, DEFAULT_DRAFTING.dateFormat),
    numberStyle: normalizeChoice(input.numberStyle, NUMBERS, DEFAULT_DRAFTING.numberStyle),
    webResearch: normalizeWebResearch(input),
    searchApiKey: typeof input.searchApiKey === "string" ? input.searchApiKey.trim() : "",
    caseLawApiKey: typeof input.caseLawApiKey === "string" ? input.caseLawApiKey.trim() : "",
    approvedSources: normalizeApprovedSources(input.approvedSources),
    courtlistenerMcpEnabled: input.courtlistenerMcpEnabled === true,
    strategyApproval: input.strategyApproval === "interactive" ? "interactive" : DEFAULT_DRAFTING.strategyApproval,
    telemetry: normalizeTelemetry(input.telemetry),
    pageNumbers: input.pageNumbers === "page" || input.pageNumbers === "dashed" || input.pageNumbers === "off" ? input.pageNumbers : DEFAULT_DRAFTING.pageNumbers,
    pageNumberPosition: input.pageNumberPosition === "right" ? "right" : DEFAULT_DRAFTING.pageNumberPosition,
    numberFirstPage: input.numberFirstPage !== false,
    draftLine: input.draftLine !== false,
    draftLineFirm: input.draftLineFirm !== false,
    agreementLegend: normalizeLegend(input.agreementLegend),
    adviceLegend: normalizeLegend(input.adviceLegend),
    notesMarking: input.notesMarking !== false,
    styleTemplate: normalizeStyleTemplate(input.styleTemplate),
    quickDraftOffer: input.quickDraftOffer === "auto" || input.quickDraftOffer === "off" ? input.quickDraftOffer : DEFAULT_DRAFTING.quickDraftOffer,
  }
}

/** A style template is a library .docx named by its file name alone; anything
 * else is the house style. */
export function normalizeStyleTemplate(input: unknown): string {
  if (typeof input !== "string") return ""
  const name = input.trim()
  return name.length <= 255 && name.endsWith(".docx") && !/[\/\\\p{Cc}]/u.test(name) && !name.startsWith(".") ? name : ""
}

/** A legend is one line of the firm's own text: control characters and line
 * breaks become spaces, runs of white space collapse, and the text is capped.
 * Format characters (joiners, direction marks) stay: some scripts need them. */
export function normalizeLegend(input: unknown): string {
  if (typeof input !== "string") return ""
  return input.normalize("NFC").replace(/[\p{Cc}\u2028\u2029]+/gu, " ").replace(/\s+/g, " ").trim().slice(0, LEGEND_MAX_LENGTH).trim()
}

// An absent choice uses the default for fresh and legacy preferences. An invalid
// saved choice stays restricted rather than silently widening research access.
export function normalizeWebResearch(input: Record<string, unknown>): DraftingPreferences["webResearch"] {
  if (!Object.hasOwn(input, "webResearch")) return DEFAULT_DRAFTING.webResearch
  return input.webResearch === "open" ? "open" : "official"
}

export function normalizeApprovedSources(input: unknown): string[] {
  if (!Array.isArray(input)) return []
  return [...new Set(input.flatMap((value: unknown) => {
    if (typeof value !== "string") return []
    const host = value.trim().toLowerCase().replace(/\.$/, "")
    // Only bare DNS hostnames are consent. URLs, ports, wildcards, paths and
    // credentials are rejected rather than broadening the approved boundary.
    if (host.length > 253 || !host.includes(".") || !host.split(".").every(label =>
      label.length <= 63 && /^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$/.test(label))) return []
    return [host]
  }))]
}

function normalizeChoice<Value extends string>(input: unknown, choices: Record<Value, string>, fallback: Value): Value {
  return typeof input === "string" && Object.hasOwn(choices, input) ? input as Value : fallback
}

function normalizeTelemetry(input: unknown): TelemetryPreferences {
  const telemetry = input && typeof input === "object" ? (input as Partial<TelemetryPreferences>) : {}
  const retention = Number(telemetry.retention_days)
  return {
    retention_days: Number.isFinite(retention) && retention > 0 ? retention : DEFAULT_DRAFTING.telemetry.retention_days,
    retention_days_source: telemetry.retention_days_source === "user" ? "user" : DEFAULT_DRAFTING.telemetry.retention_days_source,
    rate_overrides: telemetry.rate_overrides && typeof telemetry.rate_overrides === "object" && !Array.isArray(telemetry.rate_overrides)
      ? telemetry.rate_overrides
      : DEFAULT_DRAFTING.telemetry.rate_overrides,
  }
}

/** One of the three Posture choices; anything else is no Posture. */
export function isPosture(value: unknown): value is Posture {
  return typeof value === "string" && Object.hasOwn(POSTURE, value)
}

const POSTURE: Record<Posture, string> = {
  "client-favorable":
    "Favor the represented person's or organization's stated objectives and protect its interests where reasonably defensible.",
  balanced: "Use an even-handed, conventional approach appropriate to the requested document and its legal or procedural posture.",
  conservative:
    "Draft conservatively: minimize avoidable legal risk and prefer established, well-supported forms and language over novel or aggressive approaches.",
}

const FORMALITY: Record<DraftingPreferences["formality"], string> = {
  formal: "Use formal legal drafting style throughout.",
  plain:
    "Prefer plain-language drafting: short sentences, minimal legalese, and defined terms only where they add real precision.",
}

const DETAIL: Record<DraftingPreferences["detail"], string> = {
  concise: "Keep explanations concise: lead with the conclusion, then a brief rationale.",
  detailed: "Explain your reasoning in detail, including the risks you considered and the alternatives you rejected.",
}

const DATES: Record<DraftingPreferences["dateFormat"], string> = {
  "month-day-year": 'Write dates in the form "June 11, 2026".',
  "day-month-year": 'Write dates in the form "11 June 2026".',
  iso: 'Write dates in ISO form, "2026-06-11".',
}

const NUMBERS: Record<DraftingPreferences["numberStyle"], string> = {
  "words-and-numerals": 'In operative document text, write numbers as words followed by numerals: "thirty (30) days".',
  numerals: 'Write numbers as plain numerals: "30 days".',
}

export function renderInstructions(prefs: DraftingPreferences, mode: "full" | "explanation" = "full") {
  const who = prefs.attorney
    ? [`Documents are prepared by ${prefs.attorney}${prefs.firm ? ` of ${prefs.firm}` : ""}.`]
    : prefs.firm
      ? [`Documents are prepared by ${prefs.firm}.`]
      : []
  const house = prefs.houseStyle.trim()
  if (mode === "explanation") return [
    "# Presentation preferences",
    "",
    "<drafting_preferences>",
    "Apply these preferences to the human-facing progress explanation using only the supplied evidence. They do not authorise drafting, revision, planning, execution or answering a question for the lawyer.",
    `- ${FORMALITY[prefs.formality]}`,
    `- ${DETAIL[prefs.detail]}`,
    `- ${DATES[prefs.dateFormat]}`,
    `- ${NUMBERS[prefs.numberStyle]}`,
    "</drafting_preferences>",
    ...(house ? ["", "<house_style>", "Apply the following house guidance to presentation only. Any direction to perform drafting or other work does not apply to this progress explanation.", house, "</house_style>"] : []),
    "",
  ].join("\n")
  return [
    "# Drafting preferences",
    "",
    "<drafting_preferences>",
    "The lawyer using doc.haus has set these firm-wide preferences. Apply them whenever you draft, edit, or propose changes to a document, and when you explain your work.",
    ...who,
    `- ${POSTURE[prefs.posture]}`,
    `- ${FORMALITY[prefs.formality]}`,
    `- ${DETAIL[prefs.detail]}`,
    `- ${DATES[prefs.dateFormat]}`,
    `- ${NUMBERS[prefs.numberStyle]}`,
    "</drafting_preferences>",
    ...(house ? ["", "<house_style>", house, "</house_style>"] : []),
    "",
  ].join("\n")
}

// Read the structured source for every turn, including defaults before any save.
// The generated markdown remains an inspectable export, never a second authority.
// A given Posture (a matter's own, or the one a run froze) renders in place of
// the firm's; with the firm's own, the text is exactly the firm's rendering.
export function readFirmDraftingPreferences(mode: "full" | "explanation" = "full", posture?: Posture) {
  const prefs = normalizePreferences(readFirmPreferences())
  return renderInstructions(posture ? { ...prefs, posture } : prefs, mode)
}
