import type { ChronologyCoverage, ChronologyDate, ChronologyEvent } from "./chronology"

const unique = (values: string[]) => [...new Set(values.filter(Boolean))]
const months = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
]

/** Only these explicit calendar spellings are equivalent; source prose is never inferred away. */
function calendarDate(value: string): { label: string; forms: string[] } {
  const parts = /^(\d{4})(?:-(\d{2})(?:-(\d{2}))?)?$/.exec(value)
  if (!parts) return { label: value, forms: [value] }
  const [, year, month, day] = parts
  if (!month) return { label: year, forms: [value] }
  const name = months[Number(month) - 1]
  if (!name) return { label: value, forms: [value] }
  const short = name.slice(0, 3)
  if (!day) return { label: `${short} ${year}`, forms: [value, `${short} ${year}`, `${name} ${year}`] }
  const parsed = new Date(0)
  parsed.setUTCFullYear(Number(year), Number(month) - 1, Number(day))
  if (parsed.getUTCMonth() !== Number(month) - 1 || parsed.getUTCDate() !== Number(day))
    return { label: value, forms: [value] }
  return {
    label: `${Number(day)} ${short} ${year}`,
    forms: [
      value,
      `${Number(day)} ${short} ${year}`,
      `${day} ${short} ${year}`,
      `${Number(day)} ${name} ${year}`,
      `${day} ${name} ${year}`,
    ],
  }
}

function dateDetails(date: ChronologyDate): string[] {
  return [
    date.rangeKind ? (date.rangeKind === "duration" ? "Known period" : "Uncertain date range") : "",
    date.time
      ? `Time: ${date.time}${date.timezone ? ` ${date.timezone}` : " (timezone not supplied)"}`
      : date.timezone
        ? `Timezone: ${date.timezone}`
        : "",
    date.qualifier ?? "",
    date.anchorFreshness && date.anchorFreshness !== "current"
      ? `Date anchor ${date.anchorFreshness}: ${date.anchorIssue ?? "check the supporting source"}`
      : "",
  ]
}

export function chronologyExportDate(date: ChronologyDate): string {
  const start = date.start ? calendarDate(date.start) : { label: "Undated", forms: ["Undated"] }
  const end = date.end && date.end !== date.start ? calendarDate(date.end) : undefined
  const period = `${start.label}${end ? ` to ${end.label}` : ""}`
  const forms = end ? start.forms.flatMap((from) => end.forms.map((to) => `${from} to ${to}`)) : start.forms
  const comparable = (value: string) => value.trim().toLowerCase()
  const originalIsCalendar = forms.some((form) => comparable(form) === comparable(date.original))
  const precision = {
    exact: "",
    approximate: "Approximate",
    month: "Month only",
    year: "Year only",
    ambiguous: "Ambiguous date",
    relative: "Relative date",
    undated: date.start ? "Recorded as undated" : "",
  }[date.precision]
  return unique([period, originalIsCalendar ? "" : date.original, precision, ...dateDetails(date)]).join(" · ")
}

/** Evidence notes retain the normalized values and original expression without compacting either. */
function fullDate(date: ChronologyDate): string {
  const period = date.start ? `${date.start}${date.end && date.end !== date.start ? ` to ${date.end}` : ""}` : "Undated"
  return unique([
    period,
    date.original !== period ? date.original : "",
    date.precision !== "exact" ? `Date precision: ${date.precision}` : "",
    ...dateDetails(date),
  ]).join("; ")
}

export function chronologyExportQualifications(event: ChronologyEvent): string[] {
  return unique([
    event.reviewed ? "Reviewed" : "Unreviewed",
    event.hidden ? "Hidden entry explicitly included" : "",
    event.key ? "Key event" : "",
    event.disputed ? "Conflict or dispute" : "",
    event.origin === "manual" ? "Manual entry" : "",
    event.lawyerEdited ? "Edited by lawyer" : "",
    !event.accounts.length ? "Source not supplied" : "",
    ...event.accounts.flatMap((account) => [
      `Privilege: ${account.privilege}`,
      account.status !== "reported" ? `${account.status} event, not confirmation it occurred` : "",
      !account.evidence.length ? "Account has no supporting source" : "",
      ...account.evidence.flatMap((evidence) => [
        evidence.freshness !== "current" || !evidence.verified
          ? `Source support ${evidence.freshness === "current" ? "unverified" : evidence.freshness}`
          : "",
        evidence.provenance === "supplied-chronology" ? "Supplied chronology account" : "",
      ]),
    ]),
  ])
}

function dateNotes(date: ChronologyDate): string[] {
  return [
    `Date: ${fullDate(date)}`,
    ...(date.documentDate ? [`Document date: ${date.documentDate}`] : []),
    ...(date.relative
      ? [
          `Relative expression: ${date.relative.expression}`,
          `Anchor: ${date.relative.anchorDocument}${date.relative.anchorDate ? ` dated ${date.relative.anchorDate}` : ""}${date.relative.anchorVersion ? ` (version ${date.relative.anchorVersion})` : ""}`,
          ...(date.relative.calculation ? [`Calculation: ${date.relative.calculation}`] : []),
          `Interpretation: ${date.relative.interpretation ?? "not recorded"}`,
        ]
      : []),
    ...(date.correction
      ? [
          `Date corrected by ${date.correction.actor} at ${new Date(date.correction.at).toISOString()}; original: ${fullDate(date.correction.original)}`,
        ]
      : []),
  ]
}

export type ChronologyExportNote = { text: string; bold?: boolean }
const note = (text: string, options: { bold?: boolean } = {}): ChronologyExportNote => ({ text, ...options })
export function chronologyExportEventNotes(event: ChronologyEvent): ChronologyExportNote[] {
  return [
    ...dateNotes(event.date).map((text) => ({ text })),
    note(chronologyExportQualifications(event).join("; ")),
    ...(event.detail ? [note(event.detail)] : []),
    ...(event.disputed ? [note(`Dispute: ${event.dispute || "No explanation recorded"}`)] : []),
    ...(event.note ? [note(`Lawyer note: ${event.note}`)] : []),
    ...(event.reviewedAt
      ? [
          note(
            `Review recorded at ${new Date(event.reviewedAt).toISOString()}${event.reviewedBy ? ` by ${event.reviewedBy}` : ""}`,
          ),
        ]
      : []),
    ...(event.original
      ? [
          note(`Original entry: ${event.original.title}`),
          note(event.original.detail),
          note(`Original date: ${fullDate(event.original.date)}`),
        ]
      : []),
    ...event.accounts.flatMap((account, accountIndex) => [
      note(`Account ${accountIndex + 1}: ${account.title}`, { bold: true }),
      ...(account.attribution ? [note(`Attribution: ${account.attribution}`)] : []),
      note(account.detail),
      ...dateNotes(account.date).map((text) => ({ text })),
      note(`Status: ${account.status}; privilege: ${account.privilege}; significance: ${account.significance}`),
      ...(account.people.length ? [note(`People: ${account.people.join(", ")}`)] : []),
      ...account.gaps.map((gap) => note(`Gap: ${gap}`)),
      ...(!account.evidence.length ? [note("No supporting source supplied for this account.")] : []),
      ...account.evidence.flatMap((evidence) => [
        note(`Source: ${evidence.documentName}${evidence.section ? ` — ${evidence.section}` : ""}`, { bold: true }),
        note(
          `Version: ${evidence.version}; support: ${evidence.verified ? "text matched" : "unverified"}; freshness: ${evidence.freshness}; origin: ${evidence.provenance === "source" ? "source document" : "supplied chronology"}`,
        ),
        ...(evidence.issue ? [note(`Source qualification: ${evidence.issue}`)] : []),
        note(
          evidence.verified && evidence.freshness === "current" && evidence.excerpt
            ? `Excerpt: “${evidence.excerpt}”`
            : "Excerpt omitted because current verified source text is unavailable.",
        ),
      ]),
    ]),
  ]
}

export function chronologyExportCoverageNotes(coverage: ChronologyCoverage): ChronologyExportNote[] {
  return [
    note(coverage.documentName, { bold: true }),
    note(
      `Status: ${coverage.status}; sections processed: ${coverage.processedSections.length} of ${coverage.sections}; entries found: ${coverage.events}; text extracted: ${coverage.textExtracted ? "yes" : "no"}; currently uploaded: ${coverage.uploaded ? "yes" : "no"}.`,
    ),
    note(
      `Version: ${coverage.version}${coverage.suppliedChronology ? "; supplied chronology, not independent corroboration" : ""}`,
    ),
    ...coverage.issues.map((issue) => note(`Coverage qualification: ${issue}`)),
  ]
}
