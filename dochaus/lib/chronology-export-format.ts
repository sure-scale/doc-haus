import type { ChronologyCoverage, ChronologyDate, ChronologyEvent } from "./chronology"

const unique = (values: string[]) => [...new Set(values.filter(Boolean))]
export function chronologyExportDate(date: ChronologyDate): string {
  const period = date.start ? `${date.start}${date.end && date.end !== date.start ? ` to ${date.end}` : ""}` : "Undated"
  return unique([
    period,
    date.original !== period ? date.original : "",
    date.precision !== "exact" ? `Date precision: ${date.precision}` : "",
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
    `Date: ${chronologyExportDate(date)}`,
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
          `Date corrected by ${date.correction.actor} at ${new Date(date.correction.at).toISOString()}; original: ${chronologyExportDate(date.correction.original)}`,
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
          note(`Original date: ${chronologyExportDate(event.original.date)}`),
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
