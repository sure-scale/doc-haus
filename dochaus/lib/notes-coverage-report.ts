// Drafting-notes coverage report rendering for the lean gate. Pure: it lists the
// rows that were checked and the gaps.

export type NotesCoverageInput = { rows: { id?: string; location?: string }[]; issues: string[] }

export function renderNotesCoverage(coverage: NotesCoverageInput) {
  const checked = coverage.rows
    .map((row) => row.id || row.location)
    .filter((row): row is string => Boolean(row))
    .map((row) => `- ${row.replace(/\s+/g, " ").trim()}`)
  const gaps = coverage.issues.map((issue) => `- ${issue.replace(/\s+/g, " ").trim()}`)
  return [
    "# Drafting Notes Coverage Report",
    "",
    `Status: ${coverage.issues.length ? "PARTIAL" : "PASS"}`,
    "",
    "## Checked Rows",
    "",
    ...(checked.length ? checked : ["None."]),
    "",
    "## Coverage Gaps",
    "",
    ...(gaps.length ? gaps : ["None."]),
    "",
    "<!-- notes-coverage-report:end -->",
    "",
  ].join("\n")
}
