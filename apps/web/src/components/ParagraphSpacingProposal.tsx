import { useId } from "react"
import type { Redline } from "../api/ingest/redlines"
import "./ParagraphSpacingProposal.css"

type Proposal = Pick<Redline, "scope" | "paragraph_spacing" | "table_row" | "target_json" | "stale">
type Before = NonNullable<Redline["paragraph_spacing"]>["before_direct_after"]
type Location = { kind: "body" } | { kind: "table"; table: number; row: number; column: number }
type Spacing = { paragraph: string; before: Before; after: string; row: string[] | null; location?: Location }
const KIND = "paragraph_spacing_after"
const record = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === "object" && !Array.isArray(value))
const positive = (value: unknown): value is number => typeof value === "number" && Number.isSafeInteger(value) && value > 0
const literal = (value: unknown, max: number): value is string => typeof value === "string" && value.length <= max && value.isWellFormed() && !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value)
const rowText = (value: unknown): value is string[] => Array.isArray(value) && value.length <= 64 && value.every(cell => literal(cell, 65536))

export function paragraphSpacingMarker(proposal: Proposal): boolean {
  if (proposal.scope === "paragraph_spacing" || proposal.paragraph_spacing !== undefined) return true
  try {
    const binding: unknown = JSON.parse(proposal.target_json ?? "null")
    return record(binding) && (binding.version === 4 || binding.kind === KIND)
  } catch { return false }
}

function beforeValue(value: unknown): Before | undefined {
  if (!record(value)) return
  if (value.kind === "absent" && Object.keys(value).length === 1) return { kind: "absent" }
  if (value.kind === "twips" && Object.keys(value).length === 3 && typeof value.raw === "string" && /^\d{1,32}$/.test(value.raw) &&
      typeof value.normalized === "string" && /^(0|[1-9]\d{0,31})$/.test(value.normalized) && BigInt(value.raw).toString() === value.normalized)
    return { kind: "twips", raw: value.raw, normalized: value.normalized }
}

function tableLocation(value: unknown, row: string[]): Location | undefined {
  if (!record(value) || value.unavailable_reason !== undefined || !positive(value.table) || !positive(value.row) || !positive(value.column) ||
      value.row_span !== 1 || value.column_span !== 1 || value.vertical_merge !== "none" || !rowText(value.row_context) ||
      JSON.stringify(value.row_context) !== JSON.stringify(row) || !row.length || !Array.isArray(value.row_cells) || value.row_cells.length !== row.length) return
  let end = 1, target = false
  for (let index = 0; index < value.row_cells.length; index++) {
    const cell: unknown = value.row_cells[index]
    if (!record(cell) || !positive(cell.column) || !positive(cell.column_span) || !Number.isSafeInteger(cell.column + cell.column_span) || cell.column < end ||
        typeof cell.row_span !== "number" || !Number.isSafeInteger(cell.row_span) || cell.row_span < 0 || cell.vertical_merge !== "none" && cell.vertical_merge !== "restart" && cell.vertical_merge !== "continue" ||
        (cell.vertical_merge === "continue" ? cell.row_span !== 0 : cell.row_span < 1) || cell.text !== row[index]) return
    if (cell.column === value.column) target = cell.column_span === 1 && cell.row_span === 1 && cell.vertical_merge === "none"
    end = cell.column + cell.column_span
  }
  return target ? { kind: "table", table: value.table, row: value.row, column: value.column } : undefined
}

function spacingValue(proposal: Proposal): Spacing | undefined {
  if (proposal.scope !== "paragraph_spacing" || proposal.table_row !== undefined) return
  if (proposal.target_json != null) {
    try {
      const binding: unknown = JSON.parse(proposal.target_json)
      if (!record(binding) || binding.version !== 4 || binding.kind !== KIND) return
    } catch { return }
  }
  const value: unknown = proposal.paragraph_spacing
  if (!record(value) || value.kind !== KIND || !literal(value.paragraph_text, 8192) || !value.paragraph_text.length ||
      value.row_context !== null && !rowText(value.row_context) || typeof value.after_direct_after_twips !== "string" ||
      !/^(0|[1-9]\d{0,2})$/.test(value.after_direct_after_twips) || BigInt(value.after_direct_after_twips) > 960n) return
  const before = beforeValue(value.before_direct_after)
  if (!before || before.kind === "twips" && before.normalized === value.after_direct_after_twips) return
  const row = value.row_context as string[] | null
  const location = value.table_context === null && row === null ? { kind: "body" as const }
    : row !== null ? tableLocation(value.table_context, row) : undefined
  return { paragraph: value.paragraph_text, before, after: value.after_direct_after_twips, row, location }
}

export function reviewableParagraphSpacing(proposal: Proposal): boolean {
  return !proposal.stale && spacingValue(proposal)?.location !== undefined
}

function points(twips: string): string {
  const value = BigInt(twips), whole = value / 20n
  const fraction = ((value % 20n) * 5n).toString().padStart(2, "0").replace(/0+$/, "")
  return `${whole}${fraction ? `.${fraction}` : ""} pt`
}

function LiteralText({ text }: { text: string }) {
  return <><span className="paragraph-spacing-literal">{text}</span>{text === "" ? <span className="paragraph-spacing-note">Empty cell</span>
    : text.trim() === "" && <span className="paragraph-spacing-note">Whitespace only</span>}</>
}

export function ParagraphSpacingProposal({ proposal }: { proposal: Redline }) {
  const id = useId(), view = spacingValue(proposal)
  if (!view) return <p className="paragraph-spacing-unavailable">Spacing preview unavailable. Reject this proposal and prepare a fresh change list.</p>
  const location = proposal.stale ? undefined : view.location
  const contextLabel = location ? "Table context" : "Original table context"
  return <div className="paragraph-spacing-proposal">
    <p className="paragraph-spacing-title">Spacing after paragraph</p>
    <p className="paragraph-spacing-help">Wording unchanged · direct setting only</p>
    {!location && <p className="paragraph-spacing-unavailable">Location needs refreshing.</p>}
    {location?.kind === "table" && <p className="paragraph-spacing-location">Table {location.table} · Row {location.row} · Column {location.column}</p>}
    <div className="paragraph-spacing-text"><LiteralText text={view.paragraph} /></div>
    <div className="paragraph-spacing-change" role="group" aria-label="Direct spacing after">
      <dl><dt>Before</dt><dd>{view.before.kind === "absent" ? "Not set" : points(view.before.normalized)}</dd></dl><span aria-hidden="true">→</span><dl><dt>After</dt><dd>{points(view.after)}</dd></dl>
    </div>
    {view.row && view.row.length > 0 && <details className="paragraph-spacing-context">
      <summary>{contextLabel}</summary>
      <p id={`${id}-scroll`} className="sr-only">Scroll horizontally to read every cell.</p>
      <div className="paragraph-spacing-scroll" role="region" tabIndex={0} aria-label={contextLabel} aria-describedby={`${id}-scroll`}>
        <table style={{ minWidth: `${view.row.length * 7}rem` }}><caption className="sr-only">{contextLabel}. Cell text in original order; wording unchanged.</caption>
          <thead><tr>{view.row.map((_cell, index) => <th scope="col" key={index}>Cell text</th>)}</tr></thead>
          <tbody><tr>{view.row.map((cell, index) => <td key={index}><LiteralText text={cell} /></td>)}</tr></tbody>
        </table>
      </div>
    </details>}
  </div>
}
