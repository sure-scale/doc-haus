import { useId } from "react"
import type { Redline } from "../api/ingest/redlines"
import { ProposalRedline } from "./redline"
import "./TableRowProposal.css"
import { ParagraphSpacingProposal, paragraphSpacingMarker, reviewableParagraphSpacing } from "./ParagraphSpacingProposal"

type Proposal = Pick<Redline, "scope" | "table_row" | "paragraph_spacing" | "target_json" | "stale">
type RowView = NonNullable<Redline["table_row"]>
type InsertView = Extract<RowView, { kind: "table_row_insert" }>
type DeleteView = Extract<RowView, { kind: "table_row_delete" }>
const record = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === "object" && !Array.isArray(value))
const cells = (value: unknown): value is string[] => Array.isArray(value) && value.length > 0 && value.length <= 32 && value.every(cell => typeof cell === "string")

export function ordinaryProposal(proposal: Proposal): boolean {
  if (proposal.scope !== "phrase" && proposal.scope !== "clause" || proposal.table_row !== undefined || paragraphSpacingMarker(proposal)) return false
  try {
    const binding: unknown = JSON.parse(proposal.target_json ?? "null")
    if (record(binding) && (binding.version === 3 || binding.kind === "table_row_insert" || binding.kind === "table_row_delete")) return false
  } catch { /* Legacy text bindings keep their existing stale/error handling. */ }
  return true
}

/** Display the frozen cells even if the location cannot be resolved. Never
 * reconstruct cells or coordinates from legacy joined-text summaries. */
function rowView(value: unknown): RowView | undefined {
  if (!record(value) || !cells(value.cells)) return
  const located = typeof value.table === "number" && Number.isSafeInteger(value.table) && value.table > 0 &&
    typeof value.row === "number" && Number.isSafeInteger(value.row) && value.row > 0
  if (value.kind === "table_row_delete") {
    // Deletion is its own operation; insertion fields cannot change its meaning.
    if ("position" in value || "reference_cells" in value) return
    const deleted = { kind: "table_row_delete" as const, cells: value.cells }
    return located ? { ...deleted, table: value.table as number, row: value.row as number } : deleted
  }
  if (value.kind !== "table_row_insert" || value.position !== "before" && value.position !== "after") return
  const proposed: InsertView = { kind: "table_row_insert", position: value.position, cells: value.cells }
  if (located && cells(value.reference_cells) && value.reference_cells.length === value.cells.length)
    return { ...proposed, table: value.table as number, row: value.row as number, reference_cells: value.reference_cells }
  return proposed
}

/** Shared by the individual and bulk actions. The server remains the write
 * authority; a row must also have its full readable content and location. */
export function reviewableProposal(proposal: Proposal): boolean {
  if (proposal.stale) return false
  if (paragraphSpacingMarker(proposal)) return reviewableParagraphSpacing(proposal)
  if (ordinaryProposal(proposal)) return true
  if (proposal.scope !== "table_row") return false
  const view = rowView(proposal.table_row)
  return view?.kind === "table_row_insert" ? view.reference_cells !== undefined : view?.table !== undefined && view.row !== undefined
}

export function ProposalPreview({ proposal }: { proposal: Redline }) {
  if (paragraphSpacingMarker(proposal)) return <ParagraphSpacingProposal proposal={proposal} />
  if (ordinaryProposal(proposal)) return <ProposalRedline before={proposal.old_text} after={proposal.new_text} />
  return <TableRowProposal view={proposal.scope === "table_row" ? rowView(proposal.table_row) : undefined} stale={proposal.stale} />
}

function TableRowProposal({ view, stale }: { view?: RowView; stale?: boolean }) {
  if (!view) return <p className="table-row-unavailable">Table row preview unavailable. Reject this proposal and prepare a fresh change list.</p>
  return view.kind === "table_row_delete" ? <DeletedRow view={view} stale={stale} /> : <InsertedRow view={view} stale={stale} />
}

function DeletedRow({ view, stale }: { view: DeleteView; stale?: boolean }) {
  const id = useId()
  const located = !stale && view.table !== undefined && view.row !== undefined
  const title = located ? `Delete row ${view.row} in table ${view.table}` : "Delete a row"
  return <div className="table-row-proposal">
    <p id={`${id}-title`} className="table-row-title">{title}</p>
    {!located && <p className="table-row-unavailable">Location needs refreshing.</p>}
    <p id={`${id}-scroll`} className="sr-only">Scroll horizontally to read every column.</p>
    <div className="table-row-scroll" role="region" tabIndex={0} aria-labelledby={`${id}-title`} aria-describedby={`${id}-scroll`}>
      <table style={{ minWidth: `${5.25 + view.cells.length * 7}rem` }}>
        <caption className="sr-only">{title}. The complete row shown will be removed.</caption>
        <thead><tr><th scope="col"><span className="sr-only">Row</span></th>{view.cells.map((_cell, column) => <th scope="col" key={column}>Column {column + 1}</th>)}</tr></thead>
        <tbody><tr className="table-row-delete"><th scope="row">Row to delete</th>{view.cells.map((cell, column) => <td key={column}>
          <span className="table-row-cell-text">{cell}</span>
          {cell === "" ? <span className="table-row-empty" title="Empty cell"><span aria-hidden="true">—</span><span className="sr-only">Empty cell</span></span>
            : cell.trim() === "" && <span className="table-row-cell-note">Whitespace only</span>}
        </td>)}</tr></tbody>
      </table>
    </div>
  </div>
}

function InsertedRow({ view, stale }: { view: InsertView; stale?: boolean }) {
  const id = useId()
  // Stale geometry is never displayed, even if a cached response still contains
  // old reference cells. The stored position and proposed values remain visible.
  const reference = !stale && view.reference_cells !== undefined ? view : undefined
  const title = reference
    ? `Insert a row ${view.position} row ${reference.row} in table ${reference.table}`
    : `Insert a row ${view.position} the reference row`
  const rows = reference
    ? view.position === "before" ? ["new", "reference"] as const : ["reference", "new"] as const
    : ["new"] as const
  return <div className="table-row-proposal">
    <p id={`${id}-title`} className="table-row-title">{title}</p>
    {!reference && <p className="table-row-unavailable">Reference location needs refreshing.</p>}
    <p id={`${id}-scroll`} className="sr-only">Scroll horizontally to read every column.</p>
    <div className="table-row-scroll" role="region" tabIndex={0} aria-labelledby={`${id}-title`} aria-describedby={`${id}-scroll`}>
      <table style={{ minWidth: `${5.25 + view.cells.length * 7}rem` }}>
        <caption className="sr-only">{title}. {reference ? "The reference row is unchanged; the new row is proposed." : "Proposed new row; the reference row is unavailable."}</caption>
        <thead><tr><th scope="col"><span className="sr-only">Row</span></th>{view.cells.map((_cell, column) => <th scope="col" key={column}>Column {column + 1}</th>)}</tr></thead>
        <tbody>{rows.map(kind => <tr key={kind} className={kind === "new" ? "table-row-new" : "table-row-reference"}>
          <th scope="row">{kind === "new" ? "New row" : "Reference row"}</th>
          {(kind === "new" ? view.cells : reference!.reference_cells!).map((cell, column) => <td key={column}>{cell === ""
            ? <span className="table-row-empty" title="Empty cell"><span aria-hidden="true">—</span><span className="sr-only">Empty cell</span></span>
            : cell}</td>)}
        </tr>)}</tbody>
      </table>
    </div>
  </div>
}
