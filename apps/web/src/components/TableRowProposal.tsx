import { useId } from "react"
import type { Redline } from "../api/ingest/redlines"
import { ProposalRedline } from "./redline"
import "./TableRowProposal.css"

type Proposal = Pick<Redline, "scope" | "table_row" | "target_json" | "stale">
type RowView = NonNullable<Redline["table_row"]>
const record = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === "object" && !Array.isArray(value))
const cells = (value: unknown): value is string[] => Array.isArray(value) && value.length > 0 && value.length <= 32 && value.every(cell => typeof cell === "string")

export function ordinaryProposal(proposal: Proposal): boolean {
  if (proposal.scope !== "phrase" && proposal.scope !== "clause" || proposal.table_row !== undefined) return false
  try {
    const binding: unknown = JSON.parse(proposal.target_json ?? "null")
    if (record(binding) && (binding.version === 3 || binding.kind === "table_row_insert")) return false
  } catch { /* Legacy text bindings keep their existing stale/error handling. */ }
  return true
}

/** Display the frozen new cells even if the reference cannot be resolved. Never
 * reconstruct cells or coordinates from legacy joined-text summaries. */
function rowView(value: unknown): RowView | undefined {
  if (!record(value) || value.kind !== "table_row_insert" || value.position !== "before" && value.position !== "after" || !cells(value.cells)) return
  const proposed = { kind: "table_row_insert" as const, position: value.position as RowView["position"], cells: value.cells }
  if (typeof value.table === "number" && Number.isSafeInteger(value.table) && value.table > 0 &&
      typeof value.row === "number" && Number.isSafeInteger(value.row) && value.row > 0 &&
      cells(value.reference_cells) && value.reference_cells.length === value.cells.length)
    return { ...proposed, table: value.table, row: value.row, reference_cells: value.reference_cells }
  return proposed
}

/** Shared by the individual and bulk actions. The server remains the write
 * authority; a row must also have a complete readable reference before review. */
export function reviewableProposal(proposal: Proposal): boolean {
  if (proposal.stale) return false
  if (ordinaryProposal(proposal)) return true
  return proposal.scope === "table_row" && rowView(proposal.table_row)?.reference_cells !== undefined
}

export function ProposalPreview({ proposal }: { proposal: Redline }) {
  if (ordinaryProposal(proposal)) return <ProposalRedline before={proposal.old_text} after={proposal.new_text} />
  return <TableRowProposal view={proposal.scope === "table_row" ? rowView(proposal.table_row) : undefined} stale={proposal.stale} />
}

function TableRowProposal({ view, stale }: { view?: RowView; stale?: boolean }) {
  const id = useId()
  if (!view) return <p className="table-row-unavailable">Row insertion preview unavailable. Reject this proposal and prepare a fresh change list.</p>
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
