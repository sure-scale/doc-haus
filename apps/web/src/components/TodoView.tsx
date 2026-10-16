// One item of the assistant's completeness self-audit, as the upstream todowrite
// tool / SessionTodo records it: a brief description, a lifecycle status, and a
// priority. Draft mode and reviewers build this checklist to re-check every
// term, clause, number, and cross-reference item-by-item before finalizing.
export type Todo = { content: string; status: string; priority: string }

// Map each lifecycle status to a glyph and class. cancelled items are work the
// agent consciously dropped (a declined optional clause), so they read struck
// through rather than as a failure.
const STATUS = {
  completed: { glyph: "✓", className: "todo-completed" },
  in_progress: { glyph: "●", className: "todo-progress" },
  cancelled: { glyph: "—", className: "todo-cancelled" },
  pending: { glyph: "○", className: "todo-pending" },
} as const

function statusOf(status: string) {
  return STATUS[status as keyof typeof STATUS] ?? STATUS.pending
}

// In-chat render of the streamed completeness checklist: draft mode's term /
// number / cross-reference audit and the reviewer's per-area present/absent/
// deviating walk. Collapsed behind a count line like the sources list, with a
// done/total tally so the lawyer sees coverage at a glance. Mirrors how
// CitationView and DraftView card a turn's other tool output.
export default function TodoView({ todos }: { todos: Todo[] }) {
  if (todos.length === 0) return null
  const done = todos.filter((t) => t.status === "completed" || t.status === "cancelled").length
  return (
    <details className="todos">
      <summary>
        Completeness audit ({done}/{todos.length})
      </summary>
      <ul className="todo-list">
        {todos.map((t, i) => {
          const s = statusOf(t.status)
          return (
            <li className={`todo ${s.className}`} key={i}>
              <span className="todo-glyph" aria-hidden>
                {s.glyph}
              </span>
              <span className="todo-content">{t.content}</span>
            </li>
          )
        })}
      </ul>
    </details>
  )
}
