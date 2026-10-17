import { useState } from "react"
import type { LeanOpenItem } from "../../api/ingest/lean"
import { capGroups, groupOpenItems, openItemClass, openItemTag } from "./model"

const CAP = 5

// The run's open items behind one expander, grouped by kind. Long lists show the
// first few and a "Show all"; each item reads as the run recorded it. Unresolved
// checklist findings lead the list and are never cut by the cap.
export default function OpenItems({ items, titles }: { items: LeanOpenItem[]; titles: Map<string, string> }) {
  const [all, setAll] = useState(false)
  if (items.length === 0) return null
  const unresolved = items.filter((item) => item.kind === "checklist_unresolved").length
  const cap = Math.max(CAP, unresolved)
  const groups = capGroups(groupOpenItems(items), all ? items.length : cap)

  return (
    <details className="open-items">
      <summary>
        <span>Open items</span>
        <span className="rail-count">{items.length}</span>
        {unresolved > 0 && <span className="open-items-unresolved">{unresolved} not satisfied</span>}
      </summary>
      <div className="open-items-list">
        {groups.map((group) => (
          <section key={group.kind} className="open-items-group">
            <h4>
              {group.label}
              <span className="open-items-count">{group.total}</span>
            </h4>
            <ul>
              {group.items.map((item) => (
                <li key={item.id}>
                  {openItemTag(item) && (
                    <span className={`open-item-tag${item.kind === "checklist_unresolved" ? " is-unresolved" : ""}`}>{openItemTag(item)}</span>
                  )}
                  {openItemClass(item) && <span className="open-item-class">{openItemClass(item)}</span>}
                  {item.ref?.document && titles.size > 1 && titles.get(item.ref.document) && (
                    <span className="open-item-doc">{titles.get(item.ref.document)}</span>
                  )}
                  <span>{item.detail}</span>
                </li>
              ))}
            </ul>
          </section>
        ))}
        {items.length > cap && (
          <button type="button" className="linklike open-items-more" aria-expanded={all} onClick={() => setAll((on) => !on)}>
            {all ? "Show fewer" : `Show all ${items.length}`}
          </button>
        )}
      </div>
    </details>
  )
}
