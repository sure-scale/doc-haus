import type { ReactNode } from "react"
import { icons, ShapeIcon, type IconName } from "../icons"
import "../../../styles/dashboard-setup.css"

// Serializable card content. Runtime actions and interactive slots stay in the
// composition layer, so the same cards can render a JSON view or live settings.
export type SetupCardItem = {
  id: string
  icon: IconName
  title: string
  detail: string
  done: boolean
  editable?: boolean
  action?: { id: string; label: string; variant?: "text" | "outline" }
}
export type SetupChecklistConfig = { title: string; description: string; items: SetupCardItem[] }
export type DemoDataCardConfig = {
  title: string
  description: string
  features: { label: string; value?: string; icon?: IconName }[]
  action: { id: string; label: string; icon?: IconName }
  note: string
}

export function SetupItem({ item, onAction, children }: { item: SetupCardItem; onAction?: (action: string) => void; children?: ReactNode }) {
  return <li className={`dash-setup-item${item.done ? " is-done" : ""}${item.editable ? " is-editable" : ""}`}>
    <div className="dash-setup-item-main">
      <span className="dash-setup-step-icon" aria-hidden="true">{item.done ? icons.check : icons[item.icon]}</span>
      <div className="dash-setup-item-text"><h3 id={`dash-setup-${item.id}`}>{item.title}{item.done && <span className="sr-only">, complete</span>}</h3><p>{item.detail}</p></div>
      {item.action && <button type="button" className={item.action.variant === "outline" ? "dash-setup-outline" : "dash-text-button"} disabled={!onAction} onClick={() => onAction?.(item.action!.id)}>{item.action.label}{item.action.variant === "outline" && icons.arrowRight}</button>}
    </div>
    {children}
  </li>
}

export function SetupChecklist({ config, onAction, renderContent, accordion }: { config: SetupChecklistConfig; onAction?: (action: string) => void; renderContent?: (id: string) => ReactNode; accordion?: { expanded: string; onToggle: (id: string) => void } }) {
  if (accordion) return <section className="dash-setup-accordion" aria-label={config.title}>
    <ol className="dash-setup-list">{config.items.map((item, index) => {
      const expanded = accordion.expanded === item.id
      return <li key={item.id} className={`dash-setup-accordion-item${item.done ? " is-done" : ""}${expanded ? " is-expanded" : ""}`}>
        <h3 className="dash-setup-accordion-heading"><button type="button" id={`dash-setup-toggle-${item.id}`} className="dash-setup-toggle" aria-expanded={expanded} aria-controls={`dash-setup-panel-${item.id}`} onClick={() => accordion.onToggle(expanded ? "" : item.id)}>
          <span className="dash-setup-number" aria-hidden="true">{index + 1}</span>
          <span className="dash-setup-row-text"><span id={`dash-setup-${item.id}`} className="dash-setup-row-title">{item.title}{item.done && <span className="sr-only">, complete</span>}</span><span className="dash-setup-row-detail">{item.detail}</span></span>
          {item.done ? <span className="dash-setup-completed" aria-hidden="true">{icons.check}</span> : <span className={`dash-setup-row-action${config.items.find(row => !row.done)?.id === item.id ? " is-next" : ""}`} aria-hidden="true">Set up</span>}
          <span className="dash-setup-chevron" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><path d="m9 5 7 7-7 7" /></svg></span>
        </button></h3>
        <div id={`dash-setup-panel-${item.id}`} className="dash-setup-panel" role="region" aria-labelledby={`dash-setup-${item.id}`} hidden={!expanded}>
          {renderContent?.(item.id)}
          {item.action && <button type="button" className="dash-setup-outline" disabled={!onAction} onClick={() => onAction?.(item.action!.id)}>{item.action.label}{icons.arrowRight}</button>}
        </div>
      </li>
    })}</ol>
  </section>
  const count = config.items.filter((item) => item.done).length
  return <section className="dash-setup-checklist" aria-label={config.title}>
    <header className="dash-setup-head"><ShapeIcon icon="setup" /><div><h2>{config.title}</h2><p>{config.description}</p></div><span className="dash-checklist-progress"><progress value={count} max={config.items.length || 1} aria-label="Setup progress" />{count} of {config.items.length} complete</span></header>
    <ol className="dash-setup-list">{config.items.map((item) => <SetupItem key={item.id} item={item} onAction={onAction}>{renderContent?.(item.id)}</SetupItem>)}</ol>
  </section>
}

export function DemoDataCard({ config, onAction, busy = false, disabled = false, error }: { config: DemoDataCardConfig; onAction?: (action: string) => void; busy?: boolean; disabled?: boolean; error?: string }) {
  return <aside className="dash-setup-demo" aria-label={config.title}>
    <header><ShapeIcon icon="sample" /><div><h2>{config.title}</h2><p>{config.description}</p></div></header>
    <div className="dash-setup-demo-features">{config.features.map((feature) => <div key={feature.label}>{feature.value ? <strong>{feature.value}</strong> : feature.icon ? icons[feature.icon] : null}<span>{feature.label}</span></div>)}</div>
    <button type="button" className="dash-setup-demo-button" onClick={() => onAction?.(config.action.id)} disabled={disabled || busy || !onAction} aria-busy={busy}>{config.action.icon && icons[config.action.icon]}{config.action.label}</button>
    <p className="dash-setup-demo-note">{config.note}</p>{error && <p className="dash-setup-demo-error" role="alert">{error}</p>}
  </aside>
}
