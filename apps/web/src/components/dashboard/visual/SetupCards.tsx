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

export function SetupChecklist({ config, onAction, renderContent }: { config: SetupChecklistConfig; onAction?: (action: string) => void; renderContent?: (id: string) => ReactNode }) {
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
