import { createContext, useContext, type ReactNode } from "react"
import { createPortal } from "react-dom"

export const SettingsFooterContext = createContext<{ target: HTMLElement | null; tab: string } | null>(null)

export function SettingsFooter({ tab, children }: { tab: string; children: ReactNode }) {
  const footer = useContext(SettingsFooterContext)
  if (!footer) return <>{children}</>
  return footer.target && footer.tab === tab ? createPortal(children, footer.target) : null
}
