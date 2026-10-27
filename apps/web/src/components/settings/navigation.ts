import { createContext, useContext } from "react"
import type { Tab } from "./types"

export const SettingsNavigation = createContext<(tab: Tab) => void>(() => {})
export const useOpenSettings = () => useContext(SettingsNavigation)
