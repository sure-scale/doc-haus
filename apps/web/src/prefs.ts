import { readBrowserJSON, writeBrowserValue } from "./browser-storage"

// Per-browser UI preferences (Settings → Approvals / Appearance / Matter
// defaults). localStorage, not engine config, because these shape how THIS
// browser presents and gates the engine — the engine has no notion of them.
// Drafting preferences are different: they steer the model, so they live on the
// engine as an instructions file (see api/ingest.ts getPreferences).
export type Prefs = {
  textSize: "compact" | "standard" | "large"
  // Auto-approve the engine's permission asks for redline/tracked-change
  // proposals. Safe to automate: the tools only record pending tracked changes
  // the lawyer still accepts or rejects in the document viewer.
  autoApproveRedlines: boolean
  // Auto-approve individual document/template creation requests. The separate
  // Draft workflow and direct-edit permissions do not consume this preference.
  autoApproveDrafting: boolean
  // Jurisdiction codes prefilled on the new-matter form.
  defaultJurisdictions: string[]
}

export const DEFAULT_PREFS: Prefs = {
  textSize: "standard",
  autoApproveRedlines: false,
  autoApproveDrafting: false,
  defaultJurisdictions: [],
}

const KEY = "dochaus.prefs"

export function loadPrefs(): Prefs {
  const saved = readBrowserJSON(KEY)
  const values = typeof saved === "object" && saved !== null && !Array.isArray(saved) ? saved : {}
  return {
    ...DEFAULT_PREFS,
    ...values,
    autoApproveRedlines: "autoApproveRedlines" in values && values.autoApproveRedlines === true,
    autoApproveDrafting: "autoApproveDrafting" in values && values.autoApproveDrafting === true,
  }
}

export function savePrefs(patch: Partial<Prefs>, options: { requireStorage?: boolean } = {}): Prefs {
  const next = { ...loadPrefs(), ...patch }
  const saved = writeBrowserValue(KEY, JSON.stringify(next))
  if (!saved && options.requireStorage) throw new Error("Could not save this preference in your browser. Your previous setting is still in effect.")
  applyAppearance(next)
  return next
}

// Reflect text size onto <html> as the data attribute the stylesheet keys off
// ([data-text] scales the UI).
export function applyAppearance(prefs: Prefs = loadPrefs()) {
  document.documentElement.dataset.text = prefs.textSize
}

// Which permission asks each auto-approve toggle covers. Direct edits and
// redactions are absent. Native engine permissions, including separately saved
// Always allowances, are evaluated before any request reaches this browser.
const REDLINE_PERMISSIONS = new Set(["tracked-changes", "redline"])
const DRAFT_PERMISSIONS = new Set(["draft-document", "create-template"])

export function autoApproved(permission: string): boolean {
  const prefs = loadPrefs()
  if (REDLINE_PERMISSIONS.has(permission)) return prefs.autoApproveRedlines
  if (DRAFT_PERMISSIONS.has(permission)) return prefs.autoApproveDrafting
  return false
}
