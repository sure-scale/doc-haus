import { mkdirSync, watch } from "node:fs"
import path from "node:path"
import { Plugin } from "@opencode/plugin"
import { readFirmPreferences } from "../lib/preferences"
import { setResearchConsentStatus } from "../lib/research-consent-status"

// The hosted connector needs an account. Keep it out of the shipped MCP list;
// a firm explicitly enables it in Research settings and supplies its own token.
export default Plugin.define({
  id: "dochaus.research-mcp",
  setup: async (context) => {
    const location = context.location.directory
    setResearchConsentStatus(location, "pending")
    let failedActivation = false
    const failed = () => {
      failedActivation = true
      setResearchConsentStatus(location, "failed")
      console.error("Research consent settings could not be applied. Research activation is unavailable.")
    }
    const root = process.env.WORKSPACE_ROOT
    if (!root) return failed()
    const directory = path.join(root, ".preferences")
    try {
      mkdirSync(directory, { recursive: true })
      const registration = await context.mcp.transform((editor) => {
        try {
          const denied = editor.get("courtlistener")?.disabled === true
          // This transform also runs after native/account configuration refresh.
          // Never inherit an account's remote connectors into a legal matter.
          for (const [name, config] of editor.list()) {
            if (config.type === "remote" || name === "courtlistener") editor.remove(name)
          }
          const preferences = readFirmPreferences()
          const token = typeof preferences.caseLawApiKey === "string" ? preferences.caseLawApiKey.trim() : ""
          if (!denied && preferences.courtlistenerMcpEnabled === true && token && !/[\r\n\x00]/.test(token)) {
            editor.set("courtlistener", {
              type: "remote",
              url: "https://mcp.courtlistener.com/",
              oauth: false,
              headers: { Authorization: `Token ${token}` },
              timeout: { catalog: 30000, execution: 30000 },
            })
          }
          setResearchConsentStatus(location, failedActivation ? "failed" : "active")
        } catch {
          // Native transforms must not throw and silently fall back to unfiltered
          // state. The engine's readiness boundary rejects failed activation.
          failed()
        }
      })
      // Watch the directory so replacing preferences.json reloads warm matters.
      // Native reconciliation owns reconnects, tool removal and explicit denial.
      let pending: ReturnType<typeof setTimeout> | undefined
      const reload = () => {
        if (!failedActivation) setResearchConsentStatus(location, "pending")
        clearTimeout(pending)
        pending = setTimeout(() => { void context.mcp.reload().catch(failed) }, 100)
      }
      const watcher = watch(directory, (_event, filename) => {
        if (!filename || filename.toString() === "preferences.json") reload()
      })
      watcher.on("error", failed)
      reload()
      return async () => {
        setResearchConsentStatus(location, "pending")
        clearTimeout(pending)
        watcher.close()
        await registration.dispose()
      }
    } catch {
      failed()
    }
  },
})
