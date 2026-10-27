import type { PaneState } from "./sync"

export type SavePresentation = {
  badge: string
  tone: "neutral" | "changed" | "saving" | "saved" | "blocked"
  message: string
  help?: string
  mismatch?: string
  primary: "save" | "close"
}

/** A matching version ID is connection information, never a content receipt. */
export function savePresentation(state: PaneState): SavePresentation {
  const version = state.connection?.version
  const mismatch = version?.base && version.current && version.base.id !== version.current.id
    ? `This copy: version ${version.base.number}. In doc.haus: version ${version.current.number}.`
    : undefined
  const result: SavePresentation = {
    badge: "Ready", tone: "neutral", primary: "save", mismatch,
    message: version?.current ? `Version ${version.current.number} in doc.haus` : "Save to confirm this copy in doc.haus.",
  }
  if (state.status === "saving") return { ...result, badge: "Saving…", tone: "saving", message: state.message }
  if (version && version.state !== "current") {
    const copy = {
      stale: ["A newer version is saved in doc.haus.", "Load the latest version in a separate Word file. Keep any edits in this older copy with Save As."],
      removed: ["Saving is unavailable for this copy.", "This copy predates a redaction or removal. Keep your local file and reopen the current document from doc.haus."],
      "unrecorded-change": ["The matter file changed outside its saved history.", "Keep your Word file and resolve that change in doc.haus before saving."],
    }[version.state]
    return { ...result, badge: "Saving paused", tone: "blocked", message: copy[0], help: copy[1] }
  }
  if (state.connection?.lease.state === "other") return {
    ...result, badge: "Saving paused", tone: "blocked", message: "Another Word session is editing this document.",
    help: "Your local changes stay in Word. Check status or deliberately take over editing here.",
  }
  if (state.problem) return { ...result, badge: "Needs attention", tone: "blocked", message: state.message }
  if (state.status === "saved" || state.status === "done") {
    const receipt = state.acknowledgement
    const time = receipt && new Date(receipt.at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    const message = receipt ? `Version ${receipt.number} received at ${time}.` : "Save confirmed by doc.haus."
    return {
      ...result, badge: "Saved", tone: "saved", primary: "close",
      message: state.status === "done" && state.message.includes("Close it in Word") ? `${message} Close it in Word.` : message,
      help: receipt?.overlap ? "Saved as a separate version because the matter document also changed. Compare the versions in doc.haus." : undefined,
    }
  }
  if (state.content === "changed") return { ...result, badge: "Changes to save", tone: "changed", message: "Save to confirm your Word changes in doc.haus." }
  return result
}
