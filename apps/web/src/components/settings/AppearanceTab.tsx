import { useState } from "react"
import { loadPrefs, savePrefs, type Prefs } from "../../prefs"
import { SegRow } from "./controls"

// Text size, applied instantly (prefs.ts reflects it onto <html>).
export default function AppearanceTab() {
  const [prefs, setPrefs] = useState<Prefs>(loadPrefs)
  const set = (patch: Partial<Prefs>) => setPrefs(savePrefs(patch))
  return (
    <section className="settings-section">
      <p className="settings-hint muted">Display preferences for this browser. Changes apply and save immediately.</p>
      <SegRow
        label="Text size"
        hint="The size of text throughout the interface."
        value={prefs.textSize}
        onChange={(textSize) => set({ textSize: textSize as Prefs["textSize"] })}
        options={[
          { value: "compact", label: "Compact" },
          { value: "standard", label: "Standard" },
          { value: "large", label: "Large" },
        ]}
      />
      <SegRow
        label="Contrast"
        hint="Use your system setting or increase contrast in doc.haus."
        value={prefs.contrast}
        onChange={(contrast) => set({ contrast: contrast as Prefs["contrast"] })}
        options={[
          { value: "system", label: "System" },
          { value: "high", label: "High" },
        ]}
      />
      <div className="settings-type-preview">
        <span>Text preview</span>
        <h4>A comfortable place to work</h4>
        <p>Read documents, follow the conversation and review the details at a size that works for you.</p>
      </div>
    </section>
  )
}
