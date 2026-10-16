import { useState } from "react"
import { loadPrefs, savePrefs, type Prefs } from "../../prefs"
import { SegRow } from "./controls"

// Text size, applied instantly (prefs.ts reflects it onto <html>).
export default function AppearanceTab() {
  const [prefs, setPrefs] = useState<Prefs>(loadPrefs)
  const set = (patch: Partial<Prefs>) => setPrefs(savePrefs(patch))
  return (
    <section className="settings-section">
      <h3>Appearance</h3>
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
    </section>
  )
}
