import { normalizePracticeAreas } from "./drafting-preferences"
import { readFirmPreferences } from "./preferences"

/** Firm-wide background, never a classification of the current matter. */
export function firmPracticeAreaContext(workspace?: string) {
  const areas = normalizePracticeAreas(readFirmPreferences(workspace).practiceAreas)
  if (!areas.length) return ""
  // Encode markup characters so names remain data inside this context block.
  const names = JSON.stringify(areas).replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026")
  return `<practice_areas>\nThe firm works in these practice areas: ${names}. Consider these as background when doing relevant work or helping maintain library resources. Do not assume any area is the subject of this matter, establishes a legal fact or governing law, or makes a resource applicable. Follow the lawyer's instructions and the actual sources.\n</practice_areas>`
}
