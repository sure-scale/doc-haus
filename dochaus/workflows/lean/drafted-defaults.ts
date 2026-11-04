import { planActingFor } from "./plan"
import { openItem } from "./quotes"
import type { OpenItem } from "./state"

/** Points the run decided for the lawyer where the request is silent, each
 * disclosed once in the drafting notes rather than marked in every paragraph
 * it shapes: the side taken from the sources. Each id keys on its content, so
 * a changed determination is a new item. */
export function draftedDefaultItems(directory: string): OpenItem[] {
  const acting = planActingFor(directory)
  return acting.basis === "inferred"
    ? [openItem("intake", "drafted_default", `The documents are drafted for ${acting.party}, a side taken from what the requested document does and whom the sources show instructing it${acting.reason ? ` (${acting.reason})` : ""}, as the request names none. The lawyer should confirm the side, or name the party the lawyer acts for.`, { repair_target: "memo" })]
    : []
}
