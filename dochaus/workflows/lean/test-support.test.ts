import { expect, test } from "bun:test"
import { checkNotesTargets } from "./focused-notes-evidence"
import { draftingNotes } from "./test-support"

test("the shared notes fixture distinguishes target linking from prose without inventing inspected evidence", () => {
  const reply = draftingNotes("<current_target_lookup>\nLocate the relevant accepted paragraph.\n</current_target_lookup>", 0)
  const checked = checkNotesTargets(JSON.parse(reply), [])
  expect(checked.issues).toEqual([])
  expect(checked.targets).toEqual([])
  expect(checked.limitation).toContain("does not inspect current passages")
  expect(draftingNotes("<section_kind>\nopen-item\n</section_kind>\nExplain the current_target_lookup limitation in prose.", 0))
    .toBe("The document records the supplied terms. Review any stated unresolved point before using the draft.")
  expect(JSON.parse(draftingNotes("This is a separate repair-link job", 0))).toEqual({
    sections: [0], reason: "The synthetic package explanation owns this generic memo correction.",
  })
})
