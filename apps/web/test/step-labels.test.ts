import { expect, test } from "bun:test"
import { basename } from "../src/components/chat/step-labels"
import { fileLabel } from "../src/components/narration/activity"

// A Windows server sends backslash paths; each names its file as the POSIX path
// does, so redline links, step labels and run-file labels match on both.
test("a Windows path names its file and run files as a POSIX path does", () => {
  expect(basename("/Users/lawyer/matter/Supply Terms.docx")).toBe("Supply Terms.docx")
  expect(basename("C:\\Users\\lawyer\\matter\\Supply Terms.docx")).toBe("Supply Terms.docx")
  expect(basename("/Users/lawyer/matter/")).toBe("matter")
  expect(basename("C:\\Users\\lawyer\\matter\\")).toBe("matter")
  expect(basename("Supply Terms.docx")).toBe("Supply Terms.docx")

  expect(fileLabel("C:\\m\\.dochaus\\runs\\r1\\plan.json")).toBe("the drafting plan")
  expect(fileLabel(".dochaus\\lean\\coverage.json")).toBe("the plan coverage check")
  expect(fileLabel("C:\\m\\sources\\Supply Terms.docx")).toBe("Supply Terms.docx")
})
