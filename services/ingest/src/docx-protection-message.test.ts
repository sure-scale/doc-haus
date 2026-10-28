import { expect, test } from "bun:test"
import { docxProtectionMessage } from "./docx-protection-message"

test("protected Word content failures have actionable wording without internal codes or identifiers", () => {
  for (const code of ["docx_protected_inline_content:private-proposal-id", "docx_protected_inline_content_changed", "docx_field_replacement:private-proposal-id"]) {
    const message = docxProtectionMessage(new Error(code))!
    expect(message).toContain("Word")
    expect(message).toContain("prepare a fresh change")
    expect(message).not.toContain("docx_")
    expect(message).not.toContain("private-proposal-id")
    expect(message).not.toContain("unchanged")
  }
  expect(docxProtectionMessage(new Error("docx_protected_inline_content_changed_extra"))).toBeUndefined()
  expect(docxProtectionMessage(new Error("unrelated"))).toBeUndefined()
})
