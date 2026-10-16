// Run from dochaus/: bun test skill/drafting-format.test.ts
import { describe, expect, test } from "bun:test"

describe("drafting skill formatting guidance", () => {
  test("requires document-family spacing and explicit open items", async () => {
    const skill = await Bun.file("skill/drafting/SKILL.md").text()
    expect(skill).toContain("[open item: ...]")
    expect(skill).toContain("one blank line between blocks")
    expect(skill).toContain("standalone structural elements and input fields")
    expect(skill).toContain("spacing their document family requires")
  })

  test("distinguishes working placeholders from final readiness", async () => {
    const skill = (await Bun.file("skill/drafting/SKILL.md").text()).replace(/\s+/g, " ")

    expect(skill).toContain("working draft")
    expect(skill).toContain("never permission to replace an unknown fact, authority, or legal effect")
    expect(skill).toContain("do not call the document final or ready for external use")
    expect(skill).toContain("memo disclosure cannot make invented body text acceptable")
  })

  test("verification requires a semantic unresolved-content check", async () => {
    const skill = (await Bun.file("skill/verify-work-product/SKILL.md").text()).replace(/\s+/g, " ")

    expect(skill).toContain("semantic unresolved-content check")
    expect(skill).toContain("does not by itself make a final deliverable ready")
    expect(skill).toContain("Never turn missing facts or authority into neutral complete-looking prose")
  })
})
