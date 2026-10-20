import { expect, test } from "bun:test"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../lib/render"
import { presentDraft } from "./draft-presentation"

test("a draft presented without its title loses only the title, not a later Heading 1", async () => {
  const rendered = await renderMarkdownToDocx("# Notice letter\n\n## Opening\n\nWhy we write.\n\n# Later heading\n\nWhat follows.\n")
  const { bytes } = presentDraft(rendered.bytes, [{ heading: "Opening", display_heading: "" }], false)
  const body = strFromU8(unzipSync(bytes)["word/document.xml"]!)
  expect(body).not.toContain("Notice letter")
  expect(body).toContain("Later heading")
  expect(body.match(/w:val="Heading1"/g)).toHaveLength(1)
})
