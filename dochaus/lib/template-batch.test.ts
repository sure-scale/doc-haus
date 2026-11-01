// Run from dochaus/: bun test lib/template-batch.test.ts
import { describe, expect, test } from "bun:test"
import { strFromU8, unzipSync } from "fflate"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"
import {
  autoMatch,
  batchDocumentName,
  batchFolderName,
  identifyingKey,
  normaliseHeader,
  rowFill,
  templateBatchFields,
  type BatchField,
} from "./template-batch"
import { fillTemplateDocx } from "./template-fill"
import { OPEN_ITEM_HIGHLIGHT_FILL } from "./open-items"
import { FILLED_VALUE_FILL } from "./marker-patterns"

const templateMarkdown = [
  "# Synthetic services letter",
  "## 1. Parties\n\nThis letter is addressed to [Client Name] of [Client Address].",
  "## 2. Notice\n\nNotice is given on [insert notice date]. Reference [___].",
  "## 3. Non-solicitation [optional: Non-solicit]\n\nNo solicitation for twelve months.",
  "## 4. General\n\nSynthetic closing text.",
].join("\n\n")
const { bytes: template } = await renderMarkdownToDocx(templateMarkdown)

async function placeholders(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {
    return session.findPlaceholders().map((p) => ({ text: p.match.text, kind: p.kind, hint: p.hint }))
  } finally {
    session.close()
  }
}

function documentText(bytes: Uint8Array) {
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!)
  return xml
    .split(/<\/w:p>/)
    .map((paragraph) => [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(""))
    .join("\n")
}

describe("template batch fields", () => {
  test("derive labels; unnamed blanks are counted, not offered", async () => {
    const { fields, unnamedBlanks } = templateBatchFields(await placeholders(template))
    expect(fields.map(({ key, label, kind }) => ({ key, label, kind }))).toEqual([
      { key: "[Client Name]", label: "Client Name", kind: "fill" },
      { key: "[Client Address]", label: "Client Address", kind: "fill" },
      { key: "[insert notice date]", label: "Notice date", kind: "fill" },
      { key: "[optional: Non-solicit]", label: "Include: Non-solicit", kind: "optional" },
    ])
    expect(unnamedBlanks).toBe(1)
  })

  test("a repeated placeholder is one field", () => {
    const { fields } = templateBatchFields([
      { text: "[Client Name]", kind: "alternative_clause" },
      { text: "[Client Name]", kind: "alternative_clause" },
      { text: "[●]", kind: "alternative_clause" },
    ])
    expect(fields.map((field) => field.key)).toEqual(["[Client Name]"])
  })
})

const fields: BatchField[] = [
  { key: "[Client Name]", label: "Client Name", kind: "fill" },
  { key: "[Client Address]", label: "Client Address", kind: "fill" },
  { key: "[insert notice date]", label: "Notice date", kind: "fill" },
  { key: "[optional: Non-solicit]", label: "Include: Non-solicit", kind: "optional" },
]

describe("auto-match", () => {
  test("normalises headers", () => {
    expect(normaliseHeader("[Insert notice_date]")).toBe("notice date")
    expect(normaliseHeader("Include: Non-solicit")).toBe("non solicit")
    expect(normaliseHeader("  CLIENT-NAME ")).toBe("client name")
  })

  test("exact, then normalised; unmatched stays null", () => {
    expect(autoMatch(fields, ["client_name", "Notice date", "Other"])).toEqual({
      "[Client Name]": "client_name",
      "[Client Address]": null,
      "[insert notice date]": "Notice date",
      "[optional: Non-solicit]": null,
    })
  })

  test("a remembered column wins while present and each column fills one field", () => {
    const remembered = { "[Client Address]": "Client Name", "[insert notice date]": "Gone" }
    expect(autoMatch(fields, ["Client Name", "notice date", "include non solicit"], remembered)).toEqual({
      "[Client Name]": null,
      "[Client Address]": "Client Name",
      "[insert notice date]": "notice date",
      "[optional: Non-solicit]": "include non solicit",
    })
  })
})

describe("row values", () => {
  const columns = ["Client Name", "Client Address", "Notice date", "Include: Non-solicit"]
  const mapping = autoMatch(fields, columns)

  test("blank, bracketed and control-character values are missing; line breaks read as spaces", () => {
    expect(rowFill(fields, columns, ["  ", "1 Example Street\r\nExample Town", "[TBC]", "yes"], mapping)).toEqual({
      fills: { "[Client Address]": "1 Example Street Example Town" },
      include: { "[optional: Non-solicit]": true },
      missing: ["Client Name", "Notice date"],
      missingKeys: ["[Client Name]", "[insert notice date]"],
    })
    expect(rowFill(fields, columns, ["Example\u0007 Ltd", "x", "1 May", "No"], mapping).missing).toEqual(["Client Name"])
  })

  test("optional: yes/no words decide, unmapped or blank keeps, anything else is missing and kept", () => {
    expect(rowFill(fields, columns, ["A", "B", "C", "Omit"], mapping).include).toEqual({ "[optional: Non-solicit]": false })
    expect(rowFill(fields, columns, ["A", "B", "C", "maybe"], mapping)).toMatchObject({
      include: { "[optional: Non-solicit]": true },
      missing: ["Include: Non-solicit"],
    })
    expect(rowFill(fields, columns, ["A", "B", "C", ""], mapping)).toMatchObject({ include: { "[optional: Non-solicit]": true }, missing: [] })
    const unmapped = { ...mapping, "[optional: Non-solicit]": null, "[Client Address]": null }
    expect(rowFill(fields, columns, ["A", "B", "C", "no"], unmapped)).toEqual({
      fills: { "[Client Name]": "A", "[insert notice date]": "C" },
      include: { "[optional: Non-solicit]": true },
      missing: ["Client Address"],
      missingKeys: ["[Client Address]"],
    })
  })
})

describe("naming", () => {
  test("identifying field prefers a name-like label", () => {
    expect(identifyingKey(fields, { "[insert notice date]": "Date", "[Client Name]": "Name" })).toBe("[Client Name]")
    expect(identifyingKey(fields, { "[insert notice date]": "Date" })).toBe("[insert notice date]")
    expect(identifyingKey(fields, {})).toBeNull()
  })

  test("document and folder names", () => {
    expect(batchDocumentName("Services Letter.docx", "Example Holdings Ltd", 8, 3)).toBe("Services Letter – Example Holdings Ltd – row 008.docx")
    expect(batchDocumentName("Services Letter.docx", "  ", 12, 3)).toBe("Services Letter – row 012.docx")
    expect(batchDocumentName("Services Letter.docx", "Example: A/B", 8, 3)).toBe("Services Letter – Example A B – row 008.docx")
    expect(batchFolderName("Services Letter.docx", new Date(2026, 9, 31, 12))).toBe("Services Letter – 2026-10-31")
  })
})

test("fifty rows fill fifty documents, each with only its own values and no network call", async () => {
  const { fields } = templateBatchFields(await placeholders(template))
  const columns = ["Client Name", "Client Address", "Notice date", "Include: Non-solicit"]
  const mapping = autoMatch(fields, columns)
  const rows = Array.from({ length: 50 }, (_, index) => [
    `Example Holdings ${index + 1} Ltd`,
    `${index + 1} Example Street`,
    index === 7 ? "" : `${index + 1} May 2026`,
    index % 2 ? "no" : "yes",
  ])
  const originalFetch = globalThis.fetch
  globalThis.fetch = (() => {
    throw new Error("network call during a batch fill")
  }) as unknown as typeof fetch
  try {
    for (const [index, row] of rows.entries()) {
      const fill = rowFill(fields, columns, row, mapping)
      const { bytes, remaining } = await fillTemplateDocx(template, fill)
      const text = documentText(bytes)
      expect(text).toContain(`addressed to Example Holdings ${index + 1} Ltd of ${index + 1} Example Street.`)
      for (const other of [index - 1, index + 1].filter((n) => n >= 0 && n < 50))
        expect(text).not.toContain(`Example Holdings ${other + 1} Ltd`)
      expect(text).not.toContain("[optional:")
      expect(text.includes("No solicitation for twelve months")).toBe(index % 2 === 0)
      if (index === 7) {
        expect(fill.missing).toEqual(["Notice date"])
        expect(remaining).toEqual(["[insert notice date]", "[___]"])
        expect(text).toContain("Notice is given on [insert notice date].")
      } else {
        expect(text).toContain(`Notice is given on ${index + 1} May 2026.`)
        expect(remaining).toEqual(["[___]"])
      }
    }
  } finally {
    globalThis.fetch = originalFetch
  }
}, 120_000)

test("a named placeholder left unfilled is shaded", async () => {
  const { fields } = templateBatchFields(await placeholders(template))
  const columns = ["Client Address"]
  const { bytes, remaining } = await fillTemplateDocx(template, rowFill(fields, columns, ["1 Example Street"], autoMatch(fields, columns)))
  expect(remaining).toContain("[Client Name]")
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!)
  const shaded = [...xml.matchAll(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:shd\b[^>]*w:fill="([0-9A-F]+)"(?:(?!<\/w:r>)[\s\S])*?<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)]
    .map((match) => [match[2], match[1]])
  expect(shaded).toContainEqual(["[Client Name]", OPEN_ITEM_HIGHLIGHT_FILL])
})

function shadedTexts(bytes: Uint8Array) {
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!)
  return [...xml.matchAll(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:shd\b[^>]*w:fill="([0-9A-F]+)"(?:(?!<\/w:r>)[\s\S])*?<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)]
    .map((match) => [match[2], match[1]])
}

test("the preview marks exactly the values it filled in; a created document is never marked", async () => {
  const { fields } = templateBatchFields(await placeholders(template))
  const columns = ["Client Name", "Client Address"]
  const fill = rowFill(fields, columns, ["Example Holdings Ltd", "1 Example Street"], autoMatch(fields, columns))
  const preview = shadedTexts((await fillTemplateDocx(template, fill, { markFilled: true })).bytes)
  expect(preview).toContainEqual(["Example Holdings Ltd", FILLED_VALUE_FILL])
  expect(preview).toContainEqual(["1 Example Street", FILLED_VALUE_FILL])
  expect(preview.filter(([, colour]) => colour === FILLED_VALUE_FILL).map(([text]) => text).sort()).toEqual(["1 Example Street", "Example Holdings Ltd"])
  expect(preview).toContainEqual(["[insert notice date]", OPEN_ITEM_HIGHLIGHT_FILL])
  const created = await fillTemplateDocx(template, fill)
  expect(shadedTexts(created.bytes).some(([, colour]) => colour === FILLED_VALUE_FILL)).toBe(false)
  expect(documentText(created.bytes)).toBe(documentText((await fillTemplateDocx(template, fill, { markFilled: true })).bytes))
})

test("a kept component whose marker opens its paragraph starts with its own text", async () => {
  const { bytes: opening } = await renderMarkdownToDocx(
    "# Synthetic letter\n\nOpening text.\n\n[optional: Confidentiality] This request is confidential.\n\nClosing text.",
  )
  const { fields } = templateBatchFields(await placeholders(opening))
  const columns = ["Include: Confidentiality"]
  const { bytes } = await fillTemplateDocx(opening, rowFill(fields, columns, ["yes"], autoMatch(fields, columns)))
  const lines = documentText(bytes).split("\n")
  expect(lines).toContain("This request is confidential.")
  expect(lines.some((line) => /^\s+This request/.test(line))).toBe(false)
})

test("an unclear include answer keeps the component with its marker as an open item", async () => {
  const { fields } = templateBatchFields(await placeholders(template))
  const columns = ["Include: Non-solicit"]
  const { bytes, remaining } = await fillTemplateDocx(template, rowFill(fields, columns, ["maybe"], autoMatch(fields, columns)))
  expect(documentText(bytes)).toContain("[optional: Non-solicit]")
  expect(remaining).toContain("[optional: Non-solicit]")
})
