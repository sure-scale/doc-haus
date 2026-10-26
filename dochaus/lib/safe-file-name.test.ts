import { expect, test } from "bun:test"
import { safeFileName } from "./safe-file-name"

test("reserved characters become a space, or nothing before a dot or the end", () => {
  expect(safeFileName("RE: Offer.msg")).toBe("RE Offer.msg")
  expect(safeFileName('Draft "v2" <Acme|Beta>.docx')).toBe("Draft v2 Acme Beta.docx")
  expect(safeFileName("Can we meet?.msg")).toBe("Can we meet.msg")
  expect(safeFileName("Notes*")).toBe("Notes")
  expect(safeFileName("a\\b/c.pdf")).toBe("a b c.pdf")
  expect(safeFileName("Line\u0001one\ttwo.pdf")).toBe("Line one two.pdf")
  expect(safeFileName("???")).toBe("")
})

test("device names keep their place with a mark, with or without an extension", () => {
  expect(safeFileName("nul")).toBe("nul_")
  expect(safeFileName("NUL.docx")).toBe("NUL_.docx")
  expect(safeFileName("con.tar.gz")).toBe("con_.tar.gz")
  expect(safeFileName("Aux .pdf")).toBe("Aux_ .pdf")
  expect(safeFileName("prn.")).toBe("prn_")
  expect(safeFileName("COM1.msg")).toBe("COM1_.msg")
  expect(safeFileName("lpt9.xlsx")).toBe("lpt9_.xlsx")
  expect(safeFileName("COM¹.txt")).toBe("COM¹_.txt")
  for (const name of ["console.docx", "nullify.pdf", "com10.pdf", "con-tract.docx", "Aux notes.pdf"]) expect(safeFileName(name)).toBe(name)
})

test("leading and trailing dots and spaces are dropped", () => {
  expect(safeFileName("Report.docx. ")).toBe("Report.docx")
  expect(safeFileName("Notes...")).toBe("Notes")
  expect(safeFileName("  .hidden.pdf")).toBe("hidden.pdf")
})

test("a safe name, Unicode included, comes back unchanged", () => {
  for (const name of [
    "Share Purchase Agreement (2) - final.docx",
    "Smith & Jones v. Brown, 12.03.2026.pdf",
    "Two  spaces.docx",
    "Vertrag – Müller 契約書.docx",
    "😀 notes.md",
  ])
    expect(safeFileName(name)).toBe(name)
  expect(safeFileName("Résumé: été.pdf")).toBe("Résumé été.pdf")
})
