import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { openRevision, verifyPreservation } from "./docx-amendments"
import { applyAmendments } from "./drafting-amendment-apply"
import { ensureParagraphPointers, paragraphPointers, paragraphSha256, type Amendment } from "./drafting-amendment"
import { highlightOpenItemsInDocx } from "./open-items"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const DELIVERY = "The Supplier shall deliver the **Goods** within thirty days of the order date."
const PAYMENT = "Payment is due within *sixty days* after the invoice date, and late payments accrue interest."
const STEADY = "This paragraph stays unchanged with **bold text** inside."
const NOTICE = "Notices must be **in writing** and sent by post."
const plain = (text: string) => text.replace(/\*/g, "")

describe("pointer-based paragraph amendments", () => {
  let root: string
  let seed: Uint8Array
  beforeAll(async () => {
    root = await mkdtemp(path.join(tmpdir(), "dh-amend-"))
    seed = (await renderMarkdownToDocx([
      "# Supply terms", "", DELIVERY, "", PAYMENT, "", STEADY, "",
      "| Term | Detail |", "| --- | --- |", `| Notice | ${NOTICE} |`, "",
    ].join("\n"))).bytes
  })
  afterAll(async () => { await rm(root, { recursive: true, force: true }) })

  async function fixture(name: string) {
    const file = path.join(root, `${name}.docx`)
    await writeFile(file, seed)
    const pointers = await ensureParagraphPointers(file)
    const pid = (text: string) => Object.entries(pointers).find(([, value]) => value === plain(text))![0]
    return { file, out: path.join(root, `${name}.out.docx`), pointers, pid }
  }

  // A one-paragraph document.
  async function single(name: string, markdown: string) {
    const file = path.join(root, `${name}.docx`)
    await writeFile(file, (await renderMarkdownToDocx(markdown)).bytes)
    return { file, out: path.join(root, `${name}.out.docx`), pid: Object.keys(await ensureParagraphPointers(file))[0]! }
  }

  test("disjoint paragraphs apply by code; formatting outside edited words and other paragraphs is kept", async () => {
    const doc = await fixture("disjoint")
    const delivery = "The Vendor shall deliver the Goods within ten days of the order date."
    const payment = "Payment is due within sixty days after the invoice date."
    const notice = "Notices must be in writing and sent by post or email."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("a1", doc.pid(DELIVERY), plain(DELIVERY), delivery),
      amend("a2", doc.pid(PAYMENT), plain(PAYMENT), payment),
      amend("a3", doc.pid(NOTICE), plain(NOTICE), notice),
    ] })
    expect(result).toEqual({ applied: [
      { id: "a1", pid: doc.pid(DELIVERY), mode: "code" },
      { id: "a2", pid: doc.pid(PAYMENT), mode: "code" },
      { id: "a3", pid: doc.pid(NOTICE), mode: "code" },
    ], unapplied: [] })
    const before = await texts(doc.file)
    const after = await texts(doc.out)
    expect(after.get(doc.pid(DELIVERY))).toBe(delivery)
    expect(after.get(doc.pid(PAYMENT))).toBe(payment)
    expect(after.get(doc.pid(NOTICE))).toBe(notice)
    expect(after.get(doc.pid(STEADY))).toBe(before.get(doc.pid(STEADY)))
    const formatting = await measure(doc.file, doc.out)
    expect(formatting.changed).toBe(3)
    expect(formatting.outside).toBe(0)
    expect(formatting.preservation).toBe("ok")
    // "Goods" sits between two separate edits and keeps its own bold run.
    expect(formatting.word(doc.pid(DELIVERY), "Goods")).toEqual({ before: true, after: true })
    expect(formatting.word(doc.pid(PAYMENT), "sixty days")).toEqual({ before: true, after: true })
  })

  test("an amendment on a paragraph an earlier amendment in the batch changed is unapplied, never merged", async () => {
    const overlap = "An earlier amendment in this batch already changed this paragraph."
    const doc = await fixture("overlap")
    const pid = doc.pid(DELIVERY)
    const first = "The Supplier shall deliver the Goods within ten days of the order date."
    const second = "The Supplier shall deliver the Goods within thirty days of the delivery request."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("first", pid, plain(DELIVERY), first),
      amend("second", pid, plain(DELIVERY), second),
      amend("repeat", pid, plain(DELIVERY), first),
      amend("steady", doc.pid(STEADY), plain(STEADY), plain(STEADY)),
    ] })
    // A text already there is already_applied only on a paragraph no amendment in the batch changed.
    expect(result.applied).toEqual([{ id: "first", pid, mode: "code" }, { id: "steady", pid: doc.pid(STEADY), mode: "already_applied" }])
    expect(result.unapplied).toEqual([
      { id: "second", pid, reason: "stale", detail: { current: first, target: second, reason: "Reflect the agreed position.", issue: overlap } },
      { id: "repeat", pid, reason: "stale", detail: { current: first, target: first, reason: "Reflect the agreed position.", issue: overlap } },
    ])
    expect((await texts(doc.out)).get(pid)).toBe(first)
    expect((await measure(doc.file, doc.out)).word(pid, "Goods")).toEqual({ before: true, after: true })

    // Two pointers left on one paragraph are one paragraph: the second is an overlap too.
    const extra = "p0000000000aa"
    const shared = path.join(root, "overlap.shared.docx")
    await writeFile(shared, await editBody(doc.file, body => {
      const mark = Array.from(body.getElementsByTagNameNS(W, "bookmarkStart")).find(row => row.getAttributeNS(W, "name") === `_dh_${doc.pid(PAYMENT)}`)!
      const copy = mark.cloneNode(false) as Element
      copy.setAttributeNS(W, "w:name", `_dh_${extra}`)
      copy.setAttributeNS(W, "w:id", "424242")
      const end = body.createElementNS(W, "w:bookmarkEnd")
      end.setAttributeNS(W, "w:id", "424242")
      mark.parentNode!.insertBefore(copy, mark)
      mark.parentNode!.appendChild(end)
    }))
    const payment = "Payment is due within ten days after the invoice date, and late payments accrue interest."
    const pointers = await applyAmendments({ docxPath: shared, outPath: doc.out, amendments: [
      amend("payment", doc.pid(PAYMENT), plain(PAYMENT), payment),
      amend("extra", extra, plain(PAYMENT), "Payment is due within sixty days after the invoice date, and late payments accrue simple interest."),
    ] })
    expect(pointers.applied).toEqual([{ id: "payment", pid: doc.pid(PAYMENT), mode: "code" }])
    expect(pointers.unapplied).toMatchObject([{ id: "extra", pid: extra, reason: "stale", detail: { current: payment, issue: overlap } }])
    expect((await texts(doc.out)).get(extra)).toBe(payment)
  })

  test("stale: a changed paragraph is never overwritten", async () => {
    const doc = await fixture("stale")
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("old", doc.pid(PAYMENT), "Payment is due within ninety days after the invoice date.", "Payment is due within ten days after the invoice date."),
    ] })
    expect(result).toEqual({ applied: [], unapplied: [{ id: "old", pid: doc.pid(PAYMENT), reason: "stale", detail: {
      current: plain(PAYMENT), target: "Payment is due within ten days after the invoice date.", reason: "Reflect the agreed position.",
      issue: "The paragraph no longer matches the text this amendment was written against.",
    } }] })
    expect(await readFile(doc.out)).toEqual(await readFile(doc.file))
  })

  test("target_deleted: a missing or duplicated pointer never falls back to matching text", async () => {
    const doc = await fixture("deleted")
    const duplicated = path.join(root, "deleted.duplicate.docx")
    // Copy the payment paragraph's pointer onto the steady paragraph.
    await writeFile(duplicated, await editBody(doc.file, body => {
      const marks = Array.from(body.getElementsByTagNameNS(W, "bookmarkStart"))
      const source = marks.find(mark => mark.getAttributeNS(W, "name") === `_dh_${doc.pid(PAYMENT)}`)!
      const target = marks.find(mark => mark.getAttributeNS(W, "name") === `_dh_${doc.pid(STEADY)}`)!
      target.setAttributeNS(W, "w:name", source.getAttributeNS(W, "name")!)
    }))
    const missing = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("gone", "p000000000000", plain(PAYMENT), "Payment is due within ten days after the invoice date."),
    ] })
    expect(missing).toEqual({ applied: [], unapplied: [{ id: "gone", pid: "p000000000000", reason: "target_deleted", detail: {
      target: "Payment is due within ten days after the invoice date.", reason: "Reflect the agreed position.", issue: "No paragraph carries this pointer.",
    } }] })
    expect((await texts(doc.out)).get(doc.pid(PAYMENT))).toBe(plain(PAYMENT))
    // Nothing applied: the written output is the input, byte for byte.
    expect(await readFile(doc.out)).toEqual(await readFile(doc.file))
    const ambiguous = await applyAmendments({ docxPath: duplicated, outPath: doc.out, amendments: [
      amend("twice", doc.pid(PAYMENT), plain(PAYMENT), "Payment is due within ten days after the invoice date."),
    ] })
    expect(ambiguous.unapplied).toMatchObject([{ id: "twice", pid: doc.pid(PAYMENT), reason: "target_deleted", detail: { issue: "More than one paragraph carries this pointer." } }])
    expect(ambiguous.applied).toEqual([])
    expect(await readFile(doc.out)).toEqual(await readFile(duplicated))
  })

  test("a numbered list paragraph amends by code and keeps its numbering, sibling items, table and untouched runs", async () => {
    const doc = await fixture("numbered")
    await writeFile(doc.file, await numbered(doc.file, [doc.pid(DELIVERY), doc.pid(PAYMENT)]))
    const updated = "The Supplier shall deliver the Goods within ten days of the order date."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("listed", doc.pid(DELIVERY), plain(DELIVERY), updated),
    ] })
    expect(result).toEqual({ applied: [{ id: "listed", pid: doc.pid(DELIVERY), mode: "code" }], unapplied: [] })
    expect((await texts(doc.out)).get(doc.pid(DELIVERY))).toBe(updated)
    const formatting = await measure(doc.file, doc.out)
    expect(formatting.changed).toBe(1)
    expect(formatting.preservation).toBe("ok")
    expect(formatting.outside).toBe(0)
    expect(formatting.word(doc.pid(DELIVERY), "Goods")).toEqual({ before: true, after: true })
    // One shared serializer, so the engine's self-closing tag spelling is not a change.
    const before = await markup(doc.file)
    const after = await markup(doc.out)
    expect(before.numPr).toHaveLength(2)
    expect(before.numPr[0]).toContain("numId")
    expect(after.numPr).toEqual(before.numPr)
    expect(before.tables).toHaveLength(1)
    expect(after.tables).toEqual(before.tables)
    expect(after.paragraphs.get(doc.pid(DELIVERY))!.properties).toEqual(before.paragraphs.get(doc.pid(DELIVERY))!.properties)
    for (const pid of before.paragraphs.keys()) if (pid !== doc.pid(DELIVERY)) expect(after.paragraphs.get(pid)!.xml).toBe(before.paragraphs.get(pid)!.xml)
    const runs = (side: typeof before) => side.paragraphs.get(doc.pid(DELIVERY))!.runs
    expect(runs(after).filter(run => !run.text.includes("ten")).map(run => run.xml))
      .toEqual(runs(before).filter(run => !run.text.includes("thirty")).map(run => run.xml))
    expect(after.parts["word/numbering.xml"]).toEqual(before.parts["word/numbering.xml"])
  })

  test("pointers survive apply and a save round trip; re-ensuring changes nothing", async () => {
    const doc = await fixture("roundtrip")
    const tagged = await readFile(doc.file)
    expect(await ensureParagraphPointers(doc.file)).toEqual(doc.pointers)
    expect(await readFile(doc.file)).toEqual(tagged)
    expect(Object.keys(doc.pointers)).toHaveLength(8)
    expect(Object.keys(doc.pointers).every(pid => /^p[0-9a-f]{12}$/.test(pid))).toBe(true)
    expect(strFromU8(unzipSync(tagged)["word/document.xml"]!)).toMatch(/^<\?xml /)
    const updated = "Notices must be in writing and sent by courier."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("notice", doc.pid(NOTICE), plain(NOTICE), updated),
    ] })
    expect(result.applied).toEqual([{ id: "notice", pid: doc.pid(NOTICE), mode: "code" }])
    const written = await readFile(doc.out)
    const pointers = await ensureParagraphPointers(doc.out)
    expect(await readFile(doc.out)).toEqual(written)
    expect(pointers).toEqual({ ...doc.pointers, [doc.pid(NOTICE)]: updated })
    const marks = bookmarks(written)
    expect(marks).toEqual({ starts: 8, ends: 8, orphans: 0 })
    const { docxodus } = await import("./docxodus")
    const session = (await docxodus()).openDocxSession(new Uint8Array(written), {})
    try {
      // Table paragraphs also report their covering table, row and cell.
      expect(Object.keys(pointers).every(pid => session.findByBookmark(`_dh_${pid}`)
        .filter(ref => !["tbl", "tr", "tc"].includes(ref.kind)).length === 1)).toBe(true)
      expect(bookmarks(session.save())).toEqual({ starts: 8, ends: 8, orphans: 0 })
    } finally { session.close() }
  })

  test("a copied pointer is re-minted on the later paragraph", async () => {
    const doc = await fixture("copied")
    await writeFile(doc.file, await editBody(doc.file, body => {
      const marks = Array.from(body.getElementsByTagNameNS(W, "bookmarkStart"))
      const source = marks.find(mark => mark.getAttributeNS(W, "name") === `_dh_${doc.pid(PAYMENT)}`)!
      const target = marks.find(mark => mark.getAttributeNS(W, "name") === `_dh_${doc.pid(STEADY)}`)!
      target.setAttributeNS(W, "w:name", source.getAttributeNS(W, "name")!)
    }))
    const pointers = await ensureParagraphPointers(doc.file)
    expect(pointers[doc.pid(PAYMENT)]).toBe(plain(PAYMENT))
    expect(pointers[doc.pid(STEADY)]).toBeUndefined()
    expect(Object.values(pointers)).toEqual(Object.values(doc.pointers))
    expect(bookmarks(await readFile(doc.file))).toEqual({ starts: 8, ends: 8, orphans: 0 })
  })

  test("rejects malformed batches with an actionable error", async () => {
    const doc = await fixture("invalid")
    const good = amend("a", doc.pid(PAYMENT), plain(PAYMENT), "Payment is due within ten days after the invoice date.")
    const run = (amendments: Amendment[]) => applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments })
    await expect(run([good, { ...good }])).rejects.toThrow("duplicate amendment id a")
    await expect(run([{ ...good, existing_sha256: "abc" }])).rejects.toThrow("0.existing_sha256")
    await expect(run([{ ...good, updated: "Line one.\nLine two." }])).rejects.toThrow("0.updated")
  })

  test("a long paragraph keeps the formatting of words between distant edits", async () => {
    const opening = Array.from({ length: 300 }, (_, index) => `alpha${index}`).join(" ")
    const closing = Array.from({ length: 300 }, (_, index) => `omega${index}`).join(" ")
    const doc = await single("long", `Opening clause ${opening} the **Defined Term** stays bold ${closing} closing clause.`)
    const before = (await texts(doc.file)).get(doc.pid)!
    const updated = before.replace("Opening clause", "Amended opening clause").replace("closing clause.", "amended closing clause.")
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend("long", doc.pid, before, updated)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    expect((await measure(doc.file, doc.out)).word(doc.pid, "Defined Term")).toEqual({ before: true, after: true })
  })

  test("resolving a highlighted open item leaves the answer unshaded", async () => {
    const doc = await single("resolved", "Delivery takes place at [open item: delivery address] on the agreed date.")
    const before = (await texts(doc.file)).get(doc.pid)!
    expect(await shaded(doc.file)).toEqual(["[open item: delivery address]"])
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("resolve", doc.pid, before, "Delivery takes place at the main warehouse on the agreed date."),
    ] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect(await shaded(doc.out)).toEqual([])
  })

  test("words an amendment writes just before a highlighted open item are not shaded as open", async () => {
    const doc = await single("before-open-item", "The Customer may reject within 30 days [open item: confirm the period] of delivery.")
    const before = (await texts(doc.file)).get(doc.pid)!
    const updated = "The Customer may reject within 30 days and in any event within 90 days [open item: confirm both periods] of delivery."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend("before", doc.pid, before, updated)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    expect(await shaded(doc.out)).toEqual(["[open item: confirm both periods]"])
  })

  test("an open item an amendment rewrites with the words before it stays in one shaded run", async () => {
    const doc = await single("rewritten-open-item", "The exclusion applies so far as it satisfies the requirement of reasonableness under clause 12 of this Agreement [open item: what is the evidence of bargaining strength to support reasonableness].")
    const before = (await texts(doc.file)).get(doc.pid)!
    const updated = "The exclusion applies so far as it satisfies the requirement of reasonableness [open item: confirm reasonableness qualification; what is the evidence of bargaining strength to support reasonableness]."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend("rewritten", doc.pid, before, updated)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    expect(await shaded(doc.out)).toEqual(["[open item: confirm reasonableness qualification; what is the evidence of bargaining strength to support reasonableness]"])
  })

  test("an open item an amendment closes with a later open item's bracket stays in one shaded run", async () => {
    const doc = await single("merged-open-items", "Defects may be notified within a reasonable time, and in any event within [open item: confirm the notice period and its start] after delivery [open item: confirm the remedy].")
    const before = (await texts(doc.file)).get(doc.pid)!
    const updated = "Defects may be notified within a reasonable time [open item: confirm the notice period and its start date]."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend("merged", doc.pid, before, updated)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    expect(await shaded(doc.out)).toEqual(["[open item: confirm the notice period and its start date]"])
  })

  // The apply step shades the markers an amendment wrote (highlightOpenItemsInDocx);
  // a drafting placeholder is shaded on its own, as an open item is.
  for (const [name, from, to, markers] of [
    ["placeholder-before-open-item", "The Supplier's liability is limited to [open item: what is needed — no cap figure is instructed].",
      "The Supplier's liability is limited to [£ (cap figure to be confirmed)] [open item: confirm the cap figure — no cap figure is instructed].",
      ["[£ (cap figure to be confirmed)]", "[open item: confirm the cap figure — no cap figure is instructed]"]],
    ["bracket-around-open-item", "Delivery is at [open item: confirm the address].",
      "Delivery is at [the main warehouse [open item: confirm whether the warehouse is right]].",
      ["[open item: confirm whether the warehouse is right]"]],
  ] as const) test(`words an amendment writes in an open item's run around a new marker are not shaded as open (${name})`, async () => {
    const doc = await single(name, from)
    const before = (await texts(doc.file)).get(doc.pid)!
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend(name, doc.pid, before, to)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    await writeFile(doc.out, highlightOpenItemsInDocx(new Uint8Array(await readFile(doc.out))))
    expect((await texts(doc.out)).get(doc.pid)).toBe(to)
    expect(await shaded(doc.out)).toEqual([...markers])
  })

  test("words an open item held that an amendment makes clause text are not shaded as open", async () => {
    const doc = await single("open-item-to-text", "The Buyer may reject the excess [open item: confirm that this regime is accepted, including that the Buyer may not reject a slight excess where unreasonable].")
    const before = (await texts(doc.file)).get(doc.pid)!
    const updated = "The Buyer may reject the excess, provided that the Buyer may not reject a slight excess where unreasonable [open item: confirm that this regime, including what counts as slight, is accepted]."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend("to-text", doc.pid, before, updated)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    expect(await shaded(doc.out)).toEqual(["[open item: confirm that this regime, including what counts as slight, is accepted]"])
  })

  test("a split lands with its pointers; an overlapping replacement after it is stale", async () => {
    const original = "The archive record is paper. Inspection is annual."
    const split = ["The archive record is paper.", "Inspection is annual."]
    const doc = await single("split-overlap", original)
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      { ...amend("split", doc.pid, original, ""), operation: "split", updated: split },
      amend("update", doc.pid, original, "The archive record is digital. Inspection is annual."),
    ] })
    expect(result.applied.map(row => [row.id, row.mode])).toEqual([["split", "code"]])
    expect(new Set(result.applied[0]!.result_pids).size).toBe(2)
    expect(result.unapplied).toMatchObject([{ id: "update", reason: "stale", detail: { current: split.join("\n\n") } }])
    expect([...await texts(doc.out)].map(([, text]) => text)).toEqual(split)
    expect(bookmarks(await readFile(doc.out))).toEqual({ starts: 2, ends: 2, orphans: 0 })
  })

  test("an overlapping deletion never deletes; a deletion alone applies once", async () => {
    const original = "The archive record is paper."
    const doc = await single("delete-overlap", original)
    const updated = "The archive record is digital."
    const deletion: Amendment = { ...amend("delete", doc.pid, original, ""), operation: "delete", sources: [{ source_id: "instruction", quote: "Remove the archive record paragraph." }] }
    const overlap = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [amend("update", doc.pid, original, updated), deletion] })
    expect(overlap.unapplied).toMatchObject([{ id: "delete", reason: "stale" }])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    const alone = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [deletion] })
    expect(alone.applied).toEqual([{ id: "delete", pid: doc.pid, mode: "code", result_pids: [] }])
    expect(await texts(doc.out)).toEqual(new Map())
    const repeated = await applyAmendments({ docxPath: doc.out, outPath: doc.out, amendments: [deletion] })
    expect(repeated.applied).toEqual([])
    expect(repeated.unapplied[0]).toMatchObject({ id: "delete", reason: "target_deleted" })
  })

  test("section breaks and cross-paragraph bookmarks contain structural edits while adjacent replacement applies", async () => {
    for (const kind of ["section", "range"] as const) {
      const doc = await fixture(`protected-${kind}`)
      await writeFile(doc.file, await editBody(doc.file, body => {
        const target = Array.from(body.getElementsByTagNameNS(W, "p")).find(node => paragraphPointers(node).some(row => row.pid === doc.pid(DELIVERY)))!
        if (kind === "section") {
          const properties = body.createElementNS(W, "w:pPr")
          properties.appendChild(body.createElementNS(W, "w:sectPr"))
          target.insertBefore(properties, target.firstChild)
        } else {
          const start = body.createElementNS(W, "w:bookmarkStart")
          start.setAttributeNS(W, "w:id", "9000")
          start.setAttributeNS(W, "w:name", "ExternalRange")
          target.appendChild(start)
          const end = body.createElementNS(W, "w:bookmarkEnd")
          end.setAttributeNS(W, "w:id", "9000")
          Array.from(body.getElementsByTagNameNS(W, "p")).find(node => paragraphPointers(node).some(row => row.pid === doc.pid(PAYMENT)))!.appendChild(end)
        }
      }))
      const payment = "Payment is due within sixty days after the invoice date."
      const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
        { ...amend("delete", doc.pid(DELIVERY), plain(DELIVERY), ""), operation: "delete", sources: [{ source_id: "instruction", quote: "Remove the delivery paragraph." }] },
        amend("payment", doc.pid(PAYMENT), plain(PAYMENT), payment),
      ] })
      expect(result.applied.map(row => row.id)).toEqual(["payment"])
      expect(result.unapplied[0]).toMatchObject({ id: "delete", reason: "apply_failed" })
      expect((await texts(doc.out)).get(doc.pid(DELIVERY))).toBe(plain(DELIVERY))
      expect((await texts(doc.out)).get(doc.pid(PAYMENT))).toBe(payment)
    }
  })

  test("a paragraph the engine cannot write fails alone and the rest of the batch applies", async () => {
    const doc = await fixture("engine")
    // A control character is not valid in document XML, so the engine refuses the edit.
    const broken = "The Supplier shall deliver the Goods\u0001 within thirty days of the order date."
    const payment = "Payment is due within sixty days after the invoice date."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      amend("broken", doc.pid(DELIVERY), plain(DELIVERY), broken),
      amend("payment", doc.pid(PAYMENT), plain(PAYMENT), payment),
    ] })
    expect(result).toEqual({ applied: [{ id: "payment", pid: doc.pid(PAYMENT), mode: "code" }], unapplied: [{ id: "broken", pid: doc.pid(DELIVERY), reason: "apply_failed", detail: {
      original: plain(DELIVERY), current: plain(DELIVERY), target: broken, reason: "Reflect the agreed position.", issue: "The document engine could not write this paragraph.",
    } }] })
    const after = await texts(doc.out)
    expect(after.get(doc.pid(DELIVERY))).toBe(plain(DELIVERY))
    expect(after.get(doc.pid(PAYMENT))).toBe(payment)
  })
})

describe("amendments beside live cross-references", () => {
  // A numbered agreement whose clause references render as REF fields.
  const AGREEMENT = ["# Services Agreement", "## 1. Definitions", "1.1 In this Agreement the terms in clause 2.1 apply.",
    "## 2. Services", "2.1 The Supplier shall provide the Services.", "2.2 Late payment carries interest, subject to clause 2.1."].join("\n\n")
  let root: string
  beforeAll(async () => { root = await mkdtemp(path.join(tmpdir(), "dh-amend-refs-")) })
  afterAll(async () => { await rm(root, { recursive: true, force: true }) })

  async function drafted(name: string) {
    const file = path.join(root, `${name}.docx`)
    await writeFile(file, (await renderMarkdownToDocx(AGREEMENT, undefined, undefined, { form: "agreement" })).bytes)
    const pointers = await ensureParagraphPointers(file)
    const pid = (start: string) => Object.entries(pointers).find(([, text]) => text.startsWith(start))![0]
    return { file, out: path.join(root, `${name}.out.docx`), pid, text: (start: string) => pointers[pid(start)]! }
  }

  // The REF fields of each paragraph, by its opening words.
  async function fields(file: string) {
    const body = new DOMParser().parseFromString(strFromU8(unzipSync(new Uint8Array(await readFile(file)))["word/document.xml"]!), "application/xml")
    return Object.fromEntries(Array.from(body.getElementsByTagNameNS(W, "p")).flatMap(paragraph => {
      const count = Array.from(paragraph.getElementsByTagNameNS(W, "instrText")).filter(node => /^\s*REF\s/.test(node.textContent ?? "")).length
      return count ? [[Array.from(paragraph.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join("").slice(0, 12), count]] : []
    }))
  }

  test("an edit the engine cannot write keeps its paragraph's cross-references live; with nothing applied the body is unchanged", async () => {
    const doc = await drafted("ref-failed")
    expect(await fields(doc.file)).toEqual({ "In this Agre": 1, "Late payment": 1 })
    // The edit rewrites the reference's number, and its control character makes the engine refuse it.
    const broken = amend("broken", doc.pid("Late payment"), doc.text("Late payment"), "Late payment carries interest, subject to clause 2.2\u0001.")
    const alone = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [broken] })
    expect(alone.unapplied).toMatchObject([{ id: "broken", reason: "apply_failed" }])
    expect(Buffer.from(await readFile(doc.out)).equals(Buffer.from(await readFile(doc.file)))).toBe(true)
    const beside = amend("beside", doc.pid("In this Agreement"), doc.text("In this Agreement"), "In this Agreement the defined terms in clause 2.1 apply.")
    const mixed = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [broken, beside] })
    expect([mixed.applied.map(row => row.id), mixed.unapplied.map(row => row.id)]).toEqual([["beside"], ["broken"]])
    expect((await texts(doc.out)).get(doc.pid("In this Agreement"))).toBe("In this Agreement the defined terms in clause 2.1 apply.")
    expect(await fields(doc.out)).toEqual({ "In this Agre": 1, "Late payment": 1 })
  })

  test("a deletion the engine cannot write leaves the references to its clause live", async () => {
    const doc = await drafted("ref-delete-failed")
    // A section break in the clause contains the deletion.
    await writeFile(doc.file, await editBody(doc.file, body => {
      const target = Array.from(body.getElementsByTagNameNS(W, "p")).find(node => paragraphPointers(node).some(row => row.pid === doc.pid("The Supplier")))!
      target.getElementsByTagNameNS(W, "pPr")[0]!.appendChild(body.createElementNS(W, "w:sectPr"))
    }))
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: [
      { ...amend("delete", doc.pid("The Supplier"), doc.text("The Supplier"), ""), operation: "delete", sources: [{ source_id: "instruction", quote: "Remove the services clause." }] },
    ] })
    expect(result.unapplied).toMatchObject([{ id: "delete", reason: "apply_failed" }])
    expect(Buffer.from(await readFile(doc.out)).equals(Buffer.from(await readFile(doc.file)))).toBe(true)
    expect(await fields(doc.out)).toEqual({ "In this Agre": 1, "Late payment": 1 })
  })
})

// Text of every run carrying a w:shd fill in the body.
async function shaded(file: string) {
  const body = new DOMParser().parseFromString(strFromU8(unzipSync(new Uint8Array(await readFile(file)))["word/document.xml"]!), "application/xml")
  return Array.from(body.getElementsByTagNameNS(W, "r")).filter(run => run.getElementsByTagNameNS(W, "shd").length)
    .map(run => Array.from(run.getElementsByTagNameNS(W, "t")).map(text => text.textContent).join(""))
}

function amend(id: string, pid: string, existing: string, updated: string, reason = "Reflect the agreed position."): Amendment {
  return { id, pid, existing_sha256: paragraphSha256(existing), updated, reason, sources: [] }
}

async function texts(file: string) {
  const opened = await openRevision(new Uint8Array(await readFile(file)))
  try {
    return new Map(opened.paragraphs.flatMap(paragraph => paragraphPointers(paragraph.node).map(row => [row.pid, paragraph.match.text] as const)))
  } finally { opened.session.close() }
}

async function editBody(file: string, edit: (body: Document) => void) {
  const parts = unzipSync(new Uint8Array(await readFile(file)))
  const body = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  edit(body)
  return zipSync({ ...parts, "word/document.xml": strToU8(new XMLSerializer().serializeToString(body)) })
}

// Adds a decimal numbering definition and puts the given paragraphs in that list.
async function numbered(file: string, pids: string[]) {
  const parts = unzipSync(new Uint8Array(await readFile(file)))
  const body = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  for (const node of Array.from(body.getElementsByTagNameNS(W, "p")).filter(node => paragraphPointers(node).some(row => pids.includes(row.pid)))) {
    const properties = Array.from(node.childNodes).find(child => (child as Element).localName === "pPr") as Element | undefined
      ?? node.insertBefore(body.createElementNS(W, "w:pPr"), node.firstChild)
    const numPr = properties.appendChild(body.createElementNS(W, "w:numPr"))
    numPr.appendChild(body.createElementNS(W, "w:ilvl")).setAttributeNS(W, "w:val", "0")
    numPr.appendChild(body.createElementNS(W, "w:numId")).setAttributeNS(W, "w:val", "1")
  }
  const types = strFromU8(parts["[Content_Types].xml"]!).replace("</Types>",
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
  const relationships = strFromU8(parts["word/_rels/document.xml.rels"]!).replace("</Relationships>",
    '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
  return zipSync({ ...parts, "[Content_Types].xml": strToU8(types), "word/_rels/document.xml.rels": strToU8(relationships),
    "word/numbering.xml": strToU8(numbering), "word/document.xml": strToU8(new XMLSerializer().serializeToString(body)) })
}

// Serialized markup of list properties, tables and each pointer paragraph.
async function markup(file: string) {
  const parts = unzipSync(new Uint8Array(await readFile(file)))
  const body = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  const serializer = new XMLSerializer()
  const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
  return {
    parts,
    numPr: all(body, "numPr").map(node => serializer.serializeToString(node)),
    tables: all(body, "tbl").map(node => serializer.serializeToString(node)),
    paragraphs: new Map(all(body, "p").flatMap(node => paragraphPointers(node).map(row => [row.pid, {
      xml: serializer.serializeToString(node),
      properties: all(node, "pPr").map(child => serializer.serializeToString(child)),
      runs: all(node, "r").map(run => ({ text: all(run, "t").map(text => text.textContent).join(""), xml: serializer.serializeToString(run) })),
    }] as const))),
  }
}

function bookmarks(bytes: Uint8Array) {
  const body = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const starts = Array.from(body.getElementsByTagNameNS(W, "bookmarkStart")).filter(mark => mark.getAttributeNS(W, "name")?.startsWith("_dh_"))
  const ids = new Set(starts.map(mark => mark.getAttributeNS(W, "id")))
  const ends = Array.from(body.getElementsByTagNameNS(W, "bookmarkEnd")).filter(mark => ids.has(mark.getAttributeNS(W, "id")))
  const ended = new Set(ends.map(mark => mark.getAttributeNS(W, "id")))
  return { starts: starts.length, ends: ends.length, orphans: [...ids].filter(id => !ended.has(id)).length }
}

// Per-character run properties of a paragraph, in text order.
function signatures(node: Element) {
  const serializer = new XMLSerializer()
  return Array.from(node.getElementsByTagNameNS(W, "r")).flatMap(run => {
    const properties = Array.from(run.childNodes).find(child => child.nodeType === 1 && (child as Element).localName === "rPr")
    const signature = properties ? serializer.serializeToString(properties) : ""
    return Array.from(run.getElementsByTagNameNS(W, "t")).flatMap(text => Array.from(text.textContent ?? "").map(char => ({ char, signature })))
  })
}

// Counts characters whose run properties changed outside the outer changed
// span of each edited paragraph, and checks unrelated structure.
async function measure(input: string, output: string) {
  const a = await openRevision(new Uint8Array(await readFile(input)))
  const b = await openRevision(new Uint8Array(await readFile(output)))
  try {
    const changed = new Set(a.paragraphs.flatMap((paragraph, index) => paragraph.match.text !== b.paragraphs[index]!.match.text ? [index] : []))
    const preservation = (() => {
      try {
        verifyPreservation(a, b, changed)
        return "ok"
      } catch (error) { return error instanceof Error ? error.message : String(error) }
    })()
    const outside = [...changed].reduce((total, index) => {
      const x = signatures(a.paragraphs[index]!.node)
      const y = signatures(b.paragraphs[index]!.node)
      const head = x.findIndex((row, k) => row.char !== y[k]?.char)
      const start = head < 0 ? Math.min(x.length, y.length) : head
      const tail = Array.from({ length: Math.min(x.length, y.length) - start })
        .findIndex((_, k) => x[x.length - 1 - k]!.char !== y[y.length - 1 - k]!.char)
      const end = tail < 0 ? Math.min(x.length, y.length) - start : tail
      const prefix = x.slice(0, start).filter((row, k) => row.signature !== y[k]!.signature).length
      const suffix = Array.from({ length: end }).filter((_, k) => x[x.length - 1 - k]!.signature !== y[y.length - 1 - k]!.signature).length
      return total + prefix + suffix
    }, 0)
    const bold = (opened: typeof a, pid: string, word: string) => {
      const paragraph = opened.paragraphs.find(row => paragraphPointers(row.node).some(pointer => pointer.pid === pid))!
      const chars = signatures(paragraph.node)
      const at = chars.map(row => row.char).join("").indexOf(word)
      return chars.slice(at, at + word.length).every(row => /<w:(b|i)\b/.test(row.signature))
    }
    return {
      changed: changed.size,
      outside,
      preservation,
      word: (pid: string, word: string) => ({ before: bold(a, pid, word), after: bold(b, pid, word) }),
    }
  } finally {
    a.session.close()
    b.session.close()
  }
}
