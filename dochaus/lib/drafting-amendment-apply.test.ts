import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { openRevision, verifyPreservation } from "./docx-amendments"
import { applyAmendments, createAmendmentMerge, type AmendmentMergeInput } from "./drafting-amendment-apply"
import { ensureParagraphPointers, paragraphPointers, paragraphSha256, type Amendment } from "./drafting-amendment"
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
    const merges: AmendmentMergeInput[] = []
    const delivery = "The Vendor shall deliver the Goods within ten days of the order date."
    const payment = "Payment is due within sixty days after the invoice date."
    const notice = "Notices must be in writing and sent by post or email."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(merges, () => ""), amendments: [
      amend("a1", doc.pid(DELIVERY), plain(DELIVERY), delivery),
      amend("a2", doc.pid(PAYMENT), plain(PAYMENT), payment),
      amend("a3", doc.pid(NOTICE), plain(NOTICE), notice),
    ] })
    expect(result).toEqual({ applied: [
      { id: "a1", pid: doc.pid(DELIVERY), mode: "code" },
      { id: "a2", pid: doc.pid(PAYMENT), mode: "code" },
      { id: "a3", pid: doc.pid(NOTICE), mode: "code" },
    ], unapplied: [] })
    expect(merges).toHaveLength(0)
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

  test("a second amendment on the same paragraph merges against original and current", async () => {
    const doc = await fixture("pair")
    const pid = doc.pid(DELIVERY)
    const first = "The Supplier shall deliver the Goods within ten days of the order date."
    const second = "The Supplier shall deliver the Goods within thirty days of the delivery request."
    const merged = "The Supplier shall deliver the Goods within ten days of the delivery request."
    const merges: AmendmentMergeInput[] = []
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(merges, () => merged), amendments: [
      amend("first", pid, plain(DELIVERY), first),
      amend("second", pid, plain(DELIVERY), second, "Tie timing to the delivery request."),
    ] })
    expect(result.applied).toEqual([{ id: "first", pid, mode: "code" }, { id: "second", pid, mode: "merged" }])
    expect(merges).toHaveLength(1)
    expect(merges[0]).toMatchObject({ pid, original: plain(DELIVERY), current: first })
    expect(merges[0]!.target.id).toBe("second")
    expect((await texts(doc.out)).get(pid)).toBe(merged)
    expect((await measure(doc.file, doc.out)).word(pid, "Goods")).toEqual({ before: true, after: true })
  })

  test("already_applied only on an untouched paragraph; later amendments on a changed paragraph always merge", async () => {
    const doc = await fixture("already")
    const pid = doc.pid(PAYMENT)
    const both = "Payment is due within thirty days after the invoice date, and late payments accrue simple interest."
    const merges: AmendmentMergeInput[] = []
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(merges, input => input.current), amendments: [
      amend("both", pid, plain(PAYMENT), both),
      amend("repeat", pid, plain(PAYMENT), both),
      amend("subset", pid, plain(PAYMENT), "Payment is due within thirty days after the invoice date, and late payments accrue interest."),
      amend("steady", doc.pid(STEADY), plain(STEADY), plain(STEADY)),
    ] })
    expect(result.applied).toEqual([
      { id: "both", pid, mode: "code" },
      { id: "repeat", pid, mode: "merged" },
      { id: "subset", pid, mode: "merged" },
      { id: "steady", pid: doc.pid(STEADY), mode: "already_applied" },
    ])
    // Even the exact repeat of a landed change is reviewed by the merger.
    expect(merges.map(row => row.target.id)).toEqual(["repeat", "subset"])
    expect((await texts(doc.out)).get(pid)).toBe(both)
  })

  test("stale: a changed paragraph is never overwritten and the merger is not asked", async () => {
    const doc = await fixture("stale")
    const merges: AmendmentMergeInput[] = []
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(merges, () => "unused"), amendments: [
      amend("old", doc.pid(PAYMENT), "Payment is due within ninety days after the invoice date.", "Payment is due within ten days after the invoice date."),
    ] })
    expect(result).toEqual({ applied: [], unapplied: [{ id: "old", pid: doc.pid(PAYMENT), reason: "stale", detail: {
      current: plain(PAYMENT), target: "Payment is due within ten days after the invoice date.", reason: "Reflect the agreed position.",
      issue: "The paragraph no longer matches the text this amendment was written against.",
    } }] })
    expect(merges).toHaveLength(0)
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
    const missing = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "unused", amendments: [
      amend("gone", "p000000000000", plain(PAYMENT), "Payment is due within ten days after the invoice date."),
    ] })
    expect(missing).toEqual({ applied: [], unapplied: [{ id: "gone", pid: "p000000000000", reason: "target_deleted", detail: {
      target: "Payment is due within ten days after the invoice date.", reason: "Reflect the agreed position.", issue: "No paragraph carries this pointer.",
    } }] })
    expect((await texts(doc.out)).get(doc.pid(PAYMENT))).toBe(plain(PAYMENT))
    const ambiguous = await applyAmendments({ docxPath: duplicated, outPath: doc.out, merge: async () => "unused", amendments: [
      amend("twice", doc.pid(PAYMENT), plain(PAYMENT), "Payment is due within ten days after the invoice date."),
    ] })
    expect(ambiguous.unapplied).toMatchObject([{ id: "twice", pid: doc.pid(PAYMENT), reason: "target_deleted", detail: { issue: "More than one paragraph carries this pointer." } }])
  })

  test("two pointers left on one paragraph settle as one sequence", async () => {
    const doc = await fixture("twopointers")
    const extra = "p0000000000aa"
    const shared = path.join(root, "twopointers.shared.docx")
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
    const first = "Payment is due within ten days after the invoice date, and late payments accrue interest."
    const second = "Payment is due within sixty days after the invoice date, and late payments accrue simple interest."
    const merged = "Payment is due within ten days after the invoice date, and late payments accrue simple interest."
    const result = await applyAmendments({ docxPath: shared, outPath: doc.out, merge: async () => merged, amendments: [
      amend("first", doc.pid(PAYMENT), plain(PAYMENT), first),
      amend("second", extra, plain(PAYMENT), second),
    ] })
    expect(result.applied).toEqual([{ id: "first", pid: doc.pid(PAYMENT), mode: "code" }, { id: "second", pid: extra, mode: "merged" }])
    expect((await texts(doc.out)).get(extra)).toBe(merged)
  })

  test("merge_failed after a re-ask and a fresh attempt; a later attempt that keeps the landed change is accepted", async () => {
    const doc = await fixture("failed")
    const pid = doc.pid(DELIVERY)
    const first = "The Supplier shall deliver the Goods within ten days of the order date."
    const second = "The Supplier shall deliver the Goods within thirty days of the delivery request."
    const batch = [amend("first", pid, plain(DELIVERY), first), amend("second", pid, plain(DELIVERY), second)]
    const empty: AmendmentMergeInput[] = []
    const failed = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(empty, () => "  "), amendments: batch })
    expect(failed.applied).toEqual([{ id: "first", pid, mode: "code" }])
    expect(failed.unapplied).toEqual([{ id: "second", pid, reason: "merge_failed", detail: {
      original: plain(DELIVERY), current: first, target: second, reason: "Reflect the agreed position.",
      issue: "No paragraph text was found; return the full paragraph text inside <paragraph></paragraph> tags.",
    } }])
    expect(empty).toHaveLength(3)
    expect(empty[0]!.issue).toBeUndefined()
    expect(empty[1]!.issue).toContain("<paragraph>")
    expect(empty[2]!.issue).toContain("<paragraph>")
    expect(empty.map(row => row.previous)).toEqual([undefined, undefined, undefined])
    expect((await texts(doc.out)).get(pid)).toBe(first)

    // Returning the target text would revert the landed first change.
    const reverting: AmendmentMergeInput[] = []
    const blocked = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(reverting, input => input.target.updated), amendments: batch })
    expect(blocked.unapplied).toMatchObject([{ id: "second", pid, reason: "merge_failed", detail: { issue: expect.stringContaining("reverted") } }])
    expect(reverting).toHaveLength(3)
    // The re-ask shows the unusable answer; the last attempt starts fresh with only the reason.
    expect(reverting[1]).toMatchObject({ previous: second, issue: expect.stringContaining("reverted") })
    expect(reverting[2]!.previous).toBeUndefined()
    expect(reverting[2]!.issue).toContain("reverted")
    expect((await texts(doc.out)).get(pid)).toBe(first)

    const recovered: AmendmentMergeInput[] = []
    const merged = "The Supplier shall deliver the Goods within ten days of the delivery request."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch,
      merge: recordMerge(recovered, input => input.issue && !input.previous ? merged : input.target.updated) })
    expect(result.applied).toEqual([{ id: "first", pid, mode: "code" }, { id: "second", pid, mode: "merged" }])
    expect(recovered).toHaveLength(3)
    expect((await texts(doc.out)).get(pid)).toBe(merged)
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
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "unused", amendments: [
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
    const run = (amendments: Amendment[]) => applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "", amendments })
    await expect(run([good, { ...good }])).rejects.toThrow("duplicate amendment id a")
    await expect(run([{ ...good, existing_sha256: "abc" }])).rejects.toThrow("0.existing_sha256")
    await expect(run([{ ...good, updated: "Line one.\nLine two." }])).rejects.toThrow("0.updated")
  })

  test("createAmendmentMerge sends original, current, target and reason and reads only the tagged paragraph", async () => {
    const prompts: string[] = []
    const replies = ["Here it is:\n<paragraph> Merged A &lt; B &amp; C text. </paragraph>", "Merged paragraph text without tags."]
    const merge = createAmendmentMerge(async prompt => {
      prompts.push(prompt)
      return replies[prompts.length - 1]!
    })
    const target = amend("t", "p0123456789ab", "Old text.", "Target <b> & text.", "Explains the change.")
    const input = { pid: target.pid, original: "Old text.", current: "Current text.", target, previous: "Bad answer.", issue: "It was empty." }
    expect(await merge(input)).toBe("Merged A < B & C text.")
    expect(prompts[0]).toContain("<original>Old text.</original>")
    expect(prompts[0]).toContain("<current>Current text.</current>")
    expect(prompts[0]).toContain("<target>Target &lt;b&gt; &amp; text.</target>")
    expect(prompts[0]).toContain("<reason>Explains the change.</reason>")
    expect(prompts[0]).toContain("<previous_answer>Bad answer.</previous_answer>")
    expect(prompts[0]).toContain("It was empty.")
    // Text outside the tags is never treated as the paragraph.
    expect(await merge(input)).toBe("")
  })

  test("a lazy merge reply missing one piece of a multi-part change is refused", async () => {
    const doc = await single("multipart", "The Supplier shall deliver the Goods within thirty days of the order date.")
    const original = "The Supplier shall deliver the Goods within thirty days of the order date."
    const first = "The Supplier shall deliver the Goods within ten days of the order date."
    const second = "The Vendor shall deliver the Goods within thirty days of the delivery request."
    const calls: AmendmentMergeInput[] = []
    const batch = [amend("first", doc.pid, original, first), amend("second", doc.pid, original, second)]
    const lazy = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch, merge: recordMerge(calls, input => input.current) })
    expect(lazy.unapplied).toMatchObject([{ id: "second", reason: "merge_failed", detail: { original, current: first, target: second } }])
    expect(calls).toHaveLength(3)
    expect((await texts(doc.out)).get(doc.pid)).toBe(first)
    // Only part of the change: the rename landed but the second piece did not.
    const partial = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch,
      merge: async () => "The Vendor shall deliver the Goods within ten days of the order date." })
    expect(partial.unapplied.map(row => row.reason)).toEqual(["merge_failed"])
    const full = "The Vendor shall deliver the Goods within ten days of the delivery request."
    const complete = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch, merge: async () => full })
    expect(complete.applied.map(row => row.mode)).toEqual(["code", "merged"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(full)
  })

  test("an inserted word is matched as a whole word, never inside a longer word", async () => {
    const original = "The Supplier shall be liable on notice for late delivery of the Goods."
    const doc = await single("negation", original)
    const first = "The Supplier shall be liable on written notice for late delivery of the Goods."
    const second = "The Supplier shall not be liable on notice for late delivery of the Goods."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async input => input.current, amendments: [
      amend("first", doc.pid, original, first), amend("second", doc.pid, original, second),
    ] })
    expect(result.unapplied.map(row => row.reason)).toEqual(["merge_failed"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(first)
  })

  test("a merge that carries an earlier rename into the requested words is accepted", async () => {
    const original = "The Buyer shall pay the Price within ten days."
    const doc = await single("rename", original)
    const calls: AmendmentMergeInput[] = []
    const adapted = "The Purchaser shall pay the Price and the Purchaser's costs within ten days."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(calls, () => adapted), amendments: [
      amend("rename", doc.pid, original, "The Purchaser shall pay the Price within ten days."),
      amend("costs", doc.pid, original, "The Buyer shall pay the Price and the Buyer's costs within ten days."),
    ] })
    expect(result.applied.map(row => row.mode)).toEqual(["code", "merged"])
    expect(calls).toHaveLength(1)
    expect((await texts(doc.out)).get(doc.pid)).toBe(adapted)
  })

  test("an untagged merge reply is never written; a later tagged attempt is", async () => {
    const original = "Payment is due within sixty days after the invoice date."
    const doc = await single("untagged", original)
    const batch = [
      amend("first", doc.pid, original, "Payment is due within thirty days after the invoice date."),
      amend("second", doc.pid, original, "Payment is due within sixty days after receipt of the invoice."),
    ]
    const preamble = createAmendmentMerge(async () => "Merged paragraph: Payment is due within thirty days after receipt of the invoice.")
    const refused = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch, merge: preamble })
    expect(refused.unapplied.map(row => row.reason)).toEqual(["merge_failed"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(batch[0]!.updated)
    const prompts: string[] = []
    const tagged = createAmendmentMerge(async prompt => {
      prompts.push(prompt)
      return prompts.length < 3 ? "Merged paragraph: no tags." : "<paragraph>Payment is due within thirty days after receipt of the invoice.</paragraph>"
    })
    const recovered = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch, merge: tagged })
    expect(recovered.applied.map(row => row.mode)).toEqual(["code", "merged"])
    expect(prompts).toHaveLength(3)
    expect(prompts[2]).not.toContain("<previous_answer>")
    expect((await texts(doc.out)).get(doc.pid)).toBe("Payment is due within thirty days after receipt of the invoice.")
  })

  test("on a changed paragraph, a stale hash or an already-present text still goes to merge", async () => {
    const doc = await fixture("touched")
    const pid = doc.pid(DELIVERY)
    const first = "The Supplier shall deliver the Goods within ten days of the order date."
    const calls: AmendmentMergeInput[] = []
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: recordMerge(calls, input => input.current), amendments: [
      amend("first", pid, plain(DELIVERY), first),
      amend("unknown", pid, "A text this paragraph never had.", first),
      amend("present", pid, first, first),
    ] })
    expect(result.applied).toEqual([{ id: "first", pid, mode: "code" }, { id: "unknown", pid, mode: "merged" }, { id: "present", pid, mode: "merged" }])
    // Matching no state, the amendment is merged against the paragraph as it was before the batch.
    expect(calls.map(row => [row.target.id, row.original, row.current])).toEqual([["unknown", plain(DELIVERY), first], ["present", first, first]])
  })

  test("a merge reply with line breaks is written as one paragraph", async () => {
    const doc = await fixture("breaks")
    const pid = doc.pid(PAYMENT)
    const first = "Payment is due within ten days after the invoice date, and late payments accrue interest."
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out,
      merge: async () => "Payment is due within ten days after the invoice date,\nand late payments accrue simple interest.", amendments: [
        amend("first", pid, plain(PAYMENT), first),
        amend("second", pid, plain(PAYMENT), "Payment is due within sixty days after the invoice date, and late payments accrue simple interest."),
      ] })
    expect(result.applied.map(row => row.mode)).toEqual(["code", "merged"])
    expect((await texts(doc.out)).get(pid)).toBe("Payment is due within ten days after the invoice date, and late payments accrue simple interest.")
  })

  test("a long paragraph keeps the formatting of words between distant edits", async () => {
    const opening = Array.from({ length: 300 }, (_, index) => `alpha${index}`).join(" ")
    const closing = Array.from({ length: 300 }, (_, index) => `omega${index}`).join(" ")
    const doc = await single("long", `Opening clause ${opening} the **Defined Term** stays bold ${closing} closing clause.`)
    const before = (await texts(doc.file)).get(doc.pid)!
    const updated = before.replace("Opening clause", "Amended opening clause").replace("closing clause.", "amended closing clause.")
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "", amendments: [amend("long", doc.pid, before, updated)] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    expect((await measure(doc.file, doc.out)).word(doc.pid, "Defined Term")).toEqual({ before: true, after: true })
  })

  test("resolving a highlighted open item leaves the answer unshaded", async () => {
    const doc = await single("resolved", "Delivery takes place at [open item: delivery address] on the agreed date.")
    const before = (await texts(doc.file)).get(doc.pid)!
    expect(await shaded(doc.file)).toEqual(["[open item: delivery address]"])
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "", amendments: [
      amend("resolve", doc.pid, before, "Delivery takes place at the main warehouse on the agreed date."),
    ] })
    expect(result.applied.map(row => row.mode)).toEqual(["code"])
    expect(await shaded(doc.out)).toEqual([])
  })

  test("overlapping replacement preserves an existing split and refuses a collapsed merge", async () => {
    const original = "The archive record is paper. Inspection is annual."
    const split = ["The archive record is paper.", "Inspection is annual."]
    const final = ["The archive record is digital.", "Inspection is annual."]
    const doc = await single("split-overlap", original)
    const batch: Amendment[] = [
      { ...amend("split", doc.pid, original, ""), operation: "split", updated: split },
      amend("update", doc.pid, original, "The archive record is digital. Inspection is annual."),
    ]
    let calls = 0
    const collapsed = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch,
      merge: async () => { calls++; return final.join(" ") } })
    expect(calls).toBe(3)
    expect(collapsed.unapplied[0]).toMatchObject({ id: "update", reason: "merge_failed", detail: { issue: expect.stringContaining("collapsed") } })
    expect([...await texts(doc.out)].map(([, text]) => text)).toEqual(split)
    const merged = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch,
      merge: async input => { expect(input.current_paragraphs).toEqual(split); return final } })
    expect(merged.unapplied).toEqual([])
    expect(merged.applied.map(row => row.mode)).toEqual(["code", "merged"])
    expect(merged.applied[0]!.result_pids).toEqual(merged.applied[1]!.result_pids)
    expect(new Set(merged.applied[0]!.result_pids).size).toBe(2)
    expect([...await texts(doc.out)].map(([, text]) => text)).toEqual(final)
    expect(bookmarks(await readFile(doc.out))).toEqual({ starts: 2, ends: 2, orphans: 0 })
  })

  test("overlapping deletion requires a deliberate typed result; blank replies never delete", async () => {
    const original = "The archive record is paper."
    const doc = await single("delete-overlap", original)
    const updated = "The archive record is digital."
    const batch: Amendment[] = [amend("update", doc.pid, original, updated), {
      ...amend("delete", doc.pid, original, ""), operation: "delete", sources: [{ source_id: "instruction", quote: "Remove the archive record paragraph." }],
    }]
    const blank = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch, merge: async () => "" })
    expect(blank.unapplied[0]).toMatchObject({ id: "delete", reason: "merge_failed" })
    expect((await texts(doc.out)).get(doc.pid)).toBe(updated)
    const explicit = await applyAmendments({ docxPath: doc.file, outPath: doc.out, amendments: batch, merge: async () => ({ operation: "delete" }) })
    expect(explicit.unapplied).toEqual([])
    expect(explicit.applied.map(row => row.result_pids)).toEqual([[], []])
    expect(await texts(doc.out)).toEqual(new Map())
    const repeated = await applyAmendments({ docxPath: doc.out, outPath: doc.out, amendments: [batch[1]!], merge: async () => { throw new Error("not called") } })
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
      const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "", amendments: [
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
    const result = await applyAmendments({ docxPath: doc.file, outPath: doc.out, merge: async () => "", amendments: [
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

// Text of every run carrying a w:shd fill in the body.
async function shaded(file: string) {
  const body = new DOMParser().parseFromString(strFromU8(unzipSync(new Uint8Array(await readFile(file)))["word/document.xml"]!), "application/xml")
  return Array.from(body.getElementsByTagNameNS(W, "r")).filter(run => run.getElementsByTagNameNS(W, "shd").length)
    .map(run => Array.from(run.getElementsByTagNameNS(W, "t")).map(text => text.textContent).join(""))
}

function amend(id: string, pid: string, existing: string, updated: string, reason = "Reflect the agreed position."): Amendment {
  return { id, pid, existing_sha256: paragraphSha256(existing), updated, reason, sources: [] }
}

function recordMerge(calls: AmendmentMergeInput[], answer: (input: AmendmentMergeInput) => string) {
  return async (input: AmendmentMergeInput) => {
    calls.push(input)
    return answer(input)
  }
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
