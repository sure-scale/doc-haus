import { expect, test } from "bun:test"
import { WordChatHost, type WordChatContext, type WordChatWord, type WordChatApplyInput } from "./word-chat-host"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { WordChatAuthoredAmendment } from "../../../dochaus/lib/word-chat-amendment-contract"
import { wordChanges } from "../../../dochaus/lib/word-changes"

const WORD_XML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PACKAGE_XML = "http://schemas.microsoft.com/office/2006/xmlPackage"
// Minimal reviewed synthetic shape of actual Mac Word Range.getOoxml(): even
// one ordinary word exports Flat OPC with a final body-level section layout.
const escapeXml = (text: string) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
const flatOpc = (content: string, extra = "") => `<pkg:package xmlns:pkg="${PACKAGE_XML}" xmlns:w="${WORD_XML}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body>${content}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part>${extra}</pkg:package>`

/** A text-position Office host model: selections and paragraph ranges are real
 * offsets, searches return every occurrence, and writes execute only on sync. */
function office(paragraphs = ["A clause with a clear obligation.", "Another clause."]) {
  type NativeParagraph = { id: string; text: string; style?: string; styleBuiltIn?: string; outlineLevel?: number;
    list?: { id: number; level: number; label: string }; inTable?: boolean }
  const state = {
    paragraphs: paragraphs.map((text, i): NativeParagraph => ({ id: `paragraph-${i}`, text })),
    selection: [0, paragraphs[0]!.length], saved: false, tracking: "Off", modeChanges: [] as string[],
    writes: [] as Array<{ before: string; text: string; tracking: string; start: number; end: number }>,
    paragraphXml: new Map<string, string>(), comments: [] as Array<[number, number]>, revisions: [] as Array<[number, number]>, xml: undefined as string | undefined,
    selected: 0, highlights: 0, outside: false, annotationCounter: 0,
    annotations: new Map<string, unknown>(), handlers: new Map<string, Set<(args: never) => void>>(),
    throwAfterWrite: false, corruptWrite: false, afterSync: undefined as undefined | (() => void),
    driftInsertedRanges: false, renewImportedIds: false, omitNewRevisions: false, authorCalls: 0, duringAuthor: undefined as undefined | (() => void | Promise<void>),
    nativeRevisions: [] as Array<{ start: number; end: number; author: string; type: string; text: string }>,
    paragraphLoads: [] as Array<{ id: string; properties: string }>, metadataUnavailable: false, metadataFailures: 0,
  }
  const text = () => state.paragraphs.map((p) => p.text).join("\r")
  const position = (id: string) => {
    let start = 0
    for (const p of state.paragraphs) {
      if (p.id === id) return { p, start, end: start + p.text.length }
      start += p.text.length + 1
    }
    throw new Error("ItemNotFound")
  }
  const intersects = (a: number, b: number, c: number, d: number) => a === b ? c <= a && a <= d : c === d ? a <= c && c <= b : a < d && c < b
  const queue: Array<() => void> = []
  const event = (name: string) => ({ add(callback: (args: never) => void) {
    const handlers = state.handlers.get(name) ?? new Set()
    state.handlers.set(name, handlers)
    handlers.add(callback)
    return { context, remove() { handlers.delete(callback) } }
  } })
  const collection = <T>(items: T[]) => ({ items, load() {} })
  const ranges = new Set<Range>()
  class Range {
    constructor(public start: number, public end: number) { ranges.add(this) }
    get text() { return text().slice(this.start, this.end) }
    get paragraphs() {
      return collection(state.paragraphs.filter((p) => {
        const at = position(p.id)
        return this.start === this.end ? at.start <= this.start && this.start <= at.end : at.start < this.end && this.start <= at.end
      }).map(paragraph))
    }
    load() {}
    getRange(location = "Whole") {
      if (location === "Start") return new Range(this.start, this.start)
      if (location === "End") return new Range(this.end, this.end)
      return new Range(this.start, this.end)
    }
    expandTo(other: Range) { return new Range(Math.min(this.start, other.start), Math.max(this.end, other.end)) }
    compareLocationWith(other: Range) {
      return { value: state.outside ? "Unrelated" : this.start >= other.start && this.end <= other.end ? "Inside" : "Unrelated" }
    }
    search(query: string) {
      // Decode only the literal/control Word Find forms used by the adapter.
      const literal = query.replace(/\^([\^plt])/g, (_, code: string) => ({ "^": "^", p: "\r", l: "\u000b", t: "\t" })[code]!)
      const matches: Range[] = []
      let at = this.text.indexOf(literal)
      while (at >= 0) { matches.push(new Range(this.start + at, this.start + at + literal.length)); at = this.text.indexOf(literal, at + 1) }
      return collection(matches)
    }
    insertOoxml(xml: string) {
      const result = new Range(this.start, this.end)
      queue.push(() => {
        const doc = new DOMParser().parseFromString(xml, "application/xml")
        const blocks = Array.from(doc.getElementsByTagNameNS(WORD_XML, "body")[0]!.childNodes).filter((n): n is Element => n.nodeType === 1 && (n as Element).localName === "p")
        const owners = state.paragraphs.filter(p => { const at = position(p.id); return this.start <= at.start && at.end <= this.end })
        if (owners.length !== blocks.length) throw new Error("Cannot replace paragraph structure")
        const start = this.start, end = this.end, before = this.text
        const current = (node: Node): string => {
          const element = node as Element
          if (element.nodeType !== 1 || element.localName === "del") return ""
          if (element.localName === "t") return element.textContent || ""
          return Array.from(element.childNodes).map(current).join("")
        }
        const texts = blocks.map(current), replacement = texts.join("\r")
        for (const [index, owner] of owners.entries()) {
          owner.text = state.corruptWrite && index === 0 ? "WRONG" : texts[index]!
          if (state.renewImportedIds) owner.id = `imported-${owner.id}`
          state.paragraphXml.set(owner.id, new XMLSerializer().serializeToString(blocks[index]!))
        }
        const delta = replacement.length - (end - start)
        state.writes.push({ before, text: replacement, tracking: state.tracking, start, end })
        for (const kind of ["comments", "revisions"] as const) state[kind] = state[kind].map(([a, b]) => a >= end ? [a + delta, b + delta] : [a, b])
        for (const revision of state.nativeRevisions) if (revision.start >= end) { revision.start += delta; revision.end += delta }
        if (!state.omitNewRevisions) for (const [index, block] of blocks.entries()) {
          let at = position(owners[index]!.id).start
          const visit = (node: Node) => {
            const element = node as Element
            if (element.nodeType !== 1) return
            if (["ins", "del"].includes(element.localName)) {
              const added = element.localName === "ins"
              const wording = Array.from(element.getElementsByTagNameNS(WORD_XML, added ? "t" : "delText")).map(n => n.textContent).join("")
              state.nativeRevisions.push({ start: at, end: at + (added ? wording.length : 0), author: element.getAttributeNS(WORD_XML, "author") || "", type: added ? "Added" : "Deleted", text: wording })
              if (added) at += wording.length
              return
            }
            if (element.localName === "t") { at += element.textContent?.length || 0; return }
            for (const child of Array.from(element.childNodes)) visit(child)
          }
          visit(block)
        }
        state.saved = false
        if (state.driftInsertedRanges) { result.start = 0; result.end = text().length }
      })
      return result
    }
    getOoxml() { return { value: !state.writes.length && state.xml !== undefined ? state.xml : flatOpc(this.paragraphs.items.map(p => state.paragraphXml.get(p.uniqueLocalId) ?? `<w:p><w:r><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r></w:p>`).join("")) } }
    getReviewedText() { return { value: this.text } }
    getComments() { return collection(state.comments.filter(([start, end]) => intersects(this.start, this.end, start, end))) }
    getTrackedChanges() {
      return collection([
        ...state.revisions.filter(([start, end]) => intersects(this.start, this.end, start, end)).map(([start, end]) => ({ author: "Earlier reviewer", type: "Added", text: text().slice(start, end) })),
        ...state.nativeRevisions.filter(revision => intersects(this.start, this.end, revision.start, revision.end)).sort((a, b) => a.start - b.start),
      ])
    }
    select() { state.selected++ }
    highlight() { state.highlights++ }
    removeHighlight() { state.highlights-- }
    track() { return this }
    untrack() { return this }
  }
  const paragraph = (p: NativeParagraph) => ({
    get text() { return p.text }, get uniqueLocalId() { return p.id },
    get style() { return p.style ?? "Normal" }, get styleBuiltIn() { return p.styleBuiltIn ?? "Normal" },
    get outlineLevel() { return p.outlineLevel ?? 10 }, get isListItem() { return Boolean(p.list) },
    get tableNestingLevel() { return p.inTable ? 1 : 0 },
    get list() { if (!p.list) throw new Error("ItemNotFound"); return { id: p.list.id, load() {} } },
    get listItem() { if (!p.list) throw new Error("ItemNotFound"); return { level: p.list.level, listString: p.list.label, load() {} } },
    load(properties: string) {
      state.paragraphLoads.push({ id: p.id, properties })
      if (state.metadataUnavailable && properties.includes("style")) {
        state.metadataFailures++
        queue.push(() => { throw new Error("InvalidArgument: metadata unavailable") })
      }
    },
    getRange(location = "Whole") {
      const at = position(p.id)
      if (location === "Start") return new Range(at.start, at.start)
      if (location === "End") return new Range(at.end, at.end)
      return new Range(at.start, at.end)
    },
    insertAnnotations(set: unknown) {
      const id = `annotation-${++state.annotationCounter}`
      state.annotations.set(id, set)
      return { value: [id] }
    },
  })
  const context = {
    document: {
      get saved() { return state.saved },
      get changeTrackingMode() { return state.tracking },
      set changeTrackingMode(value: string) { state.modeChanges.push(value); queue.push(() => { state.tracking = value }) },
      load() {},
      getSelection: () => new Range(state.selection[0]!, state.selection[1]!),
      getParagraphByUniqueLocalId(id: string) { return paragraph(position(id).p) },
      getAnnotationById(id: string) { return { delete() { state.annotations.delete(id) } } },
      body: { get text() { return text() }, load() {}, get paragraphs() { return collection(state.paragraphs.map(paragraph)) }, getRange() { return new Range(0, text().length) } },
      onAnnotationClicked: event("annotation"), onParagraphChanged: event("change"), onParagraphAdded: event("add"), onParagraphDeleted: event("delete"),
    },
    async sync() {
      const writes = queue.splice(0)
      for (const write of writes) write()
      state.afterSync?.()
      if (state.throwAfterWrite && writes.length) throw new Error("lost native response")
    },
  } as WordChatContext
  const word = { async run<T>(first: unknown, second?: (context: WordChatContext) => Promise<T>) { return (second ?? first as (context: WordChatContext) => Promise<T>)(context) } } as WordChatWord
  const nativeHost = new WordChatHost(word, () => true)
  const author = (input: Omit<WordChatApplyInput, "author">) => async (): Promise<WordChatAuthoredAmendment> => {
    state.authorCalls++; await state.duringAuthor?.()
    const replacements = input.replacement.split("\r"), revisions: WordChatAuthoredAmendment["revisions"] = []
    let counter = 0
    const paragraphs = input.target.paragraphs.map((p, index) => {
      const start = index === 0 ? input.target.startOffset : 0, end = index === input.target.paragraphs.length - 1 ? input.target.endOffset : p.text.length
      return p.text.slice(0, start) + replacements[index]! + p.text.slice(end)
    })
    const run = (value: string) => value ? `<w:r><w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r>` : ""
    const markup = paragraphs.map((after, index) => {
      const before = input.target.paragraphs[index]!.text
      let at = 0, xml = ""
      for (const change of wordChanges(before, after)) {
        xml += run(before.slice(at, change.offset))
        if (change.replaced) { xml += `<w:del w:id="${++counter}" w:author="doc.haus"><w:r><w:delText xml:space="preserve">${escapeXml(change.replaced)}</w:delText></w:r></w:del>`; revisions.push({ author: "doc.haus", type: "Deleted", text: change.replaced }) }
        if (change.changed) { xml += `<w:ins w:id="${++counter}" w:author="doc.haus">${run(change.changed)}</w:ins>`; revisions.push({ author: "doc.haus", type: "Added", text: change.changed }) }
        at = change.offset + change.replaced.length
      }
      return `<w:p>${xml}${run(before.slice(at))}</w:p>`
    }).join("")
    return { proposalId: input.proposalId, author: "doc.haus", paragraphs, revisions, ooxml: flatOpc(markup).replace(/<w:sectPr>[\s\S]*?<\/w:sectPr>/, "") }
  }
  type TestInput = Omit<WordChatApplyInput, "author"> & { author?: WordChatApplyInput["author"] }
  const host = new Proxy(nativeHost, { get(target, key) {
    if (key === "apply") return (input: TestInput) => target.apply({ ...input, author: input.author ?? author(input) })
    const value = (target as unknown as Record<PropertyKey, unknown>)[key]
    return typeof value === "function" ? value.bind(target) : value
  } }) as Omit<WordChatHost, "apply"> & { apply(input: TestInput): ReturnType<WordChatHost["apply"]> }
  return { state, text, word, host, nativeHost, author,
    emit(name: string, args = {}) { for (const callback of state.handlers.get(name) ?? []) callback(args as never) },
  }
}

test("current context reads unsaved body without requiring a selection and empty selection stays empty", async () => {
  const { state, host } = office(["Unsaved wording.", "A different clause."])
  state.selection = [0,0]
  expect(await host.capture()).toMatchObject({text:"Unsaved wording.\rA different clause.", saved:false,truncated:false})
  expect(await host.selection()).toBeUndefined()
  expect(state.writes).toHaveLength(0)
  expect(state.selected).toBe(0)
})

test("current document context beyond 96k retains a selected later passage without truncation", async () => {
  const paragraphs = Array.from({length:6},(_,index) => `Section ${index}: ${"Ordinary current wording. ".repeat(800)}`)
  paragraphs.push("A later selected clause.")
  const {state,host,text} = office(paragraphs)
  const at = text().indexOf("later selected")
  state.selection = [at,at + "later selected".length]
  const context = await host.capture(true)
  expect(context.text.length).toBeGreaterThan(96_000)
  expect(context.text).toBe(text())
  expect(context.truncated).toBe(false)
  expect(context.selection?.text).toBe("later selected")
  expect(context.selection?.paragraphs.map(p => p.id)).toEqual(["paragraph-6"])
  expect(state.writes).toHaveLength(0)
})

test("over-limit current documents refuse clearly without sending a silently truncated buffer", async () => {
  const {host,state} = office(Array.from({length:26},() => "Ordinary text. ".repeat(1500)))
  await expect(host.capture()).rejects.toThrow("512,000 characters")
  expect(state.writes).toHaveLength(0)
})

test("paragraph identity targets a repeated clause without touching the other copy", async () => {
  const duplicate = "A clear obligation applies."
  const { state, host, text } = office([duplicate, duplicate])
  state.selection = [duplicate.length + 1, duplicate.length * 2 + 1]
  const target = (await host.selection())!
  const result = await host.apply({ proposalId: "reviewed-1", target, replacement: duplicate.replace("clear", "specific") })
  expect(text()).toBe(duplicate + "\rA specific obligation applies.")
  expect(result).toMatchObject({ status: "applied", tracked: true, saved: false, authors: ["doc.haus"] })
  expect(state.writes.map((w) => [w.before, w.text, w.tracking])).toEqual([[duplicate, "A specific obligation applies.", "Off"]])
  expect(state.tracking).toBe("Off")
})

test("multiline authored markup uses one scoped import and restores TrackMineOnly", async () => {
  const { state, host, text } = office(["The first obligation is clear.", "The second obligation is clear."])
  state.selection = [0, text().length]
  state.tracking = "TrackMineOnly"
  const target = (await host.selection())!
  await host.apply({ proposalId: "multiline", target, replacement: target.text.replaceAll("clear", "specific") })
  expect(text()).toBe("The first obligation is specific.\rThe second obligation is specific.")
  expect(state.writes).toHaveLength(1)
  expect(state.writes[0]!.text).toBe(text())
  expect(state.modeChanges).toEqual(["Off", "TrackMineOnly"])
  expect(state.tracking).toBe("TrackMineOnly")
})

test("fresh paragraph proof confirms multiline edits despite returned insertion-range drift", async () => {
  const { state, host, text } = office(["The supplier will promptly identify an issue.", "The recipient will promptly consider the issue."])
  state.selection = [0, text().length]
  state.driftInsertedRanges = true
  const target = (await host.selection())!
  const result = await host.apply({ proposalId: "native-returned-range-drift", target, replacement: target.text.replaceAll("promptly", "within a reasonable time") })
  expect(result.status).toBe("applied")
  expect(text()).toBe(target.text.replaceAll("promptly", "within a reasonable time"))
  expect(state.nativeRevisions.map(change => change.type).sort()).toEqual(["Added", "Added", "Deleted", "Deleted"])
  expect(state.tracking).toBe("Off")
})

test("native proof covers deletion-only and replaced spans without inserted-range assumptions", async () => {
  for (const replacement of ["A clause with an obligation.", "A clause with a specific duty."]) {
    const { state, host } = office()
    const target = (await host.selection())!
    expect((await host.apply({ proposalId: replacement, target, replacement })).status).toBe("applied")
    expect(state.paragraphs[0]!.text).toBe(replacement)
    expect(state.nativeRevisions.some(change => change.type === "Deleted")).toBe(true)
  }
})

test("two separated edits import in document order and preserve revisions in other paragraphs", async () => {
  const { state, host, word } = office(["A clear obligation preserves the agreed commercial position and a firm deadline.", "Earlier revision remains."])
  const priorStart = state.paragraphs[0]!.text.length + 1
  state.nativeRevisions = [
    { start: priorStart, end: priorStart + 7, author: "Earlier reviewer", type: "Added", text: "Earlier" },
    { start: priorStart + 8, end: priorStart + 8, author: "Earlier reviewer", type: "Deleted", text: "old wording " },
  ]
  const target = (await host.selection())!
  const replacement = target.text.replace("clear", "specific").replace("firm", "fixed")
  expect(wordChanges(target.text, replacement).map(edit => [edit.replaced, edit.changed])).toEqual([["clear", "specific"], ["firm", "fixed"]])
  expect((await host.apply({ proposalId: "two-separated-edits", target, replacement })).status).toBe("applied")
  expect(state.writes.map(write => [write.before, write.text])).toEqual([[target.text, replacement]])
  expect(state.nativeRevisions.filter(change => change.author === "Earlier reviewer").map(({ author, type, text }) => ({ author, type, text }))).toEqual([
    { author: "Earlier reviewer", type: "Added", text: "Earlier" }, { author: "Earlier reviewer", type: "Deleted", text: "old wording " },
  ])
  const changes = await word.run(async context => {
    const collection = context.document.body.getRange().getTrackedChanges()
    collection.load("items/author,items/type,items/text")
    await context.sync()
    return collection.items.filter(change => change.author === "doc.haus").map(change => [change.type, change.text])
  })
  expect(changes).toEqual([["Deleted", "clear"], ["Added", "specific"], ["Deleted", "firm"], ["Added", "fixed"]])
  expect(state.paragraphs[0]!.text).toBe(replacement)
})

test("preexisting revisions elsewhere cannot masquerade as new authored markup", async () => {
  const { state, host } = office(["A clear obligation.", "specific clear."])
  const priorStart = state.paragraphs[0]!.text.length + 1
  state.nativeRevisions = [
    { start: priorStart, end: priorStart + 8, author: "doc.haus", type: "Added", text: "specific" },
    { start: priorStart + 9, end: priorStart + 9, author: "doc.haus", type: "Deleted", text: "clear" },
  ]
  state.omitNewRevisions = true
  const target = (await host.selection())!
  const input = { proposalId: "preexisting-is-not-proof", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply(input)).rejects.toThrow("could not confirm")
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toThrow("may already contain")
})

test("edits outside the target are preserved; changed target paragraphs refuse before writing", async () => {
  const { state, host } = office(["A clause with a clear obligation.", "The first qualification.", "The second qualification.", "A distant clause."])
  const target = (await host.selection())!
  state.paragraphs[3]!.text += " Unsaved local drafting."
  await host.apply({ proposalId: "outside", target, replacement: target.text.replace("clear", "specific") })
  expect(state.paragraphs[3]!.text).toEndWith("Unsaved local drafting.")
  const changed = office()
  const stale = (await changed.host.selection())!
  changed.state.paragraphs[0]!.text += " Changed qualification."
  await expect(changed.host.apply({ proposalId: "stale", target: stale, replacement: "Different wording." })).rejects.toThrow("passage has changed")
  expect(changed.state.writes).toHaveLength(0)
  expect(changed.state.modeChanges).toHaveLength(0)
})

test("complete paragraph delivery keeps revision authority inside the exact selection", async () => {
  const { state, host, text } = office(["Unchanged bold prefix. A clear obligation. Unchanged italic suffix."])
  const start = text().indexOf("A clear")
  state.selection = [start, start + "A clear obligation.".length]
  const target = (await host.selection())!
  await host.apply({ proposalId: "subspan", target, replacement: "A specific obligation." })
  expect(text()).toBe("Unchanged bold prefix. A specific obligation. Unchanged italic suffix.")
  expect(state.writes.map((w) => w.before)).toEqual([target.paragraphs[0]!.text])
  expect(state.nativeRevisions.map(change => change.text)).toEqual(["clear", "specific"])
})

test("verified imported scope returns new IDs and exact selected offsets without following an adjacent cursor", async () => {
  for (const multiline of [false, true]) {
    const paragraphs = multiline ? ["Prefix clear.", "clear suffix.", "An unrelated adjacent clause."] : ["Prefix clear suffix.", "Prefix clear suffix."]
    const { state, host } = office(paragraphs)
    state.selection = [7, multiline ? paragraphs[0]!.length + 1 + 5 : 12]
    state.renewImportedIds = true
    const target = (await host.selection())!, before = structuredClone(target)
    const replacement = multiline ? "specific.\rprecise" : "very specific"
    state.duringAuthor = () => { const at = paragraphs.slice(0, -1).reduce((n, text) => n + text.length + 1, 0); state.selection = [at, at + paragraphs.at(-1)!.length] }
    const input = { proposalId: `verified-import-scope-${multiline}`, target, replacement }
    const applied = await host.apply(input), next = applied.target!
    expect(next.text).toBe(replacement)
    expect(next.paragraphs).toEqual(multiline ? [
      { id: "imported-paragraph-0", text: "Prefix specific." }, { id: "imported-paragraph-1", text: "precise suffix." },
    ] : [{ id: "imported-paragraph-0", text: "Prefix very specific suffix." }])
    expect(next.startOffset).toBe(7)
    expect(next.endOffset).toBe(multiline ? "precise".length : 7 + replacement.length)
    expect(next.saved).toBe(false)
    expect(next.id).not.toBe(target.id)
    expect(Date.parse(next.capturedAt)).toBeGreaterThanOrEqual(Date.parse(target.capturedAt))
    expect(target).toEqual(before)
    expect(state.paragraphs.at(-1)).toMatchObject({ id: `paragraph-${paragraphs.length - 1}`, text: paragraphs.at(-1)! })
    expect((await host.apply(input)).target).toBeUndefined()
    expect(state.writes).toHaveLength(1)
  }
})

test("deleting the complete selected wording does not invent an expanded follow-up target", async () => {
  const { state, host } = office(["Prefix clear suffix."])
  state.selection = [7, 13]
  const target = (await host.selection())!
  const applied = await host.apply({ proposalId: "empty-follow-up-scope", target, replacement: "" })
  expect(applied.status).toBe("applied")
  expect(applied.target).toBeUndefined()
  expect(state.paragraphs[0]!.text).toBe("Prefix suffix.")
})

test("a retry and simultaneous approvals cannot replay the same local proposal", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  const input = { proposalId: "once", target, replacement: target.text.replace("clear", "specific") }
  await Promise.all([host.apply(input), host.apply(input)])
  expect(state.writes).toHaveLength(1)
  expect((await host.apply(input)).status).toBe("already-applied")
  await expect(host.apply({ ...input, replacement: "Unreviewed replacement." })).rejects.toThrow("different wording")
  expect(state.writes).toHaveLength(1)
})

test("a lost native response restores tracking and never blindly retries a possible edit", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.throwAfterWrite = true
  const input = { proposalId: "uncertain", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply(input)).rejects.toThrow("Inspect the clause")
  expect(state.tracking).toBe("Off")
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toThrow("may already contain")
  expect(state.writes).toHaveLength(1)
})

test("comments and revisions anywhere in imported paragraphs refuse; siblings remain untouched", async () => {
  for (const kind of ["comments", "revisions"] as const) {
    const { state, host } = office()
    state.selection = [16, 21]
    const target = (await host.selection())!
    state[kind] = [[0, 1]]
    await expect(host.apply({ proposalId: kind, target, replacement: "specific" })).rejects.toThrow("existing revisions, comments")
    expect(state.writes).toHaveLength(0)
    expect(state.authorCalls).toBe(0)
    const sibling = state.paragraphs[0]!.text.length + 1
    state[kind] = [[sibling + 1, sibling + 2]]
    await host.apply({ proposalId: kind + "-outside", target, replacement: "specific" })
    expect(state[kind]).toHaveLength(1)
    expect(state.paragraphs[1]!.text).toBe("Another clause.")
  }
})

test("changed paragraph structure and field objects are disclosed before mutation", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  await expect(host.apply({ proposalId: "structure", target, replacement: target.text + "\rA new paragraph." })).rejects.toThrow("paragraph structure")
  state.xml = `<w:p xmlns:w="${WORD_XML}"><w:fldSimple/></w:p>`
  await expect(host.apply({ proposalId: "field", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("linked Word object")
  expect(state.writes).toHaveLength(0)
})

test("native Flat OPC page layout and unrelated package parts do not block ordinary text", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.xml = flatOpc(`<w:p><w:r><w:t>${escapeXml(target.text)}</w:t></w:r></w:p>`, `<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData><w:hdr><w:p><w:fldSimple w:instr="PAGE"/></w:p></w:hdr></pkg:xmlData></pkg:part>`)
  await host.apply({ proposalId: "native-flat-opc", target, replacement: target.text.replace("clear", "specific") })
  expect(state.writes).toHaveLength(1)
  expect(state.writes[0]!.tracking).toBe("Off")
  expect(state.tracking).toBe("Off")
})

test("selected Flat OPC fields, links, revisions, comments and section breaks stay guarded", async () => {
  const unsafe = [
    "<w:p><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>clear</w:t></w:r></w:fldSimple></w:p>",
    "<w:p><w:hyperlink><w:r><w:t>clear</w:t></w:r></w:hyperlink></w:p>",
    "<w:p><w:ins><w:r><w:t>clear</w:t></w:r></w:ins></w:p>",
    "<w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:t>clear</w:t></w:r></w:p>",
    "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>clear</w:t></w:r></w:p>",
    "<w:sectPr/><w:p><w:r><w:t>clear</w:t></w:r></w:p>",
  ]
  for (const [index, fragment] of unsafe.entries()) {
    const { state, host } = office()
    const target = (await host.selection())!
    state.xml = flatOpc(fragment)
    await expect(host.apply({ proposalId: `native-unsafe-${index}`, target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("linked Word object")
    expect(state.writes).toHaveLength(0)
    expect(state.modeChanges).toHaveLength(0)
  }
})

test("malformed or ambiguous native XML cannot authorize an edit", async () => {
  for (const xml of ["<w:p><broken>", "<unexpected/>", flatOpc("<w:p/>").replace("</pkg:package>", flatOpc("<w:p/>").match(/<pkg:part[\s\S]*<\/pkg:part>/)![0] + "</pkg:package>")]) {
    const { state, host } = office()
    const target = (await host.selection())!
    state.xml = xml
    await expect(host.apply({ proposalId: "malformed-native-xml", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("linked Word object")
    expect(state.writes).toHaveLength(0)
  }
})

test("unexpected inserted wording never reports a confirmed reviewed amendment", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.corruptWrite = true
  const input = { proposalId: "wrong-host-result", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply(input)).rejects.toThrow("could not confirm")
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toThrow("may already contain")
})

test("the native host requires shared authoring and never falls back to a text edit", async () => {
  const { state, nativeHost } = office()
  const target = (await nativeHost.selection())!
  await expect(nativeHost.apply({ proposalId: "missing-author", target, replacement: "Different wording." } as WordChatApplyInput)).rejects.toThrow("Reconnect to doc.haus")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test("a formatting change during shared authoring stops import before changing tracking", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.duringAuthor = () => { state.xml = flatOpc(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(target.text)}</w:t></w:r></w:p>`) }
  await expect(host.apply({ proposalId: "late-format", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("formatting changed")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test("a server artifact bound to another proposal is rejected before native mutation", async () => {
  const { state, host, author } = office()
  const target = (await host.selection())!
  const input = { proposalId: "artifact-binding", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply({ ...input, author: async () => ({ ...await author(input)(), proposalId: "other-proposal" }) })).rejects.toThrow("matches the reviewed passage")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})
