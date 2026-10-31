import { expect, test } from "bun:test"
import { prepareNativeRow, queueNativeRow, type NativeRowContext } from "./word-chat-native-row"

function fixture() {
  const batches: string[][] = [], queued: string[] = []
  const table = (prefix: string) => ({
    rowCount: 1, isUniform: true, values: [[`${prefix}1`, `${prefix}2`]], load() {},
    getCell(row: number, column: number) { return { body: {
      paragraphs: { items: [{ uniqueLocalId: `${prefix}${column}`, text: `${prefix}${column + 1}`, tableNestingLevel: 1 }], load() {} },
      getRange() { return { getReviewedText() { return { value: `${prefix}${column + 1}` } } } },
      insertText(text: string, location: string) {
        queued.push(`text:${row}:${column}:${text}:${location}`)
        const font = new Proxy({ bold: true, italic: true, underline: "Single", load() {} }, { set(target, key, value) { queued.push(`font:${String(key)}:${value}`); (target as Record<PropertyKey, unknown>)[key] = value; return true } })
        return { font }
      },
    } } },
    addRows(location: string, count: number, values: string[][]) { queued.push(`row:${location}:${count}:${JSON.stringify(values)}`) },
  })
  const document = { mode: "Off", formatting: false, load() {}, body: { tables: { items: [table("a"), table("b")], load() {} } },
    get changeTrackingMode() { return this.mode }, set changeTrackingMode(value: string) { queued.push(`tracking:${value}`); this.mode = value },
    get trackFormatting() { return this.formatting }, set trackFormatting(value: boolean) { queued.push(`formatting:${value}`); this.formatting = value },
  }
  const context: NativeRowContext = { document, async sync() { batches.push(queued.splice(0)) } }
  const binding = { tableIndex: 1, rows: [[{ id: "b0", text: "b1" }, { id: "b1", text: "b2" }]] }
  return { context, document, batches, queued, binding }
}

test("binds the reviewed cells among multiple tables and queues one mutation flush", async () => {
  const f = fixture(), prepared = await prepareNativeRow(f.context, f.binding)
  expect(f.batches).toEqual([[], [], []])
  const witness = queueNativeRow(f.context, prepared, ["First", "Second"])
  expect(witness.ranges).toHaveLength(2)
  expect(f.document.changeTrackingMode).toBe("Off")
  expect(f.document.trackFormatting).toBe(false)
  await f.context.sync()
  expect(f.batches[3]).toEqual([
    'tracking:TrackAll', 'row:End:1:[["",""]]', 'text:1:0:First:End', 'font:bold:false', 'font:italic:false', 'font:underline:None',
    'text:1:1:Second:End', 'font:bold:false', 'font:italic:false', 'font:underline:None', 'tracking:Off', 'formatting:false',
  ])
})

test("rejects a stale ordinal or changed cell without queuing a mutation", async () => {
  for (const change of ["ordinal", "text", "shape"] as const) {
    const f = fixture()
    if (change === "ordinal") f.binding.tableIndex = 0
    if (change === "text") f.binding.rows[0]![0]!.text = "changed"
    if (change === "shape") f.context.document.body.tables.items[1]!.isUniform = false
    await expect(prepareNativeRow(f.context, f.binding)).rejects.toThrow("reviewed Word table")
    expect(f.queued).toEqual([])
  }
})

test("a formatted row inserts each reviewed run with explicit flags in the same tracked flush", async () => {
  const f = fixture(), prepared = await prepareNativeRow(f.context, f.binding)
  const run = (text: string, bold = false, italic = false, underline = false) => ({ text, bold, italic, underline })
  const witness = queueNativeRow(f.context, prepared, ["Fee: 100", "Note"], [[run("Fee:", true), run(" 100")], [run("Note", false, true, true)]])
  expect(witness.ranges).toHaveLength(3)
  await f.context.sync()
  expect(f.batches[3]).toEqual([
    'tracking:TrackAll', 'row:End:1:[["",""]]',
    'text:1:0:Fee::End', 'font:bold:true', 'font:italic:false', 'font:underline:None',
    'text:1:0: 100:End', 'font:bold:false', 'font:italic:false', 'font:underline:None',
    'text:1:1:Note:End', 'font:bold:false', 'font:italic:true', 'font:underline:Single',
    'tracking:Off', 'formatting:false',
  ])
  // An empty cell in a formatted row carries no runs and stays a plain insertion.
  const g = fixture(), ready = await prepareNativeRow(g.context, g.binding)
  queueNativeRow(g.context, ready, ["Fee:", ""], [[run("Fee:", true)], []])
  await g.context.sync()
  expect(g.batches[3]!.slice(6)).toEqual(['text:1:1::End', 'font:bold:false', 'font:italic:false', 'font:underline:None', 'tracking:Off', 'formatting:false'])
})

test("formatting that does not match its cells refuses before queuing any write", async () => {
  const run = (text: string, bold = false) => ({ text, bold, italic: false, underline: false })
  for (const runs of [[[run("Fee:", true), run(" 10")], []], [[run("Fee:", true), run(" 100")]], [[{ ...run("Fee:"), bold: "yes" as unknown as boolean }, run(" 100")], []], [[run("", true), run("Fee: 100")], []]]) {
    const f = fixture(), prepared = await prepareNativeRow(f.context, f.binding)
    expect(() => queueNativeRow(f.context, prepared, ["Fee: 100", ""], runs)).toThrow("formatting does not match")
    expect(f.queued).toEqual([])
  }
})
