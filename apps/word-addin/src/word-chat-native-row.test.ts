import { expect, test } from "bun:test"
import { prepareNativeRow, queueNativeRow, type NativeRowContext } from "./word-chat-native-row"

function fixture() {
  const batches: string[][] = [], queued: string[] = []
  const table = (prefix: string) => ({
    rowCount: 1, isUniform: true, values: [[`${prefix}1`, `${prefix}2`]], load() {},
    getCell(row: number, column: number) { return { body: {
      paragraphs: { items: [{ uniqueLocalId: `${prefix}${column}`, text: `${prefix}${column + 1}`, tableNestingLevel: 1 }], load() {} },
      insertText(text: string, location: string) {
        queued.push(`text:${row}:${column}:${text}:${location}`)
        const font = new Proxy({ bold: true, italic: true, underline: "Single", load() {} }, { set(target, key, value) { queued.push(`font:${String(key)}:${value}`); return Reflect.set(target, key, value) } })
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
