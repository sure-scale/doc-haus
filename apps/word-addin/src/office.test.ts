import { expect, test } from "bun:test"
import { zipSync, strToU8 } from "fflate"
import { documentFingerprint } from "./fingerprint"
import { exportFile, type OfficeFile } from "./office"

test("DOCX fingerprint ignores ZIP timestamps, compression and entry order but detects changed content", async () => {
  const a = strToU8("<document>one</document>")
  const b = strToU8("<styles/>")
  const first = zipSync({ "word/document.xml": a, "word/styles.xml": b }, { mtime: new Date(2020, 0, 1), level: 0 })
  const second = zipSync({ "word/styles.xml": b, "word/document.xml": a }, { mtime: new Date(2025, 0, 1), level: 6 })
  const changed = zipSync({ "word/document.xml": strToU8("<document>two</document>"), "word/styles.xml": b })
  expect(await documentFingerprint(first)).toBe(await documentFingerprint(second))
  expect(await documentFingerprint(changed)).not.toBe(await documentFingerprint(first))
})

test("sliced export assembles the file and always closes its handle on success or failure", async () => {
  for (const fails of [false, true]) {
    let closed = 0
    const file: OfficeFile = {
      size: 3,
      sliceCount: 2,
      getSliceAsync: (index, callback) =>
        callback({ status: fails ? "failed" : "succeeded", value: { data: index === 0 ? [1, 2] : [3] } }),
      closeAsync: (callback) => {
        closed++
        callback({ status: "succeeded", value: undefined })
      },
    }
    if (fails) await expect(exportFile(file)).rejects.toThrow("could not read")
    else expect(await exportFile(file)).toEqual(new Uint8Array([1, 2, 3]))
    expect(closed).toBe(1)
  }
})

test("fingerprint refuses excessive parts and inflated sizes before unpacking", async () => {
  const tooMany = Object.fromEntries(Array.from({ length: 1001 }, (_, index) => [`part${index}.xml`, strToU8("<p/>")]))
  await expect(documentFingerprint(zipSync(tooMany))).rejects.toThrow("package size limit")
  const bytes = zipSync({ "word/document.xml": strToU8("<document/>") })
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let at = 0; at < bytes.length - 46; at++) {
    if (view.getUint32(at, true) === 0x02014b50) {
      view.setUint32(at + 24, 257 * 1024 * 1024, true)
      break
    }
  }
  await expect(documentFingerprint(bytes)).rejects.toThrow("package size limit")
})
