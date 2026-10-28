import { expect, test } from "bun:test"
import { FEATURE_TAB_KEY, preferRoomyPane, readFeatureTab, rememberFeatureTab } from "./feature-tabs"

test("feature preference defaults to Chat, restores each valid tab and leaves unavailable storage usable", () => {
  const values = new Map<string, string>()
  const storage = { getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => { values.set(key, value) } }
  expect(readFeatureTab(storage)).toEqual({ tab: "chat", obsolete: false })
  for (const tab of ["documents", "compare", "chat"] as const) { rememberFeatureTab(tab, storage); expect(readFeatureTab(storage)).toEqual({ tab, obsolete: false }) }
  expect([...values.keys()]).toEqual([FEATURE_TAB_KEY])
  values.set(FEATURE_TAB_KEY, "removed-feature")
  expect(readFeatureTab(storage)).toEqual({ tab: "chat", obsolete: true })
  expect(values.get(FEATURE_TAB_KEY)).toBe("removed-feature")
  expect(readFeatureTab({ getItem() { throw new Error("Disabled") } })).toEqual({ tab: "chat", obsolete: false })
})

test("initial width uses supported TaskPaneApi and tolerates silent or throwing host constraints", () => {
  const widths: number[] = []
  preferRoomyPane({ context: { requirements: { isSetSupported: (name: string, version: string) => name === "TaskPaneApi" && version === "1.1" } }, extensionLifeCycle: { taskpane: { setWidth: (width: number) => widths.push(width) } } })
  expect(widths).toEqual([500])
  preferRoomyPane({ context: { requirements: { isSetSupported: () => false } }, extensionLifeCycle: { taskpane: { setWidth: (width: number) => widths.push(width) } } })
  preferRoomyPane({ context: { requirements: { isSetSupported: () => true } }, extensionLifeCycle: { taskpane: { setWidth() { throw new Error("Host limit") } } } })
  expect(widths).toEqual([500])
})
