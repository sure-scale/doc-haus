import { expect, test } from "bun:test"
import { captureSelection, showSelection, type SelectionWord } from "./selection"

function host(text = "First clause. A neutral passage. Last clause.", selected = "A neutral passage.") {
  const state = { text, selected, saved: false, xml: "<body>Original formatting</body>", navigations: 0, searches: 0 }
  const word: SelectionWord = {
    run: async (action) =>
      action({
        document: {
          get saved() {
            return state.saved
          },
          load: () => {},
          getSelection: () => ({
            text: state.selected,
            load: () => {},
            select: () => {
              throw new Error("Selection must be searched afresh")
            },
          }),
          body: Object.assign({
            get text() {
              return state.text
            },
            load: () => {},
            search: (query: string, options: unknown) => {
              expect(options).toEqual({
                matchCase: true,
                matchWholeWord: false,
                matchWildcards: false,
                ignorePunct: false,
                ignoreSpace: false,
              })
              state.searches++
              const count = state.text.split(query).length - 1
              return {
                items: Array.from({ length: count }, () => ({
                  text: query,
                  load: () => {},
                  select: () => {
                    state.navigations++
                  },
                })),
                load: () => {},
              }
            },
          }, { getOoxml: () => ({ value: state.xml }) }),
        },
        sync: async () => {},
      }),
  }
  return { state, word }
}

test("capture only reads unsaved context; explicit Show selects one exact unchanged passage", async () => {
  const { state, word } = host()
  const snapshot = await captureSelection(word)
  expect(snapshot.text).toBe(state.selected)
  expect(snapshot.saved).toBe(false)
  expect(snapshot.navigable).toBe(true)
  expect(state.navigations).toBe(0)
  expect(state.searches).toBe(0)
  state.selected = "Last clause."
  await showSelection(snapshot, word)
  expect(state.navigations).toBe(1)
  expect(state.saved).toBe(false)
})

test("changed body text refuses a previously valid exact navigation", async () => {
  const { state, word } = host()
  const snapshot = await captureSelection(word)
  await showSelection(snapshot, word)
  expect(state.navigations).toBe(1)
  state.text += " Changed."
  await expect(showSelection(snapshot, word)).rejects.toThrow("Word text has changed")
  expect(state.navigations).toBe(1)
})

test("Office markup serialization churn does not block navigation to unchanged unique text", async () => {
  const { state, word } = host()
  const snapshot = await captureSelection(word)
  await showSelection(snapshot, word)
  state.xml = "<body>Different transient Office serialization</body>"
  state.saved = true
  await showSelection(snapshot, word)
  expect(state.navigations).toBe(2)
  expect(state.text).toBe("First clause. A neutral passage. Last clause.")
  expect(state.saved).toBe(true)
})

test("ambiguous repeated passages, no selection and nonbody context are disclosed without navigation", async () => {
  for (const [body, selection, reason] of [
    ["A clause. A clause.", "A clause.", "more than once"],
    ["A clause.", "", "Select a passage"],
    ["A clause.", "Header text.", "outside the document body"],
  ]) {
    const { state, word } = host(body, selection)
    const snapshot = await captureSelection(word)
    expect(snapshot.navigable).toBe(false)
    expect(snapshot.reason).toContain(reason)
    await expect(showSelection(snapshot, word)).rejects.toThrow()
    expect(state.navigations).toBe(0)
    if (selection === "A clause.") {
      await expect(showSelection({ ...snapshot, navigable: true }, word)).rejects.toThrow("one exact passage")
      expect(state.navigations).toBe(0)
    }
  }
})

test("long, multiline and find-code selections remain visible but are never searched as prefixes", async () => {
  for (const text of ["x".repeat(256), "First line.\rSecond line.", "Literal ^p find text."]) {
    const { state, word } = host(text, text)
    const snapshot = await captureSelection(word)
    expect(snapshot.text).toBe(text)
    expect(snapshot.reason).toContain("255")
    await expect(showSelection(snapshot, word)).rejects.toThrow("255")
    expect(state.searches).toBe(0)
    expect(state.navigations).toBe(0)
  }
})
