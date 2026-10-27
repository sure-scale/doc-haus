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
          body: {
            get text() {
              return state.text
            },
            load: () => {},
            getOoxml: () => ({ value: state.xml }),
            search: (query, options) => {
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
          },
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

test("text or OOXML changes refuse the previously valid navigation instead of guessing", async () => {
  for (const change of ["text", "xml"] as const) {
    const { state, word } = host()
    const snapshot = await captureSelection(word)
    await showSelection(snapshot, word)
    expect(state.navigations).toBe(1)
    state[change] += " Changed."
    await expect(showSelection(snapshot, word)).rejects.toThrow("has changed")
    expect(state.navigations).toBe(1)
  }
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
