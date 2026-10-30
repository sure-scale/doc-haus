import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { createReplaceConfirmation } from "./replace-confirmation"

function fixture() {
  const page = new Window()
  page.document.body.innerHTML = `<main><button id="load-latest">Load latest version</button><textarea>Keep this draft</textarea><nav class="pane-tabs"><button role="tab" aria-selected="true">Compare</button></nav></main><aside inert="inert" aria-hidden="false"></aside>`
  const document = page.document as unknown as Document
  const trigger = document.getElementById("load-latest") as HTMLButtonElement
  trigger.focus()
  const confirmation = createReplaceConfirmation(document)
  return { page, document, trigger, confirmation }
}

function key(page: Window, value: string, shiftKey = false) {
  page.document.dispatchEvent(new page.KeyboardEvent("keydown", { key: value, shiftKey, bubbles: true, cancelable: true }))
}

test("replacement shows the exact version and traps focus; Escape leaves the document and draft alone", async () => {
  const { page, document, trigger, confirmation } = fixture()
  Object.defineProperty(page, "confirm", { value: () => { throw new Error("Native browser confirmation must not be used") } })
  try {
    const answer = confirmation.confirm({ number: 7 })
    const dialog = document.querySelector('[role="alertdialog"]')!
    const cancel = dialog.querySelector<HTMLButtonElement>('[data-replace="cancel"]')!
    const accept = dialog.querySelector<HTMLButtonElement>('[data-replace="accept"]')!
    expect(dialog.getAttribute("aria-modal")).toBe("true")
    expect(document.getElementById(dialog.getAttribute("aria-labelledby")!)?.textContent).toBe("Load version 7?")
    expect(document.getElementById(dialog.getAttribute("aria-describedby")!)?.textContent).toContain("unsaved changes")
    expect(document.activeElement === cancel).toBe(true)
    key(page, "Tab"); expect(document.activeElement === accept).toBe(true)
    key(page, "Tab", true); expect(document.activeElement === cancel).toBe(true)
    trigger.focus(); expect(document.activeElement === cancel).toBe(true)
    expect(document.querySelector("main")!.hasAttribute("inert")).toBe(true)
    key(page, "Escape")
    expect(await answer).toBe(false)
    await page.happyDOM.waitUntilComplete()
    expect(document.querySelector('[role="alertdialog"]')).toBeNull()
    expect(document.activeElement === trigger).toBe(true)
    expect(document.querySelector("textarea")!.value).toBe("Keep this draft")
    expect(document.querySelector("main")!.hasAttribute("inert")).toBe(false)
    expect(document.querySelector("main")!.hasAttribute("aria-hidden")).toBe(false)
    expect(document.querySelector("aside")!.getAttribute("inert")).toBe("inert")
    expect(document.querySelector("aside")!.getAttribute("aria-hidden")).toBe("false")
  } finally { confirmation.dispose(); page.happyDOM.abort() }
})

test("Replace authorizes one write and stays busy until completion; hidden trigger restores the selected tab", async () => {
  const { page, document, trigger, confirmation } = fixture()
  let writes = 0, finish!: () => void
  const nativeWrite = new Promise<void>(done => { finish = done })
  const loading = (async () => {
    try {
      if (await confirmation.confirm({ number: 2 })) { writes++; await nativeWrite }
    } finally { confirmation.complete() }
  })()
  try {
    expect(writes).toBe(0)
    const dialog = document.querySelector<HTMLElement>('[role="alertdialog"]')!
    const cancel = dialog.querySelector<HTMLButtonElement>('[data-replace="cancel"]')!
    const accept = dialog.querySelector<HTMLButtonElement>('[data-replace="accept"]')!
    accept.click(); accept.click(); await Bun.sleep(0)
    expect(writes).toBe(1)
    expect(cancel.disabled).toBe(true); expect(accept.disabled).toBe(true)
    expect(dialog.getAttribute("aria-busy")).toBe("true")
    expect(dialog.querySelector<HTMLElement>('[role="status"]')!.hidden).toBe(false)
    cancel.click(); key(page, "Escape"); key(page, "Tab")
    expect(document.querySelector('[role="alertdialog"]') === dialog).toBe(true)
    expect(document.activeElement === dialog).toBe(true)
    // A successful load hides the stale-version action before the queue ends.
    trigger.hidden = true
    finish(); await loading; await page.happyDOM.waitUntilComplete()
    expect(document.querySelector('[role="alertdialog"]')).toBeNull()
    expect(document.activeElement === document.querySelector('[role="tab"][aria-selected="true"]')).toBe(true)
    expect(document.querySelector("main")!.hasAttribute("inert")).toBe(false)
    expect(writes).toBe(1)
  } finally { finish(); await loading; confirmation.dispose(); page.happyDOM.abort() }
})

test("Cancel and pane disposal resolve without replacement; overlapping version prompts cannot approve another version", async () => {
  const { page, document, confirmation } = fixture()
  try {
    const first = confirmation.confirm({ number: 3 })
    expect(await confirmation.confirm({ number: 4 })).toBe(false)
    expect(document.querySelectorAll('[role="alertdialog"]').length).toBe(1)
    expect(document.querySelector("h2")!.textContent).toBe("Load version 3?")
    document.querySelector<HTMLButtonElement>('[data-replace="cancel"]')!.click()
    expect(await first).toBe(false)
    const second = confirmation.confirm({ number: 4 })
    confirmation.dispose()
    expect(await second).toBe(false)
    expect(document.querySelector('[role="alertdialog"]')).toBeNull()
    expect(document.querySelector("main")!.hasAttribute("inert")).toBe(false)
    expect(await confirmation.confirm({ number: Number.NaN })).toBe(false)
    expect(document.querySelector('[role="alertdialog"]')).toBeNull()
  } finally { confirmation.dispose(); page.happyDOM.abort() }
})
