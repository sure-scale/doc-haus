// Run from apps/web: bun test/material-primitives.browser.ts
// Real shared controls, with only document-text HTTP mocked. No model or engine calls.
// PLAYWRIGHT_CHROMIUM_PATH can select an installed Chromium when the package cache differs.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-material-primitives-"))
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement as h, Fragment, useEffect, useLayoutEffect, useRef, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { Tooltip } from ${source("components/Tooltip.tsx")}
import ModalHeader from ${source("components/ModalHeader.tsx")}
import DocumentViewer from ${source("components/DocumentViewer.tsx")}
import useModalFocus from ${source("components/useModalFocus.ts")}

function NativeDialog({ onClose }) {
  const ref = useRef(null)
  useLayoutEffect(() => { ref.current.showModal() }, [])
  return h("dialog", { ref, "aria-label": "Nested confirmation", onClose },
    h("label", null, "Confirmation note", h("input", { "aria-label": "Confirmation note" })),
    h(Tooltip, { label: "Help remains above this native dialog." }, h("button", { type: "button" }, "Native help")),
    h("button", { type: "button", onClick: () => ref.current.close() }, "Close confirmation"))
}

function EmbeddedViewer({ onClose }) {
  const panel = useRef(null)
  const [ready, setReady] = useState(false)
  useModalFocus(panel, ready)
  useEffect(() => setReady(true), [])
  if (!ready) return h("p", { role: "status" }, "Loading embedded document")
  return h("div", { className: "viewer-overlay" },
    h("section", { ref: panel, className: "viewer-panel", role: "dialog", "aria-modal": true, "aria-label": "Embedded document" },
      h(ModalHeader, { title: "Embedded document", onClose }),
      h("iframe", { title: "Embedded source", srcDoc: '<!doctype html><html lang="en"><body><button type="button">Read next section</button></body></html>' })))
}

function Harness() {
  const [picker, setPicker] = useState(false)
  const [native, setNative] = useState(false)
  const [viewer, setViewer] = useState(false)
  const [embedded, setEmbedded] = useState(false)
  return h(Fragment, null,
    h("button", { type: "button", onClick: () => setPicker(true) }, "Open picker"),
    h("button", { type: "button", onClick: () => setViewer(true) }, "Read source email"),
    h("button", { type: "button", onClick: () => setEmbedded(true) }, "Open embedded viewer"),
    h("p", { id: "existing-description" }, "Existing help for this action."),
    h("div", { className: "tooltip-hover" },
      h(Tooltip, { label: "Additional context remains readable while you move the pointer onto this help." },
        h("button", { type: "button" }, "Hover help"))),
    h("div", { className: "tooltip-corner" },
      h(Tooltip, { label: "This long document action explains what happens before you continue, without covering or clipping the control." },
        h("button", { type: "button", "aria-describedby": "existing-description" }, "More information"))),
    picker && h("div", { className: "viewer-overlay" },
      h("section", { className: "picker-panel" },
        h(ModalHeader, { title: "Document actions", onClose: () => setPicker(false) }),
        h("div", { className: "picker-body" },
          h("label", null, "Document name", h("input", { "aria-label": "Document name" })),
          h("button", { type: "button", onClick: () => setNative(true) }, "Open confirmation"),
          h("button", { type: "button", disabled: true }, "Unavailable action"),
          h("button", { type: "button", hidden: true }, "Hidden action"),
          h("button", { type: "button", onClick: () => setPicker(false) }, "Done")),
        native && h(NativeDialog, { onClose: () => setNative(false) }))),
    viewer && h(DocumentViewer, { matterId: "material", name: "Source instructions.eml", onClose: () => setViewer(false) }),
    embedded && h(EmbeddedViewer, { onClose: () => setEmbedded(false) }))
}
createRoot(document.getElementById("root")).render(h(Harness))
`)
const build = await Bun.build({
  entrypoints: [entry], target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "/ingest", VITE_OPENCODE_URL: "/" }) },
})
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = (await Promise.all([
  "styles.css", "narration.css", "narration-live.css", "styles/primitives.css", "styles/upload.css",
  "styles/shell.css", "styles/chat-log.css", "styles/drafting.css", "styles/drafting-items.css",
  "styles/run-summary.css", "styles/motion.css", "styles/design-system.css", "styles/design-shell.css",
  "styles/design-documents.css", "styles/design-conversation.css", "styles/design-library.css", "styles/design-review.css",
].map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const requests: string[] = []
const server = Bun.serve({
  port: 0, hostname: "127.0.0.1",
  fetch(request) {
    const { pathname } = new URL(request.url)
    if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (pathname === "/") return new Response(`<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><style>${css}
      body { margin: 0; padding: 24px; }
      .tooltip-corner { position: fixed; right: 4px; bottom: 4px; }
      .tooltip-hover { position: fixed; left: 400px; top: 150px; }
      .picker-body { padding: 20px; display: grid; gap: 12px; }
      .picker-body input { width: 100%; }
      .picker-body [hidden] { display: none; }
    </style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    requests.push(`${request.method} ${pathname}`)
    if (request.method === "GET" && pathname === "/ingest/matters/material/documents/text")
      return Response.json({ text: "Subject: Synthetic instructions\n\nPlease review the supplied agreement.", injection: null })
    return Response.json({ error: "Unexpected request" }, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const errors: string[] = []
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, reducedMotion: "reduce" })
  page.setDefaultTimeout(5000)
  page.on("pageerror", error => errors.push(error.message))
  await page.route("**/*", route => {
    if (new URL(route.request().url()).origin === server.url.origin) return route.continue()
    errors.push(`Unexpected external request: ${route.request().url()}`)
    return route.abort()
  })
  await page.goto(server.url.href)

  for (const textSize of ["compact", "standard", "large"]) for (const width of [375, 768, 1440]) {
    await page.setViewportSize({ width, height: 740 })
    await page.evaluate(size => { document.documentElement.dataset.text = size }, textSize)
    await page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))))
    const trigger = page.getByRole("button", { name: "More information" })
    // Move focus out first so each width exercises opening from the keyboard.
    await page.getByRole("button", { name: "Open picker" }).focus()
    await trigger.focus()
    const tooltip = page.getByRole("tooltip")
    await expect(tooltip).toBeVisible()
    const description = await trigger.getAttribute("aria-describedby")
    assert.ok(description?.includes("existing-description"), "existing descriptions survive the tooltip")
    assert.ok(description?.includes((await tooltip.getAttribute("id"))!), "the focused control describes its visible tooltip")
    const bounds = await tooltip.boundingBox()
    assert.ok(bounds && bounds.x >= 11 && bounds.y >= 11 && bounds.x + bounds.width <= width - 11 && bounds.y + bounds.height <= 729,
      `tooltip is clamped inside the ${width}px viewport with ${textSize} text: ${JSON.stringify(bounds)}`)
    const triggerBounds = await trigger.boundingBox()
    const expectedTop = Math.max(12, Math.min(triggerBounds!.y + triggerBounds!.height / 2 - bounds.height / 2, 740 - bounds.height - 12))
    assert.ok(Math.abs(bounds.y - expectedTop) <= 1, `tooltip remains anchored with ${textSize} text: top=${bounds.y}, expected=${expectedTop}`)
    await page.keyboard.press("Escape")
    await expect(tooltip).toHaveCount(0)
    await expect(trigger).toBeFocused()
    assert.equal(await trigger.getAttribute("aria-describedby"), "existing-description")
  }
  await page.evaluate(() => { document.documentElement.dataset.text = "standard" })
  console.log("PASS Tooltips open on keyboard focus, preserve descriptions, anchor/clamp at 375/768/1440px in every text preference and dismiss with Escape.")

  // The pointer must cross the actual 8px gap and remain on the help beyond
  // its dismissal delay. Keyboard focus stays outside the trigger throughout.
  await page.getByRole("button", { name: "Open picker" }).focus()
  const hoverHelp = page.getByRole("button", { name: "Hover help", exact: true })
  await hoverHelp.hover()
  const hoverTooltip = page.getByRole("tooltip")
  await expect(hoverTooltip).toBeVisible()
  const hoverTriggerBounds = (await hoverHelp.boundingBox())!
  const hoverBounds = (await hoverTooltip.boundingBox())!
  assert.ok(Math.abs(hoverBounds.x - hoverTriggerBounds.x - hoverTriggerBounds.width - 8) <= 1, "hover case has the actual 8px gap")
  await page.mouse.move(hoverTriggerBounds.x + hoverTriggerBounds.width + 4, hoverTriggerBounds.y + hoverTriggerBounds.height / 2)
  await page.waitForTimeout(80)
  await expect(hoverTooltip).toBeVisible()
  await page.mouse.move(hoverBounds.x + hoverBounds.width / 2, hoverBounds.y + hoverBounds.height / 2)
  await page.waitForTimeout(300)
  await expect(hoverTooltip).toBeVisible()
  assert.equal(await hoverTooltip.evaluate(element => {
    const rect = element.getBoundingClientRect()
    return document.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2) === element
  }), true, "the tooltip itself accepts pointer hover")
  await page.mouse.move(10, 10)
  await expect(hoverTooltip).toHaveCount(0)
  await hoverHelp.hover()
  await expect(hoverTooltip).toBeVisible()
  await page.keyboard.press("Escape")
  await expect(hoverTooltip).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Open picker" })).toBeFocused()
  await page.mouse.move(10, 10)
  console.log("PASS Hover help survives the 8px pointer gap, remains hoverable, dismisses after leaving both surfaces and handles Escape with focus elsewhere.")

  await page.getByRole("button", { name: "Open picker" }).click()
  const picker = page.getByRole("dialog", { name: "Document actions" })
  await expect(picker).toBeVisible()
  await expect(picker.getByRole("button", { name: "Close", exact: true })).toBeFocused()
  await picker.getByRole("button", { name: "Done", exact: true }).focus()
  await page.keyboard.press("Tab")
  await expect(picker.getByRole("button", { name: "Close", exact: true })).toBeFocused()
  await page.keyboard.press("Shift+Tab")
  await expect(picker.getByRole("button", { name: "Done", exact: true })).toBeFocused()
  console.log("PASS Picker dialogs receive initial focus and loop Tab/Shift+Tab around enabled visible controls.")

  await picker.getByRole("button", { name: "Open confirmation" }).click()
  const native = page.getByRole("dialog", { name: "Nested confirmation" })
  await expect(native).toBeVisible()
  await expect(native.getByRole("textbox", { name: "Confirmation note" }), "native dialog receives focus when opened").toBeFocused()
  const nativeHelp = native.getByRole("button", { name: "Native help" })
  const nativeTooltip = page.getByRole("tooltip")
  await nativeHelp.hover()
  await expect(nativeTooltip).toBeVisible()
  await page.keyboard.press("Escape")
  await expect(nativeTooltip).toHaveCount(0)
  await expect(native).toBeVisible()
  await expect(native.getByRole("textbox", { name: "Confirmation note" })).toBeFocused()
  await page.mouse.move(10, 10)
  for (const textSize of ["compact", "standard", "large"]) {
    await page.evaluate(size => { document.documentElement.dataset.text = size }, textSize)
    await page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))))
    await native.getByRole("textbox", { name: "Confirmation note" }).focus()
    await page.keyboard.press("Tab")
    await expect(nativeHelp).toBeFocused()
    await expect(nativeTooltip).toBeVisible()
    const helpBounds = await nativeHelp.boundingBox()
    const tooltipBounds = await nativeTooltip.boundingBox()
    assert.ok(tooltipBounds && Math.abs(tooltipBounds.y - (helpBounds!.y + helpBounds!.height / 2 - tooltipBounds.height / 2)) <= 1,
      `native-dialog tooltip remains anchored with ${textSize} text`)
    const paintedAboveDialog = await nativeTooltip.evaluate(element => {
      const previous = element.style.pointerEvents
      element.style.pointerEvents = "auto"
      const rect = element.getBoundingClientRect()
      const hit = document.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2)
      element.style.pointerEvents = previous
      return hit === element || element.contains(hit)
    })
    assert.equal(paintedAboveDialog, true, `tooltip paints above the native dialog's top layer with ${textSize} text`)
    await page.keyboard.press("Escape")
    await expect(nativeTooltip).toHaveCount(0)
    await expect(native).toBeVisible()
  }
  await page.keyboard.press("Tab")
  await expect(native.getByRole("button", { name: "Close confirmation" })).toBeFocused()
  // Browsers may hand a native dialog's boundary Tab to browser chrome. The
  // ancestor trap must leave this native default untouched, while real interior
  // Tab presses above prove the nested controls remain keyboard-accessible.
  const prevented = await native.getByRole("button", { name: "Close confirmation" }).evaluate(element =>
    !element.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true })))
  assert.equal(prevented, false, "the ancestor picker does not prevent native-dialog Tab")
  await page.keyboard.press("Escape")
  await expect(native).toHaveCount(0)
  await expect(picker).toBeVisible()
  await expect(picker.getByRole("button", { name: "Open confirmation" })).toBeFocused()
  await page.evaluate(() => { document.documentElement.dataset.text = "standard" })
  console.log("PASS A nested native dialog owns focus and Escape; its parent stays open and regains focus.")

  await page.keyboard.press("Escape")
  await expect(picker).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Open picker" })).toBeFocused()
  console.log("PASS Closing the picker restores focus to its original opener.")

  await page.getByRole("button", { name: "Open embedded viewer" }).click()
  const embedded = page.getByRole("dialog", { name: "Embedded document" })
  await expect(embedded).toBeVisible()
  await expect(embedded.getByRole("button", { name: "Close", exact: true })).toBeFocused()
  await page.keyboard.press("Tab")
  await expect(page.frameLocator('iframe[title="Embedded source"]').getByRole("button", { name: "Read next section" })).toBeFocused()
  await page.keyboard.press("Tab")
  await expect(embedded.getByRole("button", { name: "Close", exact: true })).toBeFocused()
  await page.keyboard.press("Escape")
  await expect(embedded).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Open embedded viewer" })).toBeFocused()
  console.log("PASS A viewer mounted after loading activates its focus trap; embedded content is reachable and Tab returns to the viewer.")

  await page.setViewportSize({ width: 375, height: 812 })
  await page.getByRole("button", { name: "Read source email" }).click()
  const viewer = page.getByRole("dialog", { name: "Source instructions.eml" })
  await expect(viewer).toBeVisible()
  await expect(viewer.locator("pre.document-text")).toContainText("Please review the supplied agreement.")
  const bounds = await viewer.boundingBox()
  assert.ok(bounds && bounds.x >= 0 && bounds.y >= 0 && bounds.x + bounds.width <= 375 && bounds.y + bounds.height <= 812,
    `mobile document reader stays inside 375x812: ${JSON.stringify(bounds)}`)
  await expect(viewer.getByRole("button", { name: "Close", exact: true })).toBeFocused()
  await page.keyboard.press("Tab")
  await expect(viewer.getByRole("button", { name: "Close", exact: true })).toBeFocused()
  await page.keyboard.press("Escape")
  await expect(viewer).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Read source email" })).toBeFocused()
  assert.deepEqual(requests, ["GET /ingest/matters/material/documents/text"], "text view uses only the mocked read endpoint")
  assert.deepEqual(errors, [])
  console.log("PASS The source-email viewer traps focus, closes with Escape and restores its opener; no engine or model calls.")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
