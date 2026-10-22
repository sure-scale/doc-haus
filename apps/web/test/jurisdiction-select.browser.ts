// Isolated viewport regression for the real jurisdiction picker and product CSS.
// No API or model calls. Run from apps/web; optionally set PLAYWRIGHT_CHROMIUM_PATH.
import assert from "node:assert/strict"
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-jurisdiction-"))
const source = path.resolve(import.meta.dir, "../src")
const imports = [...(await readFile(path.join(source, "main.tsx"), "utf8")).matchAll(/import "(\.\/[^"\n]+\.css)"/g)]
const css = (await Promise.all(imports.map(([, file]) => readFile(path.resolve(source, file!), "utf8")))).join("\n")
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement, useLayoutEffect, useRef, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import JurisdictionSelect from ${JSON.stringify(path.join(source, "components/JurisdictionSelect.tsx"))}
const params = new URLSearchParams(location.search)
const jurisdictions = Array.from({ length: 30 }, (_, index) => ({ code: "jx-" + index, name: "Jurisdiction " + index }))
function Harness() {
  const [selected, setSelected] = useState([])
  const dialog = useRef(null)
  useLayoutEffect(() => { if (params.has("native")) dialog.current.showModal() }, [])
  const content = createElement(JurisdictionSelect, { jurisdictions, selected, onChange: setSelected, align: params.get("align"), id: "jurisdiction-trigger" })
  return createElement("div", null,
    params.has("native") ? createElement("dialog", { ref: dialog, className: "harness-anchor", "aria-label": "Picker host" }, content) : createElement("div", { className: "harness-anchor" }, content),
    createElement("output", { id: "selection" }, selected.join(",")))
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser" })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch(request) {
  const url = new URL(request.url)
  if (url.pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (url.pathname === "/product.css") return new Response(css, { headers: { "content-type": "text/css" } })
  const edge = url.searchParams.get("edge") ?? "bottom-right"
  return new Response(`<!doctype html><html data-text="${url.searchParams.get("scale")}"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="stylesheet" href="/product.css"><style>
    .harness-anchor { position: fixed; margin: 0; padding: 12px; border: 0; max-width: calc(var(--viewport-inline) - 16px); ${edge.includes("top") ? "top: 8px; bottom: auto;" : "bottom: 8px; top: auto;"} ${edge.includes("left") ? "left: 8px; right: auto;" : "right: 8px; left: auto;"} }
  </style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage({ reducedMotion: "reduce", hasTouch: true })
  page.setDefaultTimeout(5000)
  const errors: string[] = []
  page.on("pageerror", error => errors.push(error.message))
  const evidence = process.env.DOCHAUS_EVIDENCE_DIR
  if (evidence) await mkdir(evidence, { recursive: true })
  let cases = 0
  for (const viewport of [{ width: 375, height: 667 }, { width: 768, height: 760 }, { width: 1440, height: 900 }, { width: 1440, height: 320 }]) {
    await page.setViewportSize(viewport)
    for (const scale of ["large", "compact"]) for (const native of [false, true]) for (const align of ["left", "right"]) for (const edge of ["bottom-right", "top-left"]) {
      const query = new URLSearchParams({ scale, align, edge, ...(native ? { native: "1" } : {}) })
      await page.goto(`${server.url.href}?${query}`)
      const trigger = page.locator("#jurisdiction-trigger")
      await expect(trigger).toHaveAttribute("aria-expanded", "false")
      await trigger.click()
      const popover = page.locator(".jx-popover")
      await expect(trigger).toHaveAttribute("aria-expanded", "true")
      assert.equal(await trigger.getAttribute("aria-controls"), await popover.getAttribute("id"))
      const search = popover.getByRole("textbox", { name: "Search jurisdictions", exact: true })
      await expect(search).toBeFocused()
      const bounds = await popover.boundingBox()
      const anchor = await trigger.boundingBox()
      assert.ok(bounds && anchor)
      const context = `${viewport.width}x${viewport.height}, ${scale}, ${native ? "native dialog" : "body"}, ${align}, ${edge}`
      assert.ok(bounds.x >= 7 && bounds.x + bounds.width <= viewport.width - 7, `horizontal clipping: ${context}; ${JSON.stringify(bounds)}`)
      assert.ok(bounds.y >= 7 && bounds.y + bounds.height <= viewport.height - 7, `vertical clipping: ${context}; ${JSON.stringify(bounds)}`)
      if (edge.includes("bottom")) assert.ok(bounds.y + bounds.height <= anchor.y, `picker should open above its lower anchor: ${context}`)
      else assert.ok(bounds.y >= anchor.y + anchor.height, `picker should open below its upper anchor: ${context}`)
      assert.equal(await popover.evaluate(node => Boolean(node.closest("dialog"))), native, "native modal must own its interactive portal")
      if (evidence && scale === "large" && native && align === "left" && edge === "bottom-right") {
        await page.screenshot({ path: path.join(evidence, `jurisdictions-${viewport.width}x${viewport.height}-large.png`) })
      }
      await search.fill("Jurisdiction 29")
      await popover.getByRole("checkbox").check()
      await expect(page.locator("#selection")).toHaveText("jx-29")
      await page.keyboard.press("Escape")
      await expect(popover).toHaveCount(0)
      await expect(trigger).toHaveAttribute("aria-expanded", "false")
      await expect(trigger).toBeFocused()
      if (native) await expect(page.getByRole("dialog", { name: "Picker host" })).toBeVisible()
      cases++
    }
    console.log(`PASS ${viewport.width}x${viewport.height}: compact/large, left/right alignment, upper/lower edges, body/native modal, selection and Escape restoration`)
  }
  // A keyboard-height change must keep the active search usable, and changing
  // width must clamp the same open picker rather than leave stale coordinates.
  await page.setViewportSize({ width: 375, height: 667 })
  await page.goto(`${server.url.href}?scale=large&align=left&edge=bottom-right&native=1`)
  await page.locator("#jurisdiction-trigger").click()
  const search = page.getByRole("textbox", { name: "Search jurisdictions", exact: true })
  for (const viewport of [{ width: 375, height: 360 }, { width: 768, height: 360 }]) {
    await page.setViewportSize(viewport)
    await expect(search).toBeFocused()
    await expect.poll(async () => {
      const bounds = await page.locator(".jx-popover").boundingBox()
      return Boolean(bounds && bounds.x >= 7 && bounds.y >= 7 && bounds.x + bounds.width <= viewport.width - 7 && bounds.y + bounds.height <= viewport.height - 7)
    }).toBe(true)
  }
  await search.fill("Jurisdiction 29")
  await page.getByRole("checkbox").check()
  await expect(page.locator("#selection")).toHaveText("jx-29")
  console.log("PASS an open native picker preserves search focus and containment after height and width changes")
  assert.deepEqual(errors, [])
  console.log(`PASS ${cases} jurisdiction picker cases with no browser errors`)
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
