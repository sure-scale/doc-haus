// Focused browser regression: the real combobox, product CSS and scripted clients.
import assert from "node:assert/strict"
import { mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
const { chromium, expect } = (await import(
  process.env.PLAYWRIGHT_MODULE ?? "@playwright/test"
)) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-client-combobox-"))
const source = path.resolve(import.meta.dir, "../src")
const imports = [...(await readFile(path.join(source, "main.tsx"), "utf8")).matchAll(/import "(\.\/[^"\n]+\.css)"/g)]
const css = (await Promise.all(imports.map(([, file]) => readFile(path.resolve(source, file!), "utf8")))).join("\n")
const entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement as h, useLayoutEffect, useRef, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ClientCombobox from ${JSON.stringify(path.join(source, "components/ClientCombobox.tsx"))}
const params = new URLSearchParams(location.search)
function Harness() {
 const [value, setValue] = useState(""); const [keys, setKeys] = useState(""); const dialog = useRef(null)
 useLayoutEffect(() => { if (params.has("native")) dialog.current.showModal() }, [])
 const field = h(ClientCombobox, { value, onChange: setValue, clients: Array.from({ length: 8 }, (_, i) => "Northstar Studio " + i), "aria-label": "Client", onKeyDown: e => { setKeys(e.key); if (e.key === "Escape") e.preventDefault() } })
 return h("div", null, params.has("native") ? h("dialog", { ref: dialog, className: "harness-anchor", "aria-label": "Client host" }, field) : h("div", { className: "harness-anchor" }, field), h("output", { id: "keys" }, keys))
}
createRoot(document.getElementById("root")).render(h(Harness))
`,
)
const build = await Bun.build({
  entrypoints: [entry],
  target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://clients.invalid" }) },
})
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  fetch(request) {
    const url = new URL(request.url)
    if (url.pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (url.pathname === "/product.css") return new Response(css, { headers: { "content-type": "text/css" } })
    return new Response(
      `<!doctype html><html data-text="${url.searchParams.get("scale")}"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="stylesheet" href="/product.css"><style>.harness-anchor { position:fixed; margin:0; padding:12px; border:0; width:176px; max-width:calc(var(--viewport-inline) - 16px); ${url.searchParams.get("edge") === "top" ? "top:8px;left:8px;bottom:auto;right:auto;" : "bottom:8px;right:8px;top:auto;left:auto;"} } .harness-anchor input { width:100%; }</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`,
      { headers: { "content-type": "text/html" } },
    )
  },
})
const browser = await chromium.launch({
  headless: true,
  ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}),
})
try {
  const page = await browser.newPage({ reducedMotion: "reduce" })
  page.setDefaultTimeout(5000)
  await page.route("http://clients.invalid/**", (route) =>
    route.fulfill({ status: 200, contentType: "application/json", body: "[]" }),
  )
  const errors: string[] = []
  page.on("pageerror", (error) => errors.push(error.message))
  const field = page.getByRole("combobox", { name: "Client", exact: true })
  const list = page.getByRole("listbox", { name: "Existing clients", exact: true })
  async function contained(width: number, height: number, context: string) {
    const b = await list.boundingBox()
    assert.ok(b && b.width > 0 && b.height > 0)
    assert.ok(b.x >= 7 && b.x + b.width <= width - 7, `horizontal clipping: ${context}; ${JSON.stringify(b)}`)
    assert.ok(b.y >= 7 && b.y + b.height <= height - 7, `vertical clipping: ${context}; ${JSON.stringify(b)}`)
    return b
  }
  let cases = 0
  for (const viewport of [
    { width: 375, height: 667 },
    { width: 768, height: 760 },
    { width: 1440, height: 900 },
    { width: 900, height: 320 },
  ]) {
    await page.setViewportSize(viewport)
    for (const scale of ["large", "compact"])
      for (const native of [false, true])
        for (const edge of ["bottom", "top"]) {
          await page.goto(
            `${server.url.href}?${new URLSearchParams({ scale, edge, ...(native ? { native: "1" } : {}) })}`,
          )
          await field.fill("Northstar")
          await expect(field).toBeFocused()
          await expect(list.getByRole("option")).toHaveCount(9)
          const b = await contained(
            viewport.width,
            viewport.height,
            `${JSON.stringify(viewport)},${scale},${native},${edge}`,
          )
          const a = (await field.boundingBox())!
          assert.ok(
            edge === "bottom" ? b.y + b.height <= a.y : b.y >= a.y + a.height,
            `popup chooses available side: ${JSON.stringify({ viewport, scale, native, edge, b, a })}`,
          )
          assert.equal(await list.evaluate((node) => Boolean(node.closest("dialog"))), native)
          await page.keyboard.press("ArrowDown")
          await page.keyboard.press("Enter")
          await expect(field).toHaveValue("Northstar Studio 0")
          await expect(list).toHaveCount(0)
          await expect(page.locator("#keys")).toHaveText("")
          await field.fill("Northstar")
          await page.keyboard.press("Escape")
          await expect(list).toHaveCount(0)
          await expect(field).toBeFocused()
          await expect(page.locator("#keys")).toHaveText("")
          await page.keyboard.press("Escape")
          await expect(page.locator("#keys")).toHaveText("Escape")
          if (native) await expect(page.getByRole("dialog", { name: "Client host" })).toBeVisible()
          cases++
        }
    console.log(
      `PASS client popup ${viewport.width}x${viewport.height}: upper/lower, compact/large, body/modal, selection and Escape`,
    )
  }
  await page.setViewportSize({ width: 375, height: 667 })
  await page.goto(`${server.url.href}?scale=large&edge=bottom&native=1`)
  await field.fill("Northstar")
  for (const v of [
    { width: 375, height: 360 },
    { width: 768, height: 360 },
  ]) {
    await page.setViewportSize(v)
    await expect(field).toBeFocused()
    await expect(list).toBeVisible()
    await expect
      .poll(async () => {
        const b = await list.boundingBox()
        return Boolean(b && b.x >= 7 && b.y >= 7 && b.x + b.width <= v.width - 7 && b.y + b.height <= v.height - 7)
      })
      .toBe(true)
  }
  await list.getByRole("option", { name: "Northstar Studio 7", exact: true }).click()
  await expect(field).toHaveValue("Northstar Studio 7")
  await expect(field).toBeFocused()
  assert.deepEqual(errors, [])
  console.log(`PASS ${cases} client combobox placement cases and open viewport resize; no browser errors`)
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
