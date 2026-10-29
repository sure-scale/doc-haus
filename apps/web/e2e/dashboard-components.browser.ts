// Component contract and responsive proof. Synthetic data only, no product stack.
import { chromium, expect } from "@playwright/test"
import { mkdir, writeFile } from "node:fs/promises"
const base = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!base || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR")
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: { name: string; status: string }[] = []
const errors: string[] = []
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, reducedMotion: "reduce" })
page.on("pageerror", (error) => errors.push(error.message))
page.on("console", (message) => { if (message.type() === "error") errors.push(`${message.text()} ${message.location().url}`) })
const card = (id: string) => page.locator(`[data-component="${id}"]`)
async function open(view = "overview", width = 1440, height = 1000) {
  await page.setViewportSize({ width, height })
  await page.goto(`${base}/e2e/dashboard-components.html?view=${view}`)
  await expect(page.getByRole("heading", { name: "Dashboard", exact: true })).toBeVisible()
  await expect(page.locator(".dash-adaptive-natural").first().or(page.locator(".dash-setup-checklist"))).toBeVisible()
  await expect(page.getByRole("alert")).toHaveCount(0)
  await expect(page.locator(".component-brand img")).toHaveJSProperty("complete", true)
  expect(await page.locator(".component-brand img").evaluate((node) => (node as HTMLImageElement).naturalWidth)).toBeGreaterThan(0)
  await page.evaluate(() => document.fonts.ready)
  await page.waitForTimeout(80)
}
async function capture(name: string) {
  const overflow = await page.evaluate(() => Math.max(document.documentElement.scrollWidth - innerWidth, document.body.scrollWidth - document.body.clientWidth, document.querySelector("#root")!.scrollWidth - document.querySelector("#root")!.clientWidth))
  expect(overflow, `${name} horizontal overflow`).toBeLessThanOrEqual(0)
  await page.screenshot({ path: `${evidence}/${name}.png`, fullPage: !name.includes("ai-modal") })
}
async function check(name: string, run: () => Promise<void>) {
  await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`)
}
try {
  for (const [width, height] of [[320,640],[375,812],[768,1024],[1440,1000]]) {
    for (const view of ["overview", "catalog", "costs", "setup"]) await check(`${view} ${width}x${height}`, async () => {
      await open(view, width, height)
      if (view === "overview" && width === 1440) {
        const boxes = await Promise.all(["spend", "matter-cost", "draft-cost", "week"].map((id) => card(id).boundingBox()))
        expect(Math.max(...boxes.map((box) => box!.y)) - Math.min(...boxes.map((box) => box!.y))).toBeLessThan(1)
        const recent = await card("recent").boundingBox(), spend = await card("matter-spend").boundingBox()
        expect(Math.abs(recent!.y - spend!.y)).toBeLessThan(1)
      }
      if (view === "catalog") await expect(page.locator(".dash-adaptive-natural > [data-component]")).toHaveCount(15)
      await expect(page.getByRole("tab", { name: /Setup/ })).toContainText("1/5")
      await capture(`components-${view}-${width}x${height}`)
    })
  }
  await check("native composer, keyboard tabs and JSON count badges", async () => {
    await open()
    const plus = page.getByRole("button", { name: "New dashboard", exact: true })
    await plus.click()
    const dialog = page.getByRole("dialog", { name: "New dashboard", exact: true })
    const prompt = page.getByRole("textbox", { name: "Describe what you want", exact: true })
    await expect(prompt).toBeFocused()
    await expect(dialog.getByRole("button", { name: "Generate dashboard", exact: true })).toBeDisabled()
    await dialog.getByRole("button", { name: "Last month spend by client", exact: true }).click()
    await expect(prompt).toHaveValue("Last month spend by client")
    await expect(prompt).toBeFocused()
    await capture("components-ai-modal-1440x1000")
    await page.keyboard.press("Escape")
    await expect(dialog).not.toBeVisible()
    await expect(plus).toBeFocused()
    await page.getByRole("tab", { name: "Overview", exact: true }).focus()
    await page.keyboard.press("End")
    await expect(page.getByRole("tab", { name: /Setup/ })).toBeFocused()
    await page.keyboard.press("Home")
    await expect(page.getByRole("tab", { name: "Overview", exact: true })).toBeFocused()
    await page.evaluate(() => {
      const api = window.__dashboardComponentDemo
      const ui = api.snapshot("ui")!.data as Record<string, any>
      api.publish("ui", { ...ui, tabs: ui.tabs.map((tab: any) => tab.id === "setup" ? { ...tab, badge: { value: 2, total: 5 } } : tab.id === "overview" ? { ...tab, badge: { value: 7 } } : tab) })
    })
    await expect(page.getByRole("tab", { name: /Setup/ })).toContainText("2/5")
    await expect(page.getByRole("tab", { name: /Overview/ })).toContainText("7")
  })
  await check("adaptive cards use the gap beneath the shorter pinned card", async () => {
    await open("costs")
    const attention = await card("attention").boundingBox()
    const spend = await card("client-spend").boundingBox()
    const line = await card("conversations").boundingBox()
    expect(attention).toBeTruthy(); expect(spend).toBeTruthy(); expect(line).toBeTruthy()
    expect(Math.abs(line!.x - attention!.x)).toBeLessThan(1)
    expect(line!.y - attention!.y - attention!.height).toBeGreaterThanOrEqual(15)
    expect(line!.y - attention!.y - attention!.height).toBeLessThanOrEqual(18)
    expect(line!.y).toBeLessThan(spend!.y + spend!.height)
    await page.evaluate(() => {
      const api = window.__dashboardComponentDemo
      const value = api.snapshot("practice")!.data as Record<string, any>
      api.publish("practice", { ...value, attention: [...value.attention, { ...value.attention[0], id: "another", title: "Another question waiting" }] })
    })
    await expect(card("attention").locator("li")).toHaveCount(4)
    await page.waitForTimeout(80)
    const after = await card("attention").boundingBox()
    const lineAfter = await card("conversations").boundingBox()
    expect(lineAfter!.y).toBeGreaterThan(line!.y)
    expect(lineAfter!.y - after!.y - after!.height).toBeLessThanOrEqual(18)
  })
  await check("dynamically registered source drives chart JSON and recovers from errors", async () => {
    await open("catalog", 375, 812)
    await page.evaluate(() => {
      const api = window.__dashboardComponentDemo
      api.replaceSource("agent-query", { result: { total: 10, rows: [{ key: "a", label: "Live client A", value: 6 }, { key: "b", label: "Live client B", value: 4 }] } })
      api.setDefinitions([{ id: "dynamic", kind: "chart", frame: { title: "Live generated query", icon: "spend", size: "wide" }, props: { form: "bars", label: "Live query", format: "money", showShares: true }, bindings: { total: { source: "agent-query", path: "/result/total" }, points: { source: "agent-query", path: "/result/rows" } } }])
    })
    await expect(card("dynamic").locator(".dash-chart-total")).toHaveText("$10.00")
    await expect(card("dynamic")).toContainText("60%")
    await page.evaluate(() => window.__dashboardComponentDemo.publish("agent-query", { result: { total: 20, rows: [{ key: "c", label: "Live client C", value: 20 }] } }))
    await expect(card("dynamic").locator(".dash-chart-total")).toHaveText("$20.00")
    await expect(card("dynamic")).toContainText("Live client C")
    await expect(card("dynamic")).not.toContainText("Live client A")
    await capture("components-realtime-query-375x812")
    await page.evaluate(() => window.__dashboardComponentDemo.fail("agent-query"))
    await expect(card("dynamic").getByRole("alert")).toContainText("Could not update")
    await expect(card("dynamic").getByRole("alert")).toContainText("Last updated")
    await page.evaluate(() => window.__dashboardComponentDemo.loading("agent-query"))
    await expect(card("dynamic").getByRole("status", { name: "Loading live data" })).toBeVisible()
    await page.evaluate(() => window.__dashboardComponentDemo.publish("agent-query", { result: { total: 5, rows: [{ label: "Recovered", value: 5 }] } }))
    await expect(card("dynamic").locator(".dash-chart-total")).toHaveText("$5.00")
    await expect(page.getByRole("alert")).toHaveCount(0)
    await page.evaluate(() => window.__dashboardComponentDemo.publish("agent-query", { result: { total: 5, rows: [{ label: "Invalid", value: -5 }] } }))
    await expect(card("dynamic").getByRole("alert")).toBeVisible()
    await page.evaluate(() => {
      const api = window.__dashboardComponentDemo
      const definition = api.definitions()[0]!
      api.setDefinition("dynamic", { ...definition, bindings: { points: { source: "constructor", path: "/data" } } })
    })
    await expect(card("dynamic").getByRole("alert")).toContainText("not connected")
  })
  await check("mobile composer Stop and Escape remain operable", async () => {
    await open("costs", 320, 640)
    await page.getByRole("button", { name: "New dashboard", exact: true }).click()
    const prompt = page.getByRole("textbox", { name: "Describe what you want", exact: true })
    await prompt.fill("Show weekly conversations")
    await page.evaluate(() => {
      const api = window.__dashboardComponentDemo
      api.publish("ui", { ...(api.snapshot("ui")!.data as Record<string, any>), working: true })
    })
    await expect(page.getByRole("button", { name: "Stop", exact: true })).toBeVisible()
    await page.getByRole("button", { name: "Stop", exact: true }).click()
    await expect(prompt).toBeFocused()
    await expect(prompt).toBeEnabled()
    await capture("components-ai-modal-320x640")
    await page.keyboard.press("Escape")
    await expect(page.getByRole("dialog")).not.toBeVisible()
  })
  await check("forced colors preserve component data marks", async () => {
    await open("costs")
    await page.emulateMedia({ forcedColors: "active" })
    const bar = card("client-spend").locator(".dash-chart-track > span").first()
    const paint = await bar.evaluate((node) => ({ fill: getComputedStyle(node).backgroundColor, track: getComputedStyle(node.parentElement!).backgroundColor }))
    expect(paint.fill).not.toBe(paint.track)
    await capture("components-forced-colors-1440x1000")
    await page.emulateMedia({ forcedColors: "none" })
  })
  expect(errors).toEqual([])
} finally {
  await writeFile(`${evidence}/component-results.json`, JSON.stringify({ results, errors }, null, 2) + "\n")
  await browser.close()
}
console.log(`PASS ${results.length} component scenarios`)
