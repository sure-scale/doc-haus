// Two synthetic clients share a fixture server; no live API or account is used.
import { chromium, expect } from "@playwright/test"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const baseURL = process.env.PLAYWRIGHT_BASE_URL
if (!baseURL) throw new Error("Set PLAYWRIGHT_BASE_URL.")
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
let preferences = { ...DEFAULT_DRAFTING, practiceAreas: ["Corporate"] }
const writes: Record<string, unknown>[] = []
const errors: string[] = []

async function client() {
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 }, reducedMotion: "reduce" })
  await context.addInitScript(() => localStorage.setItem("dh.launch-notice.v1", new Date().toDateString()))
  const page = await context.newPage()
  page.on("pageerror", error => errors.push(error.message))
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || pathname.startsWith("/api/") || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === new URL(baseURL!).origin) return route.continue()
    const headers = { "access-control-allow-origin": new URL(baseURL!).origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers })
    const reply = (body: unknown) => route.fulfill({ status: 200, headers, contentType: "application/json", body: JSON.stringify(body) })
    if (pathname === "/preferences" && request.method() === "PUT") {
      const patch = request.postDataJSON() as Record<string, unknown>
      writes.push(patch)
      preferences = { ...preferences, ...patch }
      return reply(preferences)
    }
    if (pathname === "/preferences") return reply(preferences)
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply({ version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" })
    if (pathname === "/matters" || pathname === "/clients") return reply([])
    if (pathname === "/templates") return reply({ dir: "/fixture/templates", templates: [] })
    if (pathname === "/jurisdictions") return reply([{ code: "EW", name: "England and Wales" }])
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [] })
    if (pathname === "/api/model" || pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [] })
    return route.abort()
  })
  await page.goto(baseURL!)
  await page.getByRole("button", { name: "Settings", exact: true }).click()
  return { context, page, modal: page.getByRole("dialog", { name: "Settings" }) }
}
try {
  const first = await client()
  const second = await client()
  await first.modal.getByRole("tab", { name: "Matter defaults" }).click()
  await expect(first.modal.locator("#matter-practice-areas")).toHaveText(/Corporate/)
  await second.modal.getByRole("tab", { name: "Matter defaults" }).click()
  await second.modal.locator("#matter-practice-areas").click()
  await second.modal.getByRole("checkbox", { name: "Employment" }).check()
  await second.modal.getByRole("button", { name: "Save", exact: true }).click()
  await expect.poll(() => preferences.practiceAreas).toEqual(["Corporate", "Employment"])
  const count = writes.length
  await first.modal.locator("#matter-default-jurisdictions").click()
  await first.modal.getByRole("checkbox", { name: /England and Wales/ }).check()
  await first.modal.getByRole("button", { name: "Save", exact: true }).click()
  await expect.poll(() => first.page.evaluate(() => JSON.parse(localStorage.getItem("dochaus.prefs") ?? "{}").defaultJurisdictions)).toEqual(["EW"])
  expect(writes.length).toBe(count)
  expect(preferences.practiceAreas).toEqual(["Corporate", "Employment"])

  await first.modal.getByRole("tab", { name: "Drafting", exact: true }).click()
  await expect(first.modal.getByRole("tab", { name: "Drafting", exact: true })).toHaveAttribute("aria-selected", "true")
  await expect(first.modal.getByLabel("Attorney", { exact: true })).toBeVisible()
  await second.modal.locator("#matter-practice-areas").click()
  await second.modal.getByRole("checkbox", { name: "Tax", exact: true }).check()
  await second.modal.getByRole("button", { name: "Save", exact: true }).click()
  await expect.poll(() => preferences.practiceAreas).toEqual(["Corporate", "Employment", "Tax"])
  await first.modal.getByLabel("Attorney", { exact: true }).fill("Synthetic lawyer")
  await first.modal.getByRole("button", { name: "Save", exact: true }).click()
  await expect.poll(() => preferences.attorney).toBe("Synthetic lawyer")
  expect(writes.at(-1)).not.toHaveProperty("practiceAreas")
  expect(preferences.practiceAreas).toEqual(["Corporate", "Employment", "Tax"])
  expect(errors).toEqual([])
  await first.context.close()
  await second.context.close()
  console.log("PASS: jurisdiction and drafting saves preserve another client's practice-area changes")
} finally {
  await browser.close()
}
