// Run from apps/web with bun test/review-grid-presence.browser.ts.
// Exercise real BroadcastChannel delivery in two browser tabs.
import assert from "node:assert/strict"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const dir = await mkdtemp(path.join(os.tmpdir(), "review-presence-"))
const entry = path.join(dir, "entry.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import usePresence from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/useReviewGridPresence.ts"))}
function App() {
  const other = usePresence("/matters/example", new URL(location.href).searchParams.get("matter") || "one")
  return createElement("div", { role: "status" }, other ? "Other window" : "Only window")
}
createRoot(document.getElementById("root")).render(createElement(App))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser" })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs[0]!.text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/app.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response('<div id="root"></div><script type="module" src="/app.js"></script>', { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const context = await browser.newContext()
  const first = await context.newPage()
  await first.goto(server.url.href)
  await first.getByText("Only window").waitFor()
  const otherMatter = await context.newPage()
  await otherMatter.goto(`${server.url.href}?matter=two`)
  await otherMatter.getByText("Only window").waitFor()
  assert.equal(await first.getByText("Only window").count(), 1)
  const second = await context.newPage()
  await second.goto(server.url.href)
  await first.getByText("Other window").waitFor()
  await second.getByText("Other window").waitFor()
  assert.equal(await otherMatter.getByText("Only window").count(), 1)
  await second.close()
  await first.getByText("Only window").waitFor()
  console.log("PASS presence appears in both same-matter tabs, is scoped to the matter, and clears on close")
  await context.close()
} finally {
  await browser.close()
  server.stop(true)
  await rm(dir, { recursive: true, force: true })
}
