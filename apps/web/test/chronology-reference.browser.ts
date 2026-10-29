// Real ChatPanel/Composer/cards; only native chat/drafting service ports are scripted.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { referenceEvent } from "./chronology-reference.fixture"
const { chromium, expect } = (await import(
  process.env.PLAYWRIGHT_MODULE ?? "@playwright/test"
)) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "chronology-reference-"))
const root = path.resolve(import.meta.dir, "../src")
const react = JSON.stringify(fileURLToPath(import.meta.resolve("react")))
const entry = path.join(directory, "harness.tsx")
await Bun.write(
  entry,
  `
import { createElement, useState } from ${react}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter, useLocation, useNavigate } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import ChatPanel from ${JSON.stringify(path.join(root, "components/ChatPanel.tsx"))}
import { ToastProvider } from ${JSON.stringify(path.join(root, "components/Toast.tsx"))}
import { chronologyChatReference } from ${JSON.stringify(path.join(root, "components/chronology/model.ts"))}
import ${JSON.stringify(path.join(root, "styles.css"))}
import ${JSON.stringify(path.join(root, "styles/design-system.css"))}
import ${JSON.stringify(path.join(root, "styles/design-conversation.css"))}
const event = ${JSON.stringify(referenceEvent)}
window.__messages=[]; window.__updates=[]; window.__mode='failure'; window.__activeRun=false
function App(){ const [reference, setReference]=useState(); const [matter,setMatter]=useState('fixture'); const [,bump]=useState(0); const route=useLocation(),navigate=useNavigate(); const chronology=new URLSearchParams(route.search).get('view')==='chronology'
return createElement('main',{style:{maxWidth:'850px',margin:'0 auto'}},
createElement('button',{onClick:()=>setReference(chronologyChatReference(matter,event))},'Attach event'),
createElement('button',{onClick:()=>setReference(chronologyChatReference('another-matter',event))},'Wrong matter'),
createElement('button',{onClick:()=>{window.__activeRun=!window.__activeRun;bump(x=>x+1)}},'Toggle drafting'),
createElement('div',{hidden:!chronology},createElement('p',null,'Chronology table'),createElement('button',{onClick:()=>navigate('/matter/fixture?view=chat&session=chat-session')},'Back to chat')),
createElement('div',{hidden:chronology},createElement(ChatPanel,{directory:'/fixture',matterId:matter,sessionID:'chat-session',initialPrompt:'',chronologyReference:reference,onChronologyReferenceSeated:()=>setReference(undefined),agent:'qa',available:new Set(['qa']),pinned:{name:'qa',label:'Ask',description:'Ask about evidence'},onAgentChange:()=>{},onSessionCreated:()=>{}}))) }
createRoot(document.getElementById('root')).render(createElement(MemoryRouter,{initialEntries:['/matter/fixture?view=chat&session=chat-session']},createElement(ToastProvider,null,createElement(App))))
`,
)
const build = await Bun.build({
  entrypoints: [entry],
  target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/engine", VITE_INGEST_URL: "/ingest" }) },
  plugins: [
    {
      name: "scripted-chat-ports",
      setup(builder) {
        builder.onLoad({ filter: /\/chat\/use-chat-actions\.ts$/ }, () => ({
          loader: "js",
          contents: `export function useChatActions(options){const state=options.state;return{onSend:async(message,restore,dispatched,admitted)=>{window.__messages.push(message);state.setInput('');state.setBusy(true);if(window.__mode==='failure'){state.setInput(restore);state.setBusy(false);return}if(window.__mode==='cancel'){state.setInput(message);state.setBusy(false);return}if(window.__mode==='full'){options.onDraftingStarted('owner',state.turns.length);state.setBusy(false);return}dispatched?.();await new Promise(resolve=>{window.__admit=()=>{admitted?.();state.setBusy(false);resolve(true)}});return true},onStop:()=>{state.setInput('  My question\\nwith a deliberate line break.  ');state.setBusy(false)},runWorkflow:async()=>{},resendFrom:async()=>{},onPermissionReply:async()=>{},chooseDraft:async()=>{},dismissDraftOffer:()=>{},upgradeQuickDraft:async()=>{},dismissQuickLimit:()=>{},runFullInstead:async()=>{}}}`,
        }))
        builder.onLoad({ filter: /\/chat\/use-session-events\.ts$/ }, () => ({
          loader: "js",
          contents: `export function useSessionEvents(){}`,
        }))
        builder.onLoad({ filter: /\/chat\/use-turn-settlement\.ts$/ }, () => ({
          loader: "js",
          contents: `export function useTurnSettlement(){return {}}`,
        }))
        builder.onLoad({ filter: /\/narration\/use-drafting-narration\.ts$/ }, () => ({
          loader: "js",
          contents: `export function useDraftingNarration(){return {visible:window.__activeRun,status:window.__activeRun?{owner_session_id:'other',status:'running',updated_at:'2034-09-09',stages:{},questions:[],open_items:[],in_flight:true,pending_requests:0}:undefined,onForeignPart:()=>{},acceptsSession:()=>false}}`,
        }))
        builder.onLoad({ filter: /\/drafting\/useRunActions\.ts$/ }, () => ({
          loader: "js",
          contents: `export default function useRunActions(){return {busy:false,error:'',update:async text=>{window.__updates.push(text);return window.__mode==='success'},resume:async()=>{},pause:async()=>{},stop:async()=>{}}}`,
        }))
        builder.onLoad({ filter: /\/drafting\/RunHeader\.tsx$/ }, () => ({
          loader: "js",
          contents: `export default function RunHeader(){return null}`,
        }))
      },
    },
  ],
})
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text(),
  css = (await build.outputs.find((output) => output.path.endsWith(".css"))?.text()) ?? ""
let accessibleEvent = structuredClone(referenceEvent)
let accessMode: "ok" | "missing" | "error" = "ok"
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch(request) {
    const route = new URL(request.url).pathname
    if (route === "/app.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (route === "/app.css") return new Response(css, { headers: { "content-type": "text/css" } })
    if (route === "/ingest/matters/fixture/chronology")
      return accessMode === "error"
        ? Response.json({ error: "Chronology access could not be confirmed." }, { status: 503 })
        : Response.json({ events: accessMode === "missing" ? [] : [accessibleEvent], proposals: [] })
    return new Response(
      '<html><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/app.css"><div id="root"></div><script type="module" src="/app.js"></script></html>',
      { headers: { "content-type": "text/html" } },
    )
  },
})
const browser = await chromium.launch({
  headless: true,
  ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}),
})
try {
  const page = await browser.newPage({ viewport: { width: 1100, height: 900 } }),
    errors: string[] = []
  page.on("pageerror", (error) => errors.push(error.message))
  page.setDefaultTimeout(5000)
  await page.goto(server.url.href)
  const message = page.getByRole("textbox", { name: "Message" }),
    draft = "  My question\nwith a deliberate line break.  ",
    card = page.getByRole("article", { name: "Chronology reference: Reported delivery" })
  await message.fill(draft)
  await page.getByRole("button", { name: "Attach event", exact: true }).click()
  await expect(card).toBeVisible()
  await expect(message).toHaveValue(draft)
  await page.getByRole("button", { name: "Attach event", exact: true }).click()
  await expect(card).toHaveCount(1)
  await card.getByText("Included evidence", { exact: true }).click()
  await expect(card).toContainText("Supplied chronology.md")
  await expect(card).toContainText("stale")
  await expect(card).toContainText("underlying original source")
  await expect(card).not.toContainText("private-internal-version")
  await card.getByRole("link", { name: "Return to chronology" }).click()
  await expect(page.getByText("Chronology table", { exact: true })).toBeVisible()
  await page.getByRole("button", { name: "Back to chat" }).click()
  await expect(message).toHaveValue(draft)
  await expect(card).toBeVisible()
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await expect(message).toHaveValue(draft)
  await expect(card).toBeVisible()
  await expect.poll(() => page.evaluate(() => (window as any).__messages.length)).toBe(1)
  const actual = await page.evaluate(() => (window as any).__messages.at(-1))
  assert.match(actual, /The carrier recorded delivery/)
  assert.match(actual, /quotation verified; stale; supplied-chronology/)
  assert.match(actual, /My question\nwith a deliberate line break\./)
  await page.evaluate(() => {
    ;(window as any).__mode = "cancel"
  })
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await expect.poll(() => page.evaluate(() => (window as any).__messages.length)).toBe(2)
  await expect(message).toHaveValue(draft)
  await expect(card).toBeVisible()
  await page.evaluate(() => {
    ;(window as any).__mode = "hold"
  })
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await expect(card).toBeVisible()
  await expect(card.getByRole("button", { name: "Remove chronology reference: Reported delivery" })).toBeDisabled()
  await expect.poll(() => page.evaluate(() => typeof (window as any).__admit)).toBe("function")
  await page.evaluate(() => {
    ;(window as any).__admit()
  })
  await expect(card).toHaveCount(0)
  await page.getByRole("button", { name: "Wrong matter", exact: true }).click()
  await expect(card).toHaveCount(0)
  await page.getByRole("button", { name: "Attach event", exact: true }).click()
  await expect(message).toHaveValue("Please explain the evidence for this chronology entry.")
  await page.setViewportSize({ width: 375, height: 780 })
  await expect(card).toBeVisible()
  const bounds = await card.boundingBox()
  assert.ok(bounds && bounds.x >= 0 && bounds.x + bounds.width <= 375)
  await card.getByRole("button", { name: "Remove chronology reference: Reported delivery" }).click()
  await expect(card).toHaveCount(0)
  await page.getByRole("button", { name: "Attach event", exact: true }).click()
  await message.fill(draft)
  await page.getByRole("button", { name: "Toggle drafting", exact: true }).click()
  await page.evaluate(() => {
    ;(window as any).__mode = "failure"
  })
  await page.getByRole("button", { name: "Send drafting update", exact: true }).click()
  await expect(card).toBeVisible()
  await expect(message).toHaveValue(draft)
  await expect.poll(() => page.evaluate(() => (window as any).__updates.length)).toBe(1)
  assert.match(await page.evaluate(() => (window as any).__updates.at(-1)), /The carrier recorded delivery/)
  await page.evaluate(() => {
    ;(window as any).__mode = "success"
  })
  await page.getByRole("button", { name: "Send drafting update", exact: true }).click()
  await expect(card).toHaveCount(0)
  await expect(message).toHaveValue("")
  await page.getByRole("button", { name: "Toggle drafting", exact: true }).click()
  await page.getByRole("button", { name: "Attach event", exact: true }).click()
  await message.fill(draft)
  const priorSends = await page.evaluate(() => (window as any).__messages.length)
  accessibleEvent.accounts[0]!.evidence[0]!.excerpt = ""
  accessibleEvent.accounts[0]!.evidence[0]!.freshness = "unavailable"
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await expect(card).toContainText("Stale or unavailable support")
  await card.getByText("Included evidence", { exact: true }).click()
  await expect(card).toContainText("Source passage unavailable.")
  await expect(card.locator("blockquote")).toHaveCount(0)
  await expect(message).toHaveValue(draft)
  assert.equal(
    await page.evaluate(() => (window as any).__messages.length),
    priorSends,
    "revoked cached text stops this send",
  )
  await expect(
    page.getByText(
      "The chronology evidence or review status changed. Updated references are shown; review them before sending.",
      { exact: true },
    ),
  ).toBeVisible()
  accessMode = "error"
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await expect(page.getByText("Chronology access could not be confirmed.", { exact: true })).toBeVisible()
  await expect(card).toBeVisible()
  await expect(message).toHaveValue(draft)
  accessMode = "missing"
  await page.getByRole("button", { name: "Send message", exact: true }).click()
  await expect(card).toHaveCount(0)
  await expect(message).toHaveValue(draft)
  assert.equal(await page.evaluate(() => (window as any).__messages.length), priorSends)
  assert.deepEqual(errors, [])
  console.log(
    "Chronology references: visible evidence, exact draft preservation, scoped dedupe, navigation, failure/cancel, admission, removal, narrow bounds and drafting updates passed.",
  )
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
