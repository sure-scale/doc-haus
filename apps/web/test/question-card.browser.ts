// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/question-card.browser.ts
// Uses the real question dock, card and action hook; only HTTP is mocked. No live run is touched.
// Set QUESTION_CARD_SCREENSHOTS=<directory> to save the open card.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-question-card-"))
const shots = process.env.QUESTION_CARD_SCREENSHOTS
const options = [
  { id: "o1", text: "No remedy" },
  { id: "o2", text: "Liquidated damages at 1% a week" },
  { id: "o3", text: "A right to terminate after 30 days" },
]
const recommendation = {
  option: "o2",
  rationale: "The buyer asked for a remedy that does not need proof of loss.",
  basis: [
    { source_id: "request", quote: "a remedy that does not need proof of loss" },
    { source_id: "Supply terms.docx", quote: "Delivery is due within 14 days." },
  ],
}
const asked = blockedQuestion.questions[0]!
const status = {
  ...blockedQuestion,
  queued_answer_ids: [] as string[],
  questions: [
    { ...asked, id: "q-remedy", kind: "judgment", advice: { options, recommendation } },
    { ...asked, id: "q-storage", question: "Who bears the cost of storage?", kind: "judgment", advice: { options, recommendation } },
  ],
}
const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import QuestionDock from ${source("components/drafting/QuestionDock.tsx")}
import useRunActions from ${source("components/drafting/useRunActions.ts")}
import { openQuestions } from ${source("components/drafting/model.ts")}
// The dock is wired as ChatPanel wires it.
function Harness() {
  const [run, setRun] = useState(${JSON.stringify(status)})
  const actions = useRunActions("matter-q", run, setRun)
  const asked = openQuestions(run)
  return createElement(QuestionDock, {
    matterId: "matter-q", ownerSessionID: run.owner_session_id, questions: asked, answered: run.questions.length - asked.length,
    titles: new Map(), busy: actions.busy, error: actions.error, errorFor: actions.errorFor,
    onSend: (question, answer) => actions.answer(question.id, answer),
    onChoose: (question, optionID) => actions.choose(question.id, optionID),
    onDecide: (question) => actions.decide(question.id),
  })
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://questions.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs[0]!.text()
const css = (await Promise.all(["styles.css", "styles/primitives.css", "styles/drafting.css", "styles/drafting-items.css", "styles/design-system.css", "styles/design-conversation.css"]
  .map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><meta charset="utf-8"><style>body{font-family:sans-serif;padding:16px;max-width:760px}${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const posts: { url: string; body: unknown }[] = []
  await page.route("http://questions.invalid/**", async route => {
    const url = route.request().url()
    if (url.includes("/answer-drafts/")) return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ answers: {}, revision: 0 }) })
    posts.push({ url, body: route.request().postDataJSON() })
    const id = decodeURIComponent(url.slice(url.lastIndexOf("/") + 1))
    status.queued_answer_ids = [...status.queued_answer_ids, id]
    return route.fulfill({ status: 202, contentType: "application/json", body: JSON.stringify({ profile: "lean", mode: "queued", request_id: `r${posts.length}`, status }) })
  })
  await page.goto(server.url.href)
  await page.getByRole("button", { name: "Answer 2", exact: true }).click()
  const card = page.getByRole("dialog")
  const choice = card.getByRole("button", { name: /^Liquidated damages at 1% a week/ })
  // Enabled once saved answers load; nothing is chosen for the lawyer.
  await choice.and(page.locator(":enabled")).waitFor()
  const group = card.getByRole("group", { name: "Your answer" })
  assert.deepEqual(await group.locator(".question-option-text").allInnerTexts(), options.map(option => option.text))
  assert.equal(await group.getByText("Recommended", { exact: true }).count(), 1)
  assert.equal(await card.locator('[aria-pressed="true"], :focus.question-option').count(), 0)
  await card.getByText("Your instructions", { exact: true }).waitFor()
  await card.getByLabel("Other", { exact: true }).waitFor()
  if (shots) await page.screenshot({ path: path.join(shots, "question-card.png"), fullPage: true })

  await choice.click()
  await card.getByText("Who bears the cost of storage?").waitFor()
  assert.deepEqual(posts[0], { url: "http://questions.invalid/matters/matter-q/lean/questions/q-remedy", body: { option_id: "o2" } })
  console.log("PASS choosing an option posts exactly its option_id and moves to the next question")

  const decide = card.getByRole("button", { name: "Decide for me", exact: true })
  await decide.click()
  await card.waitFor({ state: "detached" })
  assert.deepEqual(posts[1], { url: "http://questions.invalid/matters/matter-q/lean/questions/q-storage", body: { decide_for_me: true } })
  assert.equal(posts.length, 2)
  console.log("PASS Decide for me posts exactly decide_for_me: true, and the dock closes once nothing is asked")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
