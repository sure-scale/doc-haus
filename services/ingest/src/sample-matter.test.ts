import { afterAll, beforeAll, expect, test } from "bun:test"
import { existsSync, mkdtempSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { sampleMatterRoutes } from "./seed"
import { deleteMatter, listMatters } from "./matter"

let workspace = ""
let previous: string | undefined

beforeAll(() => {
  workspace = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sample-matter-")))
  previous = process.env.WORKSPACE_ROOT
  process.env.WORKSPACE_ROOT = workspace
})

afterAll(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  if (workspace) rmSync(workspace, { recursive: true, force: true })
})

const add = () => sampleMatterRoutes.request("/sample-matter", { method: "POST" })

test("POST /sample-matter seeds one sample matter, idempotently", async () => {
  // Two requests at once (a double click) share one seed.
  const [first, concurrent] = await Promise.all([add(), add()])
  expect([first.status, concurrent.status].sort()).toEqual([201, 201])
  const created = (await first.json()) as { matter: { id: string; sample?: boolean; dir: string }; created: boolean }
  expect(created.created).toBe(true)
  expect(created.matter.sample).toBe(true)
  expect(existsSync(path.join(created.matter.dir, "Letter of Engagement — Aldgate Mills.docx"))).toBe(true)
  expect(existsSync(path.join(created.matter.dir, ".dochaus", "legal.db"))).toBe(true)

  const again = await add()
  expect(again.status).toBe(200)
  expect(((await again.json()) as { matter: { id: string }; created: boolean })).toMatchObject({ created: false, matter: { id: created.matter.id } })
  expect(listMatters().map((matter) => matter.id)).toEqual([created.matter.id])

  // Deleted like any matter, it can be added again.
  deleteMatter(created.matter.id)
  const readded = await add()
  expect(readded.status).toBe(201)
  expect(listMatters()).toHaveLength(1)
}, 120_000)
