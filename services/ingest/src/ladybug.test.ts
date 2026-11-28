import { afterAll, describe, expect, test } from "bun:test"
import type { QueryResult } from "@ladybugdb/core"
import { existsSync, mkdtempSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { matterGraphPath, openMatterGraph } from "./ladybug"

const workspace = mkdtempSync(path.join(tmpdir(), "dochaus-ladybug-"))
afterAll(() => rmSync(workspace, { recursive: true, force: true }))

describe("matter graph native boundary", () => {
  test("requires an explicit existing matter directory", async () => {
    expect(() => matterGraphPath("relative-matter")).toThrow("absolute matter directory")
    await expect(openMatterGraph(path.join(workspace, "missing"))).rejects.toThrow("existing matter directory")
  })

  test("creates, persists, and reopens one matter graph at the matter sidecar path", async () => {
    const matter = createMatter("lifecycle")
    const graph = await openMatterGraph(matter)
    expect(graph.databasePath).toBe(path.join(realpathSync(matter), ".dochaus", "ladybug", "matter.lbdb"))
    await graph.connection.query("CREATE NODE TABLE Item(id STRING PRIMARY KEY)")
    await graph.connection.query(`CREATE (:Item {id: "a"})`)
    await graph.close()

    expect(existsSync(graph.databasePath)).toBe(true)
    const reopened = await openMatterGraph(matter)
    expect(await rows(await reopened.connection.query("MATCH (item:Item) RETURN item.id"))).toEqual([
      { "item.id": "a" },
    ])
    await reopened.close()
  })

  test("isolates matters and enforces read-only opens without creating a missing sidecar", async () => {
    const first = createMatter("first")
    const second = createMatter("second")
    const empty = createMatter("read-only-missing")
    await seed(first, "first")
    await seed(second, "second")

    const firstRead = await openMatterGraph(first, { readOnly: true })
    const secondRead = await openMatterGraph(second, { readOnly: true })
    expect(await rows(await firstRead.connection.query("MATCH (item:Item) RETURN item.id"))).toEqual([
      { "item.id": "first" },
    ])
    expect(await rows(await secondRead.connection.query("MATCH (item:Item) RETURN item.id"))).toEqual([
      { "item.id": "second" },
    ])
    await expect(firstRead.connection.query(`CREATE (:Item {id: "forbidden"})`)).rejects.toThrow()
    await firstRead.close()
    await secondRead.close()

    await expect(openMatterGraph(empty, { readOnly: true })).rejects.toThrow("does not exist")
    expect(existsSync(path.join(empty, ".dochaus"))).toBe(false)
  })
})

function createMatter(name: string) {
  return mkdtempSync(path.join(workspace, `${name}-`))
}

async function seed(matter: string, id: string) {
  const graph = await openMatterGraph(matter)
  await graph.connection.query("CREATE NODE TABLE Item(id STRING PRIMARY KEY)")
  await graph.connection.query(`CREATE (:Item {id: "${id}"})`)
  await graph.close()
}

async function rows(result: QueryResult | QueryResult[]) {
  if (Array.isArray(result)) throw new Error("Expected one Ladybug query result.")
  return result.getAll()
}
