import { afterAll, describe, expect, test } from "bun:test"
import { Database, type QueryResult } from "@ladybugdb/core"
import { existsSync, mkdtempSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { matterGraphPath, openMatterGraph } from "./ladybug"

const workspace = mkdtempSync(path.join(tmpdir(), "dochaus-ladybug-"))
afterAll(() => rmSync(workspace, { recursive: true, force: true }))

describe("matter graph native boundary", () => {
  test("loads the exact qualified Ladybug version", () => {
    expect(Database.getVersion()).toBe("0.19.1")
  })

  test("requires an explicit existing matter directory", async () => {
    expect(() => matterGraphPath("relative-matter")).toThrow("absolute matter directory")
    await expect(openMatterGraph(path.join(workspace, "missing"))).rejects.toThrow("existing matter directory")
  })

  test("creates, traverses, updates, deletes, closes, and reopens one matter graph", async () => {
    const matter = await createMatter("lifecycle")
    const graph = await openMatterGraph(matter)
    expect(graph.databasePath).toBe(path.join(realpathSync(matter), ".dochaus", "ladybug", "matter.lbdb"))

    await graph.connection.query("CREATE NODE TABLE Item(id STRING PRIMARY KEY, name STRING)")
    await graph.connection.query("CREATE REL TABLE Links(FROM Item TO Item, kind STRING)")
    await graph.connection.query(`CREATE (:Item {id: "a", name: "Alpha"})`)
    await graph.connection.query(`CREATE (:Item {id: "b", name: "Beta"})`)
    await graph.connection.query(
      `MATCH (a:Item {id: "a"}), (b:Item {id: "b"}) CREATE (a)-[:Links {kind: "supports"}]->(b)`,
    )
    expect(await rows(await graph.connection.query("MATCH (a:Item)-[r:Links]->(b:Item) RETURN a.id, r.kind, b.id")))
      .toEqual([{ "a.id": "a", "r.kind": "supports", "b.id": "b" }])

    await graph.connection.query(`MATCH (item:Item {id: "b"}) SET item.name = "Updated"`)
    expect(await rows(await graph.connection.query(`MATCH (item:Item {id: "b"}) RETURN item.name`))).toEqual([
      { "item.name": "Updated" },
    ])
    await graph.connection.query(`MATCH (item:Item {id: "a"}) DETACH DELETE item`)
    expect(await rows(await graph.connection.query("MATCH (item:Item) RETURN item.id ORDER BY item.id"))).toEqual([
      { "item.id": "b" },
    ])
    await graph.close()

    expect(existsSync(graph.databasePath)).toBe(true)
    const reopened = await openMatterGraph(matter)
    expect(await rows(await reopened.connection.query("MATCH (item:Item) RETURN item.id, item.name"))).toEqual([
      { "item.id": "b", "item.name": "Updated" },
    ])
    await reopened.close()
  })

  test("isolates matters and enforces read-only opens", async () => {
    const first = await createMatter("first")
    const second = await createMatter("second")
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
  })

  test("does not create a missing sidecar during a read-only open", async () => {
    const matter = await createMatter("read-only-missing")
    await expect(openMatterGraph(matter, { readOnly: true })).rejects.toThrow("does not exist")
    expect(existsSync(path.join(matter, ".dochaus"))).toBe(false)
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
