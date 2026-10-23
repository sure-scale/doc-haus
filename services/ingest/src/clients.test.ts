import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const root = mkdtempSync(path.join(tmpdir(), "dochaus-clients-"))
process.env.WORKSPACE_ROOT = root
const { clientRoutes } = await import("./clients")
const { listMatters } = await import("./matter")
const app = clientRoutes()
afterAll(() => rmSync(root, { recursive: true, force: true }))

const post = (body: unknown) =>
  app.request("/clients", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
const list = async () => (await (await app.request("/clients")).json()) as { id: number; name: string; reference?: string }[]

test("a new client is stored with its number and listed, and the store is not a matter", async () => {
  const res = await post({ name: "  Northwind\tTraders Ltd ", reference: " C-0042 " })
  expect(res.status).toBe(201)
  const client = await res.json()
  expect(client).toMatchObject({ name: "Northwind Traders Ltd", reference: "C-0042" })
  expect(await list()).toEqual([client])
  expect(listMatters()).toEqual([])
})

test("adding a name that is already a client returns the existing client", async () => {
  const first = await (await post({ name: "Contoso Ltd" })).json()
  for (const name of ["contoso  ltd.", "CONTOSO LTD", "Contoso Ltd;"]) {
    const again = await post({ name, reference: "ignored" })
    expect(again.status).toBe(200)
    expect(await again.json()).toEqual(first)
  }
  // Composed and decomposed accents are one client too.
  const cafe = await (await post({ name: "Café Holdings" })).json()
  expect(await (await post({ name: "Café Holdings" })).json()).toEqual(cafe)
  expect((await list()).map((c) => c.name)).toEqual(["Café Holdings", "Contoso Ltd", "Northwind Traders Ltd"])
})

test("a blank or overlong name is refused and a client can be removed", async () => {
  for (const body of [{}, { name: "  " }, { name: 7 }, { name: "x".repeat(201) }]) expect((await post(body)).status).toBe(400)
  expect(((await (await post({ name: "" })).json()) as { error: string }).error).toBe("Name the client to add it.")
  const target = (await list()).find((c) => c.name === "Contoso Ltd")!
  expect((await app.request(`/clients/${target.id}`, { method: "DELETE" })).status).toBe(200)
  expect((await list()).some((c) => c.name === "Contoso Ltd")).toBe(false)
})
