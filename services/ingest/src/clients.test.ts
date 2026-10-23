import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const root = mkdtempSync(path.join(tmpdir(), "dochaus-clients-"))
process.env.WORKSPACE_ROOT = root
const { clientRoutes } = await import("./clients")
const { createMatter, getMatter, listMatters } = await import("./matter")
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

const patch = (id: number, body: unknown) =>
  app.request(`/clients/${id}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
const rename = (body: unknown) =>
  app.request("/clients/rename", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
const byName = async (name: string) => (await list()).find((c) => c.name === name)
type Edited = { client: { id: number; name: string; reference?: string; created_at: number }; matters: number; merged: boolean }

test("editing a client's name and number renames it on every matter that names it", async () => {
  const added = await (await post({ name: "Tailspin Toys", reference: "C-1" })).json()
  const lease = createMatter("Lease", undefined, undefined, undefined, "Tailspin Toys")
  const loan = createMatter("Loan", undefined, undefined, undefined, "tailspin  toys.")
  const other = createMatter("Other", undefined, undefined, undefined, "Fourth Coffee")
  const res = await patch(added.id, { name: " Tailspin Toys Ltd ", reference: "C-2" })
  expect(res.status).toBe(200)
  const edited = (await res.json()) as Edited
  expect(edited).toMatchObject({ client: { id: added.id, name: "Tailspin Toys Ltd", reference: "C-2" }, matters: 2, merged: false })
  expect([lease, loan, other].map((m) => getMatter(m.id).client)).toEqual(["Tailspin Toys Ltd", "Tailspin Toys Ltd", "Fourth Coffee"])
  // A blank number clears it; a case-only rename keeps the record.
  expect(((await (await patch(added.id, { name: "TAILSPIN TOYS LTD", reference: "" })).json()) as Edited).client).toEqual({ id: added.id, name: "TAILSPIN TOYS LTD", created_at: added.created_at })
  expect(getMatter(lease.id).client).toBe("TAILSPIN TOYS LTD")
  for (const body of [{ name: "" }, { name: "x".repeat(201) }, { name: "Ok", reference: 4 }]) expect((await patch(added.id, body)).status).toBe(400)
  expect((await patch(99999, { name: "Nobody" })).status).toBe(404)
})

test("renaming a client to another client's name merges it, keeping that client's name and number", async () => {
  const from = await (await post({ name: "Litware", reference: "L-9" })).json()
  const into = await (await post({ name: "Proseware Inc" })).json()
  const deal = createMatter("Deal", undefined, undefined, undefined, "Litware")
  const kept = createMatter("Kept", undefined, undefined, undefined, "Proseware Inc")
  const edited = (await (await patch(from.id, { name: "proseware inc.", reference: "L-9" })).json()) as Edited
  expect(edited).toMatchObject({ client: { id: into.id, name: "Proseware Inc", reference: "L-9" }, matters: 1, merged: true })
  expect(await byName("Litware")).toBeUndefined()
  expect((await list()).filter((c) => c.name.startsWith("Proseware"))).toEqual([{ ...into, reference: "L-9" }])
  expect([deal, kept].map((m) => getMatter(m.id).client)).toEqual(["Proseware Inc", "Proseware Inc"])
})

test("editing a client named only on matters stores it and renames those matters", async () => {
  const first = createMatter("First", undefined, undefined, undefined, "Adventure Works")
  const second = createMatter("Second", undefined, undefined, undefined, "adventure works")
  expect(await byName("Adventure Works")).toBeUndefined()
  const edited = (await (await rename({ from: "adventure works", name: "Adventure Works LLP", reference: "A-1" })).json()) as Edited
  expect(edited).toMatchObject({ client: { name: "Adventure Works LLP", reference: "A-1" }, matters: 2, merged: false })
  expect(await byName("Adventure Works LLP")).toEqual(edited.client)
  expect([first, second].map((m) => getMatter(m.id).client)).toEqual(["Adventure Works LLP", "Adventure Works LLP"])
  // A matter-only client renamed onto a stored one merges without a new record.
  createMatter("Third", undefined, undefined, undefined, "Blue Yonder")
  const count = (await list()).length
  expect(((await (await rename({ from: "Blue Yonder", name: "Adventure Works LLP" })).json()) as Edited).merged).toBe(true)
  expect((await list()).length).toBe(count)
  expect(listMatters().filter((m) => m.client === "Adventure Works LLP").map((m) => m.title).sort()).toEqual(["First", "Second", "Third"])
  expect((await rename({ name: "No source" })).status).toBe(400)
})
