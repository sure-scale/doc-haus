import { expect, test } from "bun:test"
import type { Matter } from "../src/api/ingest/matters"
import { groupClients, suggestClients } from "../src/lib/clients"

const matter = (id: string, client: string | undefined, created_at: number): Matter => ({ id, title: id, client, dir: `/tmp/${id}`, created_at })

test("spellings that differ only in case or spacing are one client under the most common spelling", () => {
  const groups = groupClients([
    matter("a", "northwind traders", 1),
    matter("b", "Northwind  Traders", 2),
    matter("c", " Northwind Traders ", 3),
    matter("d", "Contoso", 4),
    matter("e", undefined, 5),
    matter("f", "", 6),
  ])
  expect(groups.map((group) => [group.name, group.matters.map((m) => m.id)])).toEqual([
    ["Contoso", ["d"]],
    ["Northwind Traders", ["c", "b", "a"]],
  ])
})

test("suggestions put prefix matches before contains matches and leave out an exact match", () => {
  const names = ["Bank of Acme", "Acme Holdings", "Contoso", "acme", "Acme Ltd"]
  expect(suggestClients(names, "ac")).toEqual(["Acme Holdings", "acme", "Acme Ltd", "Bank of Acme"])
  expect(suggestClients(names, "ACME")).toEqual(["Acme Holdings", "Acme Ltd", "Bank of Acme"])
  expect(suggestClients(names, "zzz")).toEqual([])
  expect(suggestClients(names, "", 2)).toEqual(["Bank of Acme", "Acme Holdings"])
})
