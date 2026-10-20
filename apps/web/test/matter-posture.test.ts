import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { MemoryRouter } from "react-router-dom"
import { isPosture } from "../../../dochaus/lib/drafting-preferences"
import { POSTURES } from "../src/api/ingest/preferences"
import type { MatterDetail } from "../src/api/ingest/matters"
import MatterHeader from "../src/components/matter/MatterHeader"

const noop = () => {}
const done = async () => {}
const header = (fields: Partial<MatterDetail>) =>
  renderToStaticMarkup(
    createElement(MemoryRouter, null, createElement(MatterHeader, {
      matter: { id: "matter-fixture", title: "Supply note", dir: "/tmp/matter-fixture", created_at: 0, documents: [], client: "Northwind Traders Ltd", ...fields },
      jurisdictions: [],
      editingRef: false, setEditingRef: noop, draftRef: "", setDraftRef: noop, commitRef: done,
      editingClient: false, setEditingClient: noop, draftClient: "", setDraftClient: noop, commitClient: done,
      renaming: false, setRenaming: noop, draftTitle: "", setDraftTitle: noop, commitRename: done,
      patch: done,
    })),
  )

test("the web offers the server's three Postures, and nothing else", () => {
  expect(POSTURES.map(row => row.value)).toEqual(["client-favorable", "balanced", "conservative"])
  for (const row of POSTURES) expect(isPosture(row.value)).toBe(true)
})

test("the matter header shows the matter's Posture, else the firm default", () => {
  const unset = header({})
  expect(unset).toContain('class="matter-posture is-default"')
  expect(unset).toMatch(/<option value="" selected="">Posture: firm default<\/option>/)
  const set = header({ posture: "conservative" })
  expect(set).toContain('class="matter-posture"')
  expect(set).toMatch(/<option value="conservative" selected="">Posture: Conservative<\/option>/)
  expect(set).not.toMatch(/<option value="" selected="">/)
})
