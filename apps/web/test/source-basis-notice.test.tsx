import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import type { SourceDependencyStatus } from "../../../dochaus/lib/source-version-dependencies-contract"
import type { LeanDraftingStatus } from "../src/api/ingest/lean"
import type { DocumentVersion } from "../src/api/ingest/versions"
import SourceBasisNotice from "../src/components/SourceBasisNotice"
import VersionsPanel from "../src/components/DocumentVersions"
import RunHeader from "../src/components/drafting/RunHeader"
import type { RunActions } from "../src/components/drafting/useRunActions"

const stale: SourceDependencyStatus = { status: "stale", complete: true, sources: [{ documentId: "source", name: "Agreement.docx", versionId: "v1", versionNumber: 1, sha256: "a".repeat(64), currentVersionId: "v3", currentVersionNumber: 3, status: "newer-version" }] }
const unknown: SourceDependencyStatus = { status: "unknown", complete: false, sources: [] }
const actions: RunActions = { busy: false, pausing: false, stopping: false, error: "", errorFor: undefined, resuming: false, resume: async () => true, answer: async () => true, choose: async () => true, decide: async () => true, update: async () => true, refine: async () => true, pause: async () => true, stop: async () => true }

test("completed work retains its stale source-version notice and does not become a fresh approved result", () => {
  const status = { ...completed, source_dependencies: stale } as unknown as LeanDraftingStatus
  const html = renderToStaticMarkup(createElement(RunHeader, { status, actions, own: true, onOpen: () => {}, onShowDrafts: () => {} }))
  expect(html).toContain("Source material changed")
  expect(html).toContain("Agreement.docx · Built from version 1 · Current version 3")
  expect(html).toContain("latest available sources before relying on it")
  const legacy = renderToStaticMarkup(createElement(RunHeader, { status: { ...status, source_dependencies: unknown }, actions, own: true, onOpen: () => {}, onShowDrafts: () => {} }))
  expect(legacy).toContain("Source basis needs checking")
  expect(legacy).toContain("was not recorded or could not be verified")
})

test("derived versions show stale or legacy-unknown basis alongside their kept version metadata", () => {
  const version: DocumentVersion = { id: "derived1", number: 1, sha256: "f".repeat(64), ext: ".docx", bytes: 100, parent: null, derived_from: null, origin: "generated", actor: "doc.haus", run: "run", label: null, created_at: 1, current: true, source_dependencies: stale }
  const render = (basis: SourceDependencyStatus) => renderToStaticMarkup(createElement(VersionsPanel, { matterId: "matter", name: "Advice.docx", data: { name: "Advice.docx", document: { id: "derived", name: "Advice.docx", current_version: version.id }, versions: [{ ...version, source_dependencies: basis }], events: [], working_version: version.id }, unavailable: {}, onRestored: () => {}, onUnavailable: () => {} }))
  expect(render(stale)).toContain("Built from version 1 · Current version 3")
  expect(render(unknown)).toContain("Source basis needs checking")
  expect(render({ status: "current", complete: true, sources: [] })).not.toContain("source-basis-notice")
})

test("a newer linked received branch is labeled by its real source while the original source remains current", () => {
  const linked: SourceDependencyStatus = { ...stale, sources: [{ ...stale.sources[0]!, currentVersionId: "v1", currentVersionNumber: 1, status: "newer-received-branch", linkedReturn: { exchangeId: 44, sourceName: "Separate returned copy.docx", versionId: "other-v77", versionNumber: 77 } }] }
  const html = renderToStaticMarkup(createElement(SourceBasisNotice, { basis: linked }))
  expect(html).toContain("Agreement.docx · Built from version 1")
  expect(html).toContain("New linked return: Separate returned copy.docx · Version 77")
  expect(html).not.toContain("Source unavailable")
  expect(html).not.toContain("Current version 77")
})
