import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { buildChecklistRefinementContext } from "./checklist-refinement-context"
const roots: string[] = [], workspace = process.env.WORKSPACE_ROOT
afterEach(() => { roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }));if (workspace === undefined) delete process.env.WORKSPACE_ROOT;else process.env.WORKSPACE_ROOT = workspace })
function fixture() {
  const dir = mkdtempSync(path.join(tmpdir(), "refinement-units-"));roots.push(dir)
  const put = (file: string, value: unknown) => { mkdirSync(path.dirname(path.join(dir, file)), { recursive: true });writeFileSync(path.join(dir, file), typeof value === "string" ? value : JSON.stringify(value)) }
  put(".dochaus/lean/plan.json", { schema: "dochaus.lean.plan/1", documents: [{ id: "d1", file: "agreement.docx", title: "Agreement", purpose: "Record terms", parts: [{ heading: "Terms", scope: "Terms" }] }] });put(".dochaus/lean/request.md", "Review original sources; preserve conditional proposals.\n")
  const texts = ["# Custody\n\nKeep records, unless the owner directs otherwise.\n\n", "# Returns\n\nPropose a receipt, subject to confirmation."]
  const text = texts.join("");put("sources/guide.txt", text + "\n");const document = { document: "guide.txt", extracted_text: "sources/guide.txt", characters: text.length, chunks: texts.map((quote, i) => ({ id: `guide-c${i + 1}`, heading_path: [quote.split("\n")[0]], start_char: i ? texts[0].length : 0, end_char: i ? text.length : texts[0].length })) }
  put("source-index.json", { sources: [document] });put("source-map.json", { documents: [document] })
  const item = { check: "Check source-supported concerns.", basis: "directive", source_refs: [{ document: "guide.txt", locator: "Custody", quote: "Keep records, unless the owner directs otherwise." }], qualifications: ["Preserve all conditions."], target: { document: "d1", role: "Agreement", treatment: "operative" } }
  const save = () => put(".dochaus/lean/checklist/extraction.json", { sources: [{ items: [{ id: "parent", item }] }] });save();return { dir, put, text, texts, item, document, save }
}
test("finite collection retains existing IDs, exact canonical coverage and adjacent context without current documents", () => {
  const f = fixture(), before = readFileSync(path.join(f.dir, ".dochaus/lean/checklist/extraction.json")), result = buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })
  expect(result.units.map(unit => unit.id)).toEqual(["guide-c1", "guide-c2"]);expect(result.units.map(unit => unit.passages[0].quote).join("")).toBe(f.text);expect(result.units[0].passages[0].quote).toContain("unless the owner directs otherwise")
  expect(result.units[0].passages[1].locator).toContain("adjacent boundary");expect(result.scope).toBe("source_collection");expect(result.instructions[0].document).toBe("request")
  f.put(".dochaus/lean/d1/body.docx", "GENERATED OUTPUT");expect(buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })).toEqual(result);expect(readFileSync(path.join(f.dir, ".dochaus/lean/checklist/extraction.json"))).toEqual(before)
})
test("local quote selects its region, cross-boundary and whole-source citations select finite units without copying a giant source", () => {
  const f = fixture();expect(buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "cited_sources" }).units.map(unit => unit.id)).toEqual(["guide-c1"])
  f.item.source_refs[0].quote = "owner directs otherwise.\n\n# Returns";f.save();expect(buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "cited_sources" }).units).toHaveLength(2)
  f.item.source_refs[0].quote = f.text;f.save();const result = buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "cited_sources" });expect(result.units.map(unit => unit.passages[0].quote)).toEqual(f.texts)
})
test("unknown source, altered quote, gapped map and missing map fail instead of dropping source coverage", () => {
  const f = fixture();f.item.source_refs[0].document = "missing.txt";f.save();expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "cited_sources" })).toThrow("unavailable")
  f.item.source_refs[0].document = "guide.txt";f.item.source_refs[0].quote = "Always retain every record.";f.save();expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "cited_sources" })).toThrow("not found")
  f.document.chunks[1].start_char++;f.put("source-map.json", { documents: [f.document] });expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })).toThrow("gapped")
  rmSync(path.join(f.dir, "source-map.json"));expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })).toThrow("unavailable")
})
test("actual answers remain separate; library guidance is bounded on paragraph/word boundaries with exact coverage", () => {
  const f = fixture();f.put(".dochaus/lean/questions.json", { questions: [{ id: "q1", question: "MODEL QUESTION", status: "answered", answer: "Keep the fallback conditional." }, { id: "q2", question: "OPEN MODEL QUESTION", status: "open", answer: "UNSETTLED" }] })
  process.env.WORKSPACE_ROOT = f.dir;const library = "# Guidance\n\n" + "Preserve conditional proposals and exceptions.\n\n".repeat(200)
  f.put("matter.json", { playbook: "playbook-unit-guidance" });f.put(".playbooks/playbook-unit-guidance/SKILL.md", library)
  const result = buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" }), units = result.units.filter(unit => unit.passages[0].document === "playbook:playbook-unit-guidance")
  expect(units.length).toBeGreaterThan(1);expect(units.every(unit => unit.passages[0].quote.length <= 4500)).toBe(true);expect(units.map(unit => unit.passages[0].quote).join("")).toBe(library);expect(result.instructions.map(ref => ref.quote)).toContain("Keep the fallback conditional.");expect(result.instructions.find(ref => ref.quote === "Keep the fallback conditional.")?.question).toBe("MODEL QUESTION");expect(result.instructions.filter(ref => ref.quote.includes("MODEL QUESTION"))).toEqual([]);expect(JSON.stringify(result)).not.toContain("OPEN MODEL QUESTION");expect(JSON.stringify(result)).not.toContain("UNSETTLED")
  rmSync(path.join(f.dir, ".playbooks/playbook-unit-guidance/SKILL.md"));expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })).toThrow("unavailable")
})

test("duplicated chunk identities and missing canonical source text do not silently lose units", () => {
  const f = fixture();f.document.chunks[1].id = f.document.chunks[0].id;f.put("source-map.json", { documents: [f.document] })
  expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })).toThrow("not unique")
  rmSync(path.join(f.dir, "sources/guide.txt"));expect(() => buildChecklistRefinementContext(f.dir, { parent_id: "parent", scope: "source_collection" })).toThrow()
})
