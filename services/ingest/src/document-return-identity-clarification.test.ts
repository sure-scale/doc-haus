import { expect, spyOn, test } from "bun:test"
import { Hono } from "hono"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { documentExchangeRoutes } from "./document-exchange-routes"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { readCurrentVersion, readDocumentVersions } from "./document-versions"
import { openDb } from "./db"
import { checkIdentityClarification, clarifyIdentity } from "./document-return-identity-clarification"

const model = { providerID: "protocol-fixture", id: "configured-fast" }
const identity = "Sale Agreement Orchard C-47 between Cedar and Riverstone for Lot 14."
const excluded = "The transfer excludes the northern strip. The seller retains access, except for emergency use."
const included = "The transfer includes the northern strip. The seller retains no access rights."
const docx = (...text: string[]) => Packer.toBuffer(new Document({ sections: [{ children: text.map(item => new Paragraph(item)) }] }))
const block = (prompt: string, tag: string) => { const found = prompt.match(new RegExp(`<${tag} encoding="json">([\\s\\S]*?)<\\/${tag}>`)); return found ? JSON.parse(found[1]!) : undefined }
const address = (text: string) => ({ text, scope: "body", story: "body", paragraph: 2 })

test("clarification needs complete NEW exact passage evidence, notes and all required coverage; it cannot rewrite original different-document", () => {
  const triple = { part: 1, incoming: address(excluded), left: address(excluded), right: address(included) }
  const context = { versionId: "old", rivalVersionId: "current", side: "left" as const, triple, incomingQuotes: [identity], candidateQuotes: [identity] }
  const answer = { versionId: "old", relation: "same-document", reason: "The newly supplied complete qualification resolves the original identity omission.", limitations: [], evidence: [{ versionId: "old", incomingQuote: excluded, candidateQuote: excluded }] }
  const checked = { ...checkIdentityClarification(JSON.stringify(answer), context), model }
  expect(clarifyIdentity({ versionId: "old", relation: "uncertain" }, { requiredParts: 1, processedParts: 1 }, [], [checked]).relation).toBe("same-document")
  expect(clarifyIdentity({ versionId: "old", relation: "different-document" }, { requiredParts: 1, processedParts: 1 }, [], [checked]).relation).not.toBe("same-document")
  expect(clarifyIdentity({ versionId: "old", relation: "uncertain" }, { requiredParts: 2, processedParts: 1 }, [], [checked]).relation).toBe("uncertain")
  expect(clarifyIdentity({ versionId: "old", relation: "uncertain" }, { requiredParts: 1, processedParts: 1 }, ["A required external qualification is missing."], [checked]).relation).toBe("uncertain")
  const date = { part: 2, incoming: address("Completion is Wednesday."), left: address("Completion is Monday."), right: address("Completion is Tuesday.") }
  const neutralAnswer = { ...answer, relation: "neutral", reason: "This complete date-change passage is independently non-identifying; the separate property qualification still needs its own resolution.", evidence: [{ versionId: "old", incomingQuote: date.incoming.text, candidateQuote: date.left.text }] }
  const neutral = { ...checkIdentityClarification(JSON.stringify(neutralAnswer), { ...context, triple: date }), model }
  expect(clarifyIdentity({ versionId: "old", relation: "uncertain" }, { requiredParts: 1, processedParts: 1 }, [], [neutral]).relation).toBe("uncertain")
  expect(clarifyIdentity({ versionId: "old", relation: "uncertain" }, { requiredParts: 2, processedParts: 2 }, [], [checked, neutral]).relation).toBe("same-document")
  expect(clarifyIdentity({ versionId: "old", relation: "uncertain" }, { requiredParts: 2, processedParts: 2 }, ["The original material identity concern is still unaddressed."], [checked, neutral]).relation).toBe("uncertain")
  expect(() => checkIdentityClarification(JSON.stringify({ ...answer, evidence: [{ versionId: "old", incomingQuote: identity, candidateQuote: identity }] }), context)).toThrow("complete new passage pair")
  expect(() => checkIdentityClarification(JSON.stringify({ ...answer, evidence: [{ versionId: "old", incomingQuote: excluded, candidateQuote: "The transfer excludes the northern strip." }] }), context)).toThrow("exact supplied source")
  const withoutNotes = { ...answer } as Partial<typeof answer>; delete withoutNotes.limitations
  expect(() => checkIdentityClarification(JSON.stringify(withoutNotes), context)).toThrow("bounded evidence contract")
  expect(() => checkIdentityClarification(JSON.stringify({ ...answer, limitations: ["The schedule identifying the transaction remains unavailable."] }), context)).toThrow("bounded evidence contract")
  expect(() => checkIdentityClarification(JSON.stringify({ ...neutralAnswer, limitations: ["The identity-bearing cross-reference is missing."] }), { ...context, triple: date })).toThrow("bounded evidence contract")
})

test("uncertain identity reaches new contrasts and only complete explicit clarification can qualify a base", async () => {
  for (const mode of ["resolved", "still-uncertain", "external-unknown", "both-uncertain", "missing-coverage", "all-different", "source-drift"] as const) {
    const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-identity-pivot-"))
    let fetch: ReturnType<typeof spyOn> | undefined, contrasts = 0, clarifications = 0, reconciliations = 0, calls = 0, drifted = false
    try {
      await writeDeliverable(directory, "Agreement.docx", await docx(identity, excluded, "Completion is Monday."))
      const old = readCurrentVersion(directory, "Agreement.docx")!.current
      await writeDeliverableRevision(directory, "Agreement.docx", await docx(identity, included, "Completion is Tuesday."))
      const returned = mode === "missing-coverage" ? await docx(identity, "Completion is Wednesday.") : await docx(identity, excluded, "Completion is Wednesday.")
      const before = readDocumentVersions(directory, "Agreement.docx"), canonical = readFileSync(path.join(directory, "Agreement.docx"))
      fetch = spyOn(globalThis, "fetch").mockImplementation((async (url, init) => {
        expect(String(url).endsWith("/api/dochaus/generate")).toBe(true)
        const { prompt, modelRole } = JSON.parse(String(init?.body)); expect(modelRole).toBe("fast"); calls++
        const reading = block(prompt, "source-part"), candidate = block(prompt, "candidate"), left = block(prompt, "left-version"), clarification = block(prompt, "identity-clarification")
        let text: string
        if (reading) text = JSON.stringify({ quotes: [{ paragraph: 1, text: identity }], limitations: mode === "external-unknown" ? ["The schedule identifying the transaction is unavailable."] : [] })
        else if (candidate) {
          const relation = mode === "all-different" ? "different-document" : candidate.versionId === old.id || mode === "both-uncertain" ? "uncertain" : "same-document"
          expect(prompt).toContain("Do not decide historical descent")
          text = JSON.stringify({ versionId: candidate.versionId, relation, reason: relation === "uncertain" ? "The selected incoming reading omits the candidate's complete qualification; its identifying scope is unresolved." : "Synthetic identity protocol judgment.", evidence: relation === "uncertain" ? [] : [{ versionId: candidate.versionId, incomingQuote: identity, candidateQuote: identity }] })
        } else if (left) {
          contrasts++
          const right = block(prompt, "right-version"), incoming = block(prompt, "incoming")
          text = JSON.stringify({ support: incoming.text === left.passage.text ? "left" : incoming.text === right.passage.text ? "right" : "neither", reason: "Synthetic directional protocol judgment.", limitations: [], evidence: { incomingQuote: incoming.text, leftQuote: left.passage.text, rightQuote: right.passage.text } })
        } else if (clarification) {
          clarifications++
          const initial = block(prompt, "initial-identity"), readings = block(prompt, "initial-readings")
          expect(initial.relation).toBe("uncertain")
          expect(initial.reason).toContain("complete qualification")
          expect(readings.incoming.parts[0].quotes[0].text).toBe(identity)
          expect(readings.incoming.parts[0].quotes.some((quote: { text: string }) => quote.text === excluded)).toBe(false)
          expect(clarification.coverage.requiredParts).toBe(2)
          expect(clarification.triple.incoming.text).not.toBe(identity)
          expect(Buffer.byteLength(prompt)).toBeLessThanOrEqual(64 * 1024)
          if (mode === "external-unknown") expect(JSON.stringify(readings)).toContain("schedule identifying the transaction")
          const relation = clarification.triple.incoming.text.startsWith("Completion") ? "neutral" : mode === "still-uncertain" || mode === "external-unknown" ? "uncertain" : "same-document"
          expect(prompt).toContain("Neutral NEVER means missing correspondence")
          text = JSON.stringify({ versionId: clarification.versionId, relation, reason: relation === "uncertain" ? "The new passage does not resolve the original identifying qualification." : relation === "neutral" ? "This fully read completion term is independently non-identifying; it neither resolves nor contradicts the separate property qualification." : "The complete new paragraph resolves the omitted identifying scope alongside the original transaction-specific identity quote.", limitations: relation === "uncertain" ? ["Transaction identity remains unresolved."] : [],
            evidence: relation === "uncertain" ? [] : [{ versionId: clarification.versionId, incomingQuote: clarification.triple.incoming.text, candidateQuote: clarification.triple[clarification.candidateSide].text }] })
          if (mode === "source-drift" && !drifted) {
            drifted = true
            const db = openDb(directory); try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [old.id]) } finally { db.close() }
          }
        } else {
          reconciliations++
          const summary = block(prompt, "checked-base-contrasts")
          if (mode === "resolved" || mode === "both-uncertain") expect(summary.allowedBaseIds).toEqual([old.id])
          else expect(summary.allowedBaseIds).toEqual([])
          text = mode === "all-different" ? JSON.stringify({ choice: "new-document", reason: "All original identity pairs identify different documents; protocol fixture only.", evidence: before.versions.map(row => ({ versionId: row.id, incomingQuote: identity, candidateQuote: identity })) })
            : mode === "missing-coverage" ? JSON.stringify({ choice: "uncertain", reason: "Required correspondence is unavailable.", evidence: [] })
              : JSON.stringify({ choice: "candidate", versionId: old.id, reason: "Synthetic protocol choice; rejected unless both identity and version support qualify.", evidence: [{ versionId: old.id, incomingQuote: excluded, candidateQuote: excluded }] })
        }
        return Response.json({ data: { text, model, usage: { tokens: { input: 12, output: 8, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
      }) as typeof globalThis.fetch)
      const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { throw new Error("Advice cannot save.") } }))
      const body = new FormData(); body.append("file", new File([new Uint8Array(returned)], "Return.docx"))
      const response = await app.request("/matters/one/documents/returns/tie-break", { method: "POST", body }), result = await response.json()
      if (mode === "source-drift") { expect(response.status).toBe(409); expect(result.code).toBe("match_source_changed"); expect(reconciliations).toBe(0) }
      else if (mode === "still-uncertain" || mode === "external-unknown") {
        expect(response.status).toBe(409); expect(result.error).toContain("reconciliation did not follow its checked candidate evidence"); expect(contrasts).toBe(2); expect(clarifications).toBe(2)
      } else {
        expect(response.status).toBe(200)
        if (mode === "all-different") { expect(result.advisory.choice).toBe("new-document"); expect(contrasts).toBe(0); expect(clarifications).toBe(0); expect(calls).toBe(6) }
        else if (mode === "missing-coverage") {
          expect(result.advisory.choice).toBe("uncertain")
          expect(result.advisory.identityClarifications[0].relation).toBe("uncertain")
          expect(result.advisory.identityClarifications[0].unresolved.length).toBeGreaterThan(0)
        } else {
          expect(result.advisory.versionId).toBe(old.id); expect(contrasts).toBe(2); expect(clarifications).toBe(mode === "both-uncertain" ? 4 : 2)
          expect(result.advisory.judgments.find((row: { versionId: string }) => row.versionId === old.id).relation).toBe("uncertain")
          expect(result.advisory.identityClarifications.find((row: { versionId: string }) => row.versionId === old.id).relation).toBe("same-document")
          expect(result.advisory.identityClarifications.every((row: { coverage: { requiredParts: number; processedParts: number } }) => row.coverage.requiredParts === row.coverage.processedParts)).toBe(true)
        }
      }
      if (mode !== "source-drift") expect(readDocumentVersions(directory, "Agreement.docx")).toEqual(before)
      expect(readFileSync(path.join(directory, "Agreement.docx"))).toEqual(canonical)
    } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
  }
}, 30_000)
