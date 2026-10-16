import { expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { prepareLegalWebFetch, projectLegalWebFetch } from "../plugin/legal/research-webfetch"
import { LegalPlugin } from "../plugin/legal/policy"
import { frameUntrustedResearchText } from "../lib/untrusted"

test("remote HTML retains visible qualifications while excluding explicit hidden instructions", () => {
  const args = { url: "https://example.gov/document", format: "markdown" }
  prepareLegalWebFetch(args)
  expect(args.format).toBe("html")
  const output = {
    title: `${args.url} (text/html; charset=utf-8)`,
    output: '<p>Not every request requires <strong>approval</strong>.</p><p hidden>PRIVATE-HIDDEN-INSTRUCTION</p><p>Timing &amp; notice remain separate.</p>',
    metadata: {} as Record<string, unknown>,
  }
  projectLegalWebFetch(args, output)
  expect(output.output).toContain("Not every request requires approval.")
  expect(output.output).toContain("Timing & notice remain separate.")
  expect(output.output).not.toContain("PRIVATE-HIDDEN-INSTRUCTION")
  expect(output.metadata.sourceSnapshot).toBeUndefined()
  expect(frameUntrustedResearchText(output.output).output).toContain(output.output)
})

test("plain text is not parsed as HTML and truncated HTML never claims a complete source", () => {
  const args = { url: "https://example.gov/document", format: "html" }
  const plain = { title: `${args.url} (text/plain)`, output: "Use <term> literally; x < y & z > w.", metadata: {} }
  projectLegalWebFetch(args, plain)
  expect(plain.output).toBe("Use <term> literally; x < y & z > w.")
  const truncated = { title: `${args.url} (application/xhtml+xml)`, output: "<p>First visible passage.</p>", metadata: { truncated: true } as Record<string, unknown> }
  projectLegalWebFetch(args, truncated)
  expect(truncated.output).toContain("Source excerpt only")
  expect(truncated.metadata.sourceProjection).toMatchObject({ completeness: "truncated" })
  const malformed = { title: `${args.url} (text/html)`, output: "<div>".repeat(600), metadata: {} }
  expect(() => projectLegalWebFetch(args, malformed)).toThrow("nesting limit")
})

test("the production hooks preserve raw HTML until projection and frame only its visible result", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "legal-webfetch-hook-"))
  try {
    const hooks = await LegalPlugin({ directory } as Parameters<typeof LegalPlugin>[0])
    const input = { tool: "webfetch", sessionID: "ses_fetch", callID: "call_fetch" }
    const args = { url: "https://example.gov/document", format: "markdown" }
    await hooks["tool.execute.before"]!(input, { args })
    expect(args.format).toBe("html")
    const output = {
      title: `${args.url} (text/html; charset=utf-8)`,
      output: '<p>The exception does not remove the qualification.</p><span hidden>PRIVATE-HIDDEN-HOOK</span>',
      metadata: {} as Record<string, unknown>,
    }
    await hooks["tool.execute.after"]!({ ...input, args }, output)
    expect(output.output).toContain("The exception does not remove the qualification.")
    expect(output.output).not.toContain("PRIVATE-HIDDEN-HOOK")
    expect(output.output).toContain('<web-content untrusted="true">')
    expect(output.metadata.sourceProjection).toMatchObject({ completeness: "tool_returned_body" })
    expect(output.metadata.sourceSnapshot).toBeUndefined()
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})
