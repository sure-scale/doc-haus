import { describe, expect, test } from "bun:test"
import { collectResearchSources, evaluateSourceIdentity, parseJsonFence, summarizeSourceChecks } from "../lib/research-source-audit"

describe("research source audit helpers", () => {
  test("parses research source rows from fenced JSON", () => {
    const rows = parseJsonFence(
      '```research\n[{"question":"q","tier":"verified","sources":[{"title":"Official Guidance","url":"https://example.test/guidance","kind":"official"}]}]\n```',
      "research",
    )
    expect(collectResearchSources(rows)).toEqual([
      {
        row: 1,
        question: "q",
        tier: "verified",
        title: "Official Guidance",
        url: "https://example.test/guidance",
        kind: "official",
      },
    ])
  })

  test("keeps the same source when it supports different research rows", () => {
    const rows = parseJsonFence(
      [
        "```research",
        "[",
        '{"question":"q1","tier":"verified","sources":[{"title":"15 U.S.C. § 80b-5","url":"https://example.test/80b-5","kind":"official"}]},',
        '{"question":"q2","tier":"verified","sources":[{"title":"15 U.S.C. § 80b-5","url":"https://example.test/80b-5","kind":"official"}]}',
        "]",
        "```",
      ].join("\n"),
      "research",
    )
    expect(collectResearchSources(rows).map((source) => source.row)).toEqual([1, 2])
  })

  test("blocks a fetched source whose named identity does not match the claimed title", () => {
    const result = evaluateSourceIdentity(
      {
        row: 1,
        question: "q",
        tier: "verified",
        title: "Agency Order - Alpha Advisory LLC, Rel. No. 1234",
        url: "https://example.test/order.pdf",
        kind: "official",
      },
      "Release No. 1234. In the Matter of Beta Advisory LLC. Findings and order.",
    )
    expect(result.ok).toBe(false)
    expect(result.reason).toContain("identity")
  })

  test("accepts a fetched source whose named identity and identifier match", () => {
    const result = evaluateSourceIdentity(
      {
        row: 1,
        question: "q",
        tier: "verified",
        title: "Agency Order - Alpha Advisory LLC, Rel. No. 1234",
        url: "https://example.test/order.pdf",
        kind: "official",
      },
      "Release No. 1234. In the Matter of Alpha Advisory LLC. Findings and order.",
    )
    expect(result.ok).toBe(true)
  })

  test("accepts a citation-style title when the fetched text contains the legal identifier", () => {
    const result = evaluateSourceIdentity(
      {
        row: 1,
        question: "q",
        tier: "verified",
        title: "State Agencies Regs. §12a-34-56b",
        url: "https://example.test/regulation",
        kind: "official",
      },
      "Current regulation text. Section 12a-34-56b states the operative rule.",
    )
    expect(result.ok).toBe(true)
  })

  test("accepts a citation-style title when the URL carries the base section", () => {
    const result = evaluateSourceIdentity(
      {
        row: 1,
        question: "q",
        tier: "verified",
        title: "15 U.S.C. § 99a-12(b)(3)",
        url: "https://official.example/statute/99a-12",
        kind: "official",
      },
      "Current section text without a repeated heading.",
    )
    expect(result.ok).toBe(true)
  })

  test("accepts a part-level CFR title when the URL carries the part", () => {
    const result = evaluateSourceIdentity(
      {
        row: 1,
        question: "q",
        tier: "verified",
        title: "17 CFR Part 248 — Regulation S-P: Privacy of Consumer Financial Information",
        url: "https://www.ecfr.gov/current/title-17/chapter-II/part-248",
        kind: "official",
      },
      "Current eCFR page text without the full title repeated.",
    )
    expect(result.ok).toBe(true)
  })

  test("treats redundant fetch failures as warnings when a verified row has another matching source", () => {
    const audit = summarizeSourceChecks([
      {
        source: {
          row: 1,
          question: "q",
          tier: "verified",
          title: "Official Source A",
          url: "https://official.example/a",
          kind: "official",
        },
        ok: true,
        reason: "matched",
        evidence: "matched title tokens",
      },
      {
        source: {
          row: 1,
          question: "q",
          tier: "verified",
          title: "Official Source B",
          url: "https://official.example/b",
          kind: "official",
        },
        ok: false,
        reason: "could not fetch or read source URL",
        evidence: "HTTP 403 Forbidden",
      },
    ])
    expect(audit.status).toBe("PARTIAL")
    expect(audit.failures).toHaveLength(0)
    expect(audit.warnings).toHaveLength(1)
  })
})
