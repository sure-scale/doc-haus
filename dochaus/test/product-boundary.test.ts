import { describe, expect, test } from "bun:test"
import { readFileSync, readdirSync } from "node:fs"
import path from "node:path"

// Product-boundary guard (issue #130): the generic runtime under plugin/ and
// lib/ must stay jurisdiction-neutral. Named statutes, regulators, statute
// families, and jurisdiction abbreviations belong in jurisdiction packs
// (dochaus/jurisdiction/<code>/), which are data and deliberately NOT scanned
// here. If this test fails, move the offending reference into a pack or
// replace it with a generic shape.
const BANNED = [
  /\bU\.?S\.?C\b/,
  /\bC\.?F\.?R\b/,
  /\bSEC\b/,
  /sec\.gov/i,
  /\bAdvisers Act\b/i,
  /\bExchange Act\b/i,
  /\bSecurities Act\b/i,
  /\bGen\.?\s?Stat\b/i,
  /\bState\s+Agencies\b/i,
  /\bConn\b/,
  /\bDGCL\b/,
  /\bGDPR\b/i,
  /\bFINRA\b/i,
]

const ROOT = path.join(import.meta.dir, "..")

function tsFiles(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) return tsFiles(full)
    return entry.name.endsWith(".ts") && !entry.name.endsWith(".test.ts") ? [full] : []
  })
}

describe("product boundary", () => {
  test("tool discovery contains no test modules", () => {
    expect(
      readdirSync(path.join(ROOT, "tool"), { withFileTypes: true })
        .filter((entry) => entry.isFile() && /\.(?:test|spec)\.[cm]?[jt]s$/.test(entry.name))
        .map((entry) => entry.name)
        .sort(),
    ).toEqual([])
  })

  test("generic runtime names no jurisdiction-specific authorities", () => {
    const files = [...tsFiles(path.join(ROOT, "plugin")), ...tsFiles(path.join(ROOT, "lib"))]
    expect(files.length).toBeGreaterThan(0)
    const violations = files.flatMap((file) => {
      const text = readFileSync(file, "utf8")
      return BANNED.flatMap((pattern) =>
        pattern.test(text) ? [`${path.relative(ROOT, file)} matches ${pattern}`] : [],
      )
    })
    expect(violations).toEqual([])
  })
})
