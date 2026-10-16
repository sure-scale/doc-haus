import { readFileSync } from "node:fs"
import path from "node:path"
import { describe, expect, test } from "bun:test"

const repoRoot = path.resolve(import.meta.dir, "..", "..")
const quietTurnMS = 20 * 60_000
const timeoutOptions = ["timeout", "headerTimeout", "chunkTimeout"] as const

function explicitTimeoutViolations(config: {
  provider?: Record<string, { options?: Partial<Record<(typeof timeoutOptions)[number], number | false>> }>
}) {
  // Check explicit settings only. Omitted timeouts leave defaults and transport
  // behavior to the provider/executor; this check makes no guarantee about them.
  return Object.entries(config.provider ?? {}).flatMap(([providerID, provider]) =>
    timeoutOptions.flatMap((option) => {
      const timeout = provider.options?.[option]
      if (timeout === undefined || timeout === false) return []
      return typeof timeout === "number" && timeout >= quietTurnMS ? [] : [`${providerID}.${option}`]
    }),
  )
}

describe("explicit generation timeouts do not undercut documented quiet turns", () => {
  test("checks every configured provider in the shipped config without requiring model pins", () => {
    const config = JSON.parse(readFileSync(path.join(repoRoot, "dochaus/opencode.json"), "utf8"))
    expect(explicitTimeoutViolations(config)).toEqual([])
  })

  test("allows a config without saved providers or model selections", () => {
    expect(explicitTimeoutViolations({})).toEqual([])
    expect(explicitTimeoutViolations({ provider: {} })).toEqual([])
  })

  for (const option of timeoutOptions) {
    test.each([undefined, false, quietTurnMS])(`allows ${option} = %p`, (timeout) => {
      expect(explicitTimeoutViolations({ provider: { synthetic: { options: { [option]: timeout } } } })).toEqual([])
    })

    test(`rejects a short ${option} on a provider without a model pin`, () => {
      const config = {
        model: "selected/model",
        provider: {
          selected: { options: { timeout: quietTurnMS } },
          unselected: { options: { [option]: quietTurnMS } },
        },
      }
      expect(explicitTimeoutViolations(config)).toEqual([])

      config.provider.unselected.options[option] = quietTurnMS - 1
      expect(explicitTimeoutViolations(config)).toEqual([`unselected.${option}`])
    })
  }
})

describe("draft mode lean owner contract", () => {
  const prompt = readFileSync(path.join(repoRoot, "dochaus/mode/draft.md"), "utf8")
  const frontmatter = prompt.split("---")[1]!

  test("enables no tools", () => {
    const tools = frontmatter.slice(frontmatter.indexOf("tools:"))

    expect(tools).toContain('  "*": false')
    expect([...tools.matchAll(/^  ["']?([a-z*-]+)["']?: true$/gm)]).toEqual([])
  })

  test("does not publish a task permission catalog", () => {
    expect(frontmatter).not.toContain("permission:")
    expect(frontmatter).not.toContain("subagent_type")
  })
})
