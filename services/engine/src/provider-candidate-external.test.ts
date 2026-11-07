import { afterEach, expect, test } from "bun:test"
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Credential } from "@opencode/core/credential"
import { Integration } from "@opencode/core/integration"
import { Effect, Logger } from "effect"
import { providerCandidates, candidateInput } from "./provider-candidate"
import { createProviderSetupSession, type SetupCredential } from "./provider-setup-session"
import { credentialReuse } from "./credential-reuse"

// Claude Agent checks the local `claude` program without a prompt. These tests
// never run a real `claude`: PATH is emptied or OPENCODE_CLAUDE_PATH points at a
// synthetic stand-in that answers only the start-up request.
const integrationID = Integration.ID.make("claude-agent")
const projectVariables = ["GOOGLE_VERTEX_PROJECT", "GOOGLE_CLOUD_PROJECT", "GCP_PROJECT", "GCLOUD_PROJECT"]
const saved = Object.fromEntries(
  ["PATH", "OPENCODE_CLAUDE_PATH", ...projectVariables].map((name) => [name, process.env[name]]),
)
const directories: string[] = []
afterEach(() => {
  for (const [name, value] of Object.entries(saved)) {
    if (value === undefined) delete process.env[name]
    else process.env[name] = value
  }
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
})

function temporary() {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-external-test-"))
  directories.push(directory)
  return directory
}

// A synthetic `claude` that records its flags and the messages it receives.
function fakeClaude(account: Record<string, unknown>) {
  const directory = temporary()
  const log = path.join(directory, "log.jsonl")
  const executable = path.join(directory, "claude")
  const reply = {
    commands: [],
    agents: [],
    models: [
      {
        value: "default",
        resolvedModel: "synthetic-main",
        displayName: "Default",
        description: "Synthetic default",
        supportedEffortLevels: ["low", "high"],
      },
      {
        value: "synthetic-main",
        resolvedModel: "synthetic-main",
        displayName: "Synthetic Main",
        description: "Synthetic model",
        supportedEffortLevels: ["low", "high"],
      },
    ],
    account,
    claude_code_version: "2.1.250",
  }
  writeFileSync(
    executable,
    `#!${process.execPath}
const fs = require("node:fs")
const log = ${JSON.stringify(log)}
const reply = ${JSON.stringify(reply)}
fs.appendFileSync(log, JSON.stringify({ args: process.argv.slice(2), memory: process.env.CLAUDE_CODE_DISABLE_AUTO_MEMORY, instructions: process.env.CLAUDE_CODE_DISABLE_CLAUDE_MDS }) + "\\n")
let buffer = ""
process.stdin.setEncoding("utf8")
process.stdin.on("data", (chunk) => {
  buffer += chunk
  for (let index = buffer.indexOf("\\n"); index >= 0; index = buffer.indexOf("\\n")) {
    const line = buffer.slice(0, index)
    buffer = buffer.slice(index + 1)
    if (!line.trim()) continue
    const message = JSON.parse(line)
    fs.appendFileSync(log, JSON.stringify({ received: message.type, subtype: message.request?.subtype }) + "\\n")
    if (message.type === "control_request" && message.request?.subtype === "initialize")
      process.stdout.write(JSON.stringify({ type: "control_response", response: { subtype: "success", request_id: message.request_id, response: reply } }) + "\\n")
  }
})
process.stdin.on("end", () => process.exit(0))
`,
  )
  chmodSync(executable, 0o755)
  return {
    executable,
    log: () =>
      readFileSync(log, "utf8")
        .trim()
        .split("\n")
        .map((line) => JSON.parse(line)),
  }
}

async function withManager<T>(
  credential: SetupCredential | undefined,
  use: (manager: ReturnType<typeof providerCandidates>, ids: Credential.ID[]) => Promise<T>,
) {
  const legal = temporary()
  writeFileSync(path.join(legal, "opencode.json"), "{}")
  return Effect.runPromise(
    Effect.scoped(
      Effect.gen(function* () {
        const baseline = yield* createProviderSetupSession({ legal, settings: {}, credential })
        const ids = (yield* baseline.credentials.all()).map((item) => item.id)
        const manager = providerCandidates({
          legal,
          credentials: baseline.credentials,
          client: undefined as never,
          settings: () => ({}),
          inspect: (effect) => Effect.runPromise(baseline.provide(effect)),
          reuse: credentialReuse({
            directory: path.join(legal, "absent"),
            targetDirectory: legal,
            create: async () => {
              throw new Error("Unexpected import")
            },
          }),
        })
        return yield* Effect.promise(() => use(manager, ids))
      }),
    ).pipe(Effect.provide(Logger.layer([]))),
  )
}

test("external connections accept only a method identifier", () => {
  const input = { providerID: "claude-agent", connection: { type: "external", methodID: "claude-login" } }
  expect(candidateInput(input)).toBe(true)
  expect(candidateInput({ ...input, connection: { type: "external" } })).toBe(false)
  expect(candidateInput({ ...input, connection: { type: "external", methodID: "../claude" } })).toBe(false)
  expect(candidateInput({ ...input, connection: { type: "external", methodID: "claude-login", key: "x" } })).toBe(false)
  expect(
    candidateInput({ ...input, connection: { type: "external", methodID: "adc", answer: { project: "p" } } }),
  ).toBe(true)
  expect(candidateInput({ ...input, connection: { type: "external", methodID: "adc", answer: "p" } })).toBe(false)
})

test("a missing Claude Code passes the native needs_auth message through unchanged", async () => {
  const empty = temporary()
  mkdirSync(path.join(empty, "bin"))
  process.env.PATH = path.join(empty, "bin")
  delete process.env.OPENCODE_CLAUDE_PATH
  await withManager(undefined, async (manager) => {
    await expect(
      manager.handle("session", { providerID: "claude-agent", connection: { type: "external", methodID: "unknown" } }),
    ).rejects.toThrow("Select an available sign-in for this provider.")
    const created = (await manager.handle("session", {
      providerID: "claude-agent",
      connection: { type: "external", methodID: "claude-login" },
    })) as { sessionID: string; external: unknown }
    expect(created.external).toEqual({
      status: "needs_auth",
      message: "Claude Code is not installed. Install it, then run `claude auth login`.",
      url: "https://code.claude.com/docs/en/setup",
    })
    await manager.close(created.sessionID)
  })
})

test("a signed-in Claude Code is ready and lists models without a prompt or user settings", async () => {
  const claude = fakeClaude({ email: "user@example.com", subscriptionType: "Synthetic", apiProvider: "firstParty" })
  process.env.OPENCODE_CLAUDE_PATH = claude.executable
  await withManager(undefined, async (manager) => {
    const created = (await manager.handle("session", {
      providerID: "claude-agent",
      connection: { type: "external", methodID: "claude-login" },
    })) as { sessionID: string; external: unknown }
    expect(created.external).toEqual({ status: "ready" })
    const models = (await manager.handle("models", { sessionID: created.sessionID })) as { data: { id: string }[] }
    expect(models.data.map((model) => model.id)).toEqual(["synthetic-main"])
    await manager.close(created.sessionID)
  })
  const log = claude.log()
  const spawns = log.filter((line) => line.args)
  expect(spawns.length).toBeGreaterThan(0)
  for (const spawn of spawns) {
    expect(spawn.args).toContain("--setting-sources=")
    expect(spawn.args).toContain("--strict-mcp-config")
    expect(spawn.instructions).toBe("1")
    expect(spawn.memory).toBe("1")
  }
  expect(log.some((line) => line.received === "user")).toBe(false)
})

test("a saved external connection is checked again and reports a signed-out Claude Code", async () => {
  const claude = fakeClaude({ tokenSource: "none", apiProvider: "firstParty" })
  process.env.OPENCODE_CLAUDE_PATH = claude.executable
  const credential = {
    integrationID,
    label: "Claude Agent",
    value: Credential.External.make({ type: "external", methodID: Integration.MethodID.make("claude-login") }),
  }
  await withManager(credential, async (manager, [id]) => {
    const created = (await manager.handle("session", {
      providerID: "claude-agent",
      connection: { type: "credential", id: id! },
    })) as { sessionID: string; external: unknown }
    expect(created.external).toEqual({
      status: "needs_auth",
      message: "Claude Code is not logged in. Run `claude auth login` in a terminal.",
    })
    await manager.close(created.sessionID)
  })
  expect(claude.log().some((line) => line.received === "user")).toBe(false)
})

// Vertex's Google Cloud credentials are an external method with fields. The
// answer configures the connection; no Google request is made.
test("an external method with fields connects with its answers and rejects incomplete ones", async () => {
  for (const name of projectVariables) delete process.env[name]
  await withManager(undefined, async (manager) => {
    const connect = (answer?: Record<string, string>) =>
      manager.handle("session", {
        providerID: "google-vertex",
        connection: { type: "external", methodID: "google-adc", ...(answer ? { answer } : {}) },
      }) as Promise<{ sessionID: string; external: unknown }>
    await expect(connect()).rejects.toThrow("Complete the provider connection fields.")
    await expect(connect({ location: "global" })).rejects.toThrow("Complete the provider connection fields.")
    await expect(connect({ project: "synthetic-project", location: "global", extra: "x" })).rejects.toThrow(
      "Complete the provider connection fields.",
    )
    // The native route validates the required project again, so success shows
    // the answer reached it.
    const created = await connect({ project: "synthetic-project", location: "global" })
    expect(created.external).toEqual({ status: "ready" })
    const models = (await manager.handle("models", { sessionID: created.sessionID })) as { data: unknown[] }
    expect(models.data.length).toBeGreaterThan(0)
    await manager.close(created.sessionID)
  })
})
