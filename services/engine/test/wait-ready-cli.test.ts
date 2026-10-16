import { expect, test } from "bun:test"
import path from "node:path"

const repository = path.resolve(import.meta.dir, "../../..")
const waiter = path.join(repository, "services/engine/script/wait-ready.ts")
const prefix = "DOCHAUS_ENGINE_READINESS "

async function run(body: string, status: number, observe: boolean) {
  const server = Bun.serve({
    hostname: "127.0.0.1",
    port: 0,
    fetch: request => {
      expect(new URL(request.url).pathname).toBe("/api/dochaus/ready")
      return new Response(body, { status, headers: { "content-type": "application/json" } })
    },
  })
  const child = Bun.spawn([process.execPath, "--no-env-file", "run", waiter], {
    cwd: repository,
    env: {
      WORKSPACE_ROOT: "/tmp",
      OPENCODE_HOSTNAME: "127.0.0.1",
      DOC_HAUS_ENGINE_PORT: String(server.port),
      ...(observe ? { DOCHAUS_STARTUP_OBSERVATION: "1" } : {}),
    },
    stdout: "pipe",
    stderr: "pipe",
  })
  try {
    const stdout = new Response(child.stdout).text()
    const stderr = new Response(child.stderr).text()
    return { exitCode: await child.exited, stdout: await stdout, stderr: await stderr }
  } finally {
    if (child.exitCode === null) child.kill()
    server.stop(true)
  }
}

function receipt(stdout: string) {
  const line = stdout.split("\n").find((value) => value.startsWith(prefix))
  expect(line).toBeDefined()
  return JSON.parse(line!.slice(prefix.length))
}

test("wait-ready emits bounded opt-in terminal receipts for native readiness", async () => {
  const active = await run(JSON.stringify({ ready: true, legalState: "active", legalPluginCount: 1, consentState: "active" }), 200, true)
  expect(active.exitCode).toBe(0)
  expect(receipt(active.stdout)).toEqual(expect.objectContaining({
    outcome: "active",
    attempts: 1,
    lastHttpStatus: 200,
    transportCategory: null,
    envelope: "valid_data",
    legalState: "active",
    legalPluginCount: 1,
    consentState: "active",
  }))

  const failed = await run(JSON.stringify({ ready: false, legalState: "failed", legalPluginCount: 1, consentState: "pending" }), 503, true)
  expect(failed.exitCode).not.toBe(0)
  expect(receipt(failed.stdout)).toEqual(expect.objectContaining({
    outcome: "plugin_failed",
    envelope: "valid_data",
    legalState: "failed",
    legalPluginCount: 1,
  }))

  const missingConsent = await run(JSON.stringify({ ready: false, legalState: "active", legalPluginCount: 1, consentState: "missing" }), 503, true)
  expect(missingConsent.exitCode).not.toBe(0)
  expect(receipt(missingConsent.stdout)).toMatchObject({ outcome: "consent_failed", consentState: "missing" })

  const malformed = await run("not-json", 200, true)
  expect(malformed.exitCode).not.toBe(0)
  expect(receipt(malformed.stdout)).toEqual(expect.objectContaining({
    outcome: "json_failure",
    lastHttpStatus: 200,
    transportCategory: null,
    envelope: "invalid_json",
    legalState: null,
    legalPluginCount: null,
  }))

  const defaultOff = await run(JSON.stringify({ ready: true, legalState: "active", legalPluginCount: 1, consentState: "active" }), 200, false)
  expect(defaultOff.exitCode).toBe(0)
  expect(defaultOff.stdout.includes(prefix)).toBe(false)
})
