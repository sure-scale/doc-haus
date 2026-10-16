import { createHash } from "node:crypto"

export const startupObservationEnabled = process.env.DOCHAUS_STARTUP_OBSERVATION === "1"
const phases = new Set(["graph", "migration", "listener", "config", "api.plugin", "api.config", "legal.setup", "legal.tools", "legal.reader", "legal.policy", "tools.import", "tools.transform", "policy.runtime", "policy.initialize", "policy.reset", "heartbeat", "network.denied", "preload"])
const names = new Set(["PluginModule.load", "Plugin.load", "PluginSupervisor.activate", "Plugin.list"])
let emitted = 0
let dropped = 0

/** Diagnostic metadata only; never include attributes, values or native causes. */
export function startupObservation(phase: string, state: "enter" | "exit" | "failed" | "sample", directory?: string, numbers: Record<string, number> = {}) {
  if (!startupObservationEnabled || !phases.has(phase)) return
  if (emitted >= 512 && phase !== "heartbeat") { dropped++; return }
  emitted++
  console.error("DOCHAUS_STARTUP_OBSERVATION " + JSON.stringify({
    phase, state, ms: performance.now(),
    ...(directory ? { location: createHash("sha256").update(directory).digest("hex") } : {}),
    ...numbers, dropped,
  }))
}

export function observeStartup<A>(phase: string, directory: string, operation: () => Promise<A>): Promise<A> {
  startupObservation(phase, "enter", directory)
  return operation().then(value => {
    startupObservation(phase, "exit", directory)
    return value
  }, error => {
    startupObservation(phase, "failed", directory)
    throw error
  })
}

export function startupObservationSpan(name: string, state: "enter" | "exit", span: string, status?: "Success" | "Failure") {
  if (!startupObservationEnabled || !names.has(name)) return
  if (emitted >= 512) { dropped++; return }
  emitted++
  console.error("DOCHAUS_STARTUP_OBSERVATION " + JSON.stringify({
    phase: name, state, ms: performance.now(), span,
    ...(status ? { status } : {}), dropped,
  }))
}
