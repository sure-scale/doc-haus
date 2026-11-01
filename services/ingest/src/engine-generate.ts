import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { openTelemetryDb, recordGeneration, type GenerationUsage } from "../../../dochaus/lib/telemetry"

// One stateless call to the engine's generate route: the user's selected model,
// no tools, no fallback and no session. Callers frame the prompt and check the
// text. `signal` (a closed request, a Cancel) stops the call before the timeout,
// and the engine interrupts the model when the request goes.
export class EngineGenerationError extends Error {}

export async function generate(directory:string,prompt:string,signal?:AbortSignal) {
  const password = process.env.OPENCODE_SERVER_PASSWORD
  const base = process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
  const timeout = AbortSignal.timeout(480000)
  const response = await fetch(`${base}/api/dochaus/generate`,{method:"POST",headers:{"content-type":"application/json","x-opencode-directory":directory,
    ...(password ? {authorization:`Basic ${btoa(`opencode:${password}`)}`} : {})},body:JSON.stringify({prompt}),signal:signal ? AbortSignal.any([timeout,signal]) : timeout})
  if (!response.ok) throw new EngineGenerationError(`The configured legal generation failed (HTTP ${response.status}).`)
  const result = object(await response.json());const value=object(result.data)
  if (typeof value.text !== "string" || !value.text) throw new EngineGenerationError("Legal generation returned no text.")
  return {text:value.text,model:value.model,elapsed_ms:value.elapsed_ms,usage:value.usage}
}

// The same call, with the provider-reported usage kept in the matter's usage
// ledger under `agent`. The ledger is a report: a failed write never fails the job.
export async function generateFor(agent:string,directory:string,prompt:string,signal?:AbortSignal) {
  const result = await generate(directory,prompt,signal)
  recordGenerationUsage(directory,agent,result)
  return result
}

function recordGenerationUsage(directory:string,agent:string,result:{model:unknown;usage:unknown}) {
  const usage = generationUsage(result.usage)
  if (!usage) return
  const model = object(result.model)
  try {
    const db = openTelemetryDb(directory)
    try {
      recordGeneration(db,{agent,...typeof model.providerID === "string" && typeof model.id === "string" ? {model:`${model.providerID}/${model.id}`} : {},...usage})
    } finally { db.close() }
  } catch (error) { console.warn("[engine-generate] usage was not recorded",{agent,...diagnosticErrorMetadata(error)}) }
}

function generationUsage(value:unknown):Pick<GenerationUsage,"tokens"|"cost">|undefined {
  const usage = object(value), tokens = object(usage.tokens), cache = object(tokens.cache)
  const count = (item:unknown) => typeof item === "number" && Number.isFinite(item) && item >= 0 ? item : undefined
  const input = count(tokens.input), output = count(tokens.output), reasoning = count(tokens.reasoning), read = count(cache.read), write = count(cache.write)
  if (input === undefined || output === undefined || reasoning === undefined || read === undefined || write === undefined) return
  const cost = usage.cost === null ? null : count(usage.cost)
  if (cost === undefined) return
  return {tokens:{input,output,reasoning,cache:{read,write}},cost}
}

function object(value:unknown):Record<string,unknown> {return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string,unknown> : {}}
