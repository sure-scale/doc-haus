import { DOCUMENT_AMENDMENT_LEGAL_OUTPUT } from "../../../dochaus/lib/document-amendment-legal-contract"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { openTelemetryDb, recordGeneration, type GenerationUsage } from "../../../dochaus/lib/telemetry"

// One stateless call to the engine's generate route: the user's selected model,
// no tools, no fallback and no session. Callers frame the prompt and check the
// text. `signal` (a closed request, a Cancel) stops the call before the timeout,
// and the engine interrupts the model when the request goes.
export class EngineGenerationError extends Error {
  constructor(message: string, readonly generation?: { model: unknown; usage: unknown }) { super(message) }
}

export async function generate(directory:string,prompt:string,signal?:AbortSignal,outputContract?:typeof DOCUMENT_AMENDMENT_LEGAL_OUTPUT,modelRole?:"fast") {
  const password = process.env.OPENCODE_SERVER_PASSWORD
  const base = process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
  const timeout = AbortSignal.timeout(480000)
  const response = await fetch(`${base}/api/dochaus/generate`,{method:"POST",headers:{"content-type":"application/json","x-opencode-directory":directory,
    ...(password ? {authorization:`Basic ${btoa(`opencode:${password}`)}`} : {})},body:JSON.stringify({prompt,...outputContract ? {outputContract} : {},...modelRole ? {modelRole} : {}}),signal:signal ? AbortSignal.any([timeout,signal]) : timeout})
  const result = object(await response.json().catch(() => undefined));const value=object(result.data)
  if (!response.ok) throw new EngineGenerationError(`The configured legal generation failed (HTTP ${response.status}).`,
    outputContract && ((response.status === 422 && object(value.error).code === "invalid_structured_output") ||
      (response.status === 503 && object(value.error).code === "structured_generation_failed"))
      ? { model: value.model, usage: value.usage } : undefined)
  if (typeof value.text !== "string" || !value.text) throw new EngineGenerationError("Legal generation returned no text.", outputContract ? { model: value.model, usage: value.usage } : undefined)
  return {text:value.text,model:value.model,elapsed_ms:value.elapsed_ms,usage:value.usage}
}

// The same call, with the provider-reported usage kept in the matter's usage
// ledger under `agent`. The ledger is a report: a failed write never fails the job.
export async function generateFor(agent:string,directory:string,prompt:string,signal?:AbortSignal,outputContract?:typeof DOCUMENT_AMENDMENT_LEGAL_OUTPUT,modelRole?:"fast") {
  try {
    const result = await generate(directory,prompt,signal,outputContract,modelRole)
    recordGenerationUsage(directory,agent,result,Boolean(outputContract) || Boolean(modelRole))
    return result
  } catch (error) {
    if (outputContract || modelRole) recordGenerationUsage(directory,agent,
      error instanceof EngineGenerationError && error.generation ? error.generation : { model: undefined, usage: undefined },true)
    throw error
  }
}

const unknownUsage: Pick<GenerationUsage,"tokens"|"cost"> = { tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, cost: null }
function recordGenerationUsage(directory:string,agent:string,result:{model:unknown;usage:unknown},keepUnknown = false) {
  // A lost structured response cannot prove that its model request was free.
  // Retain the existing explicit unknown-cost representation for reconciliation.
  const usage = generationUsage(result.usage) ?? (keepUnknown ? unknownUsage : undefined)
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
