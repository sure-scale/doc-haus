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

function object(value:unknown):Record<string,unknown> {return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string,unknown> : {}}
