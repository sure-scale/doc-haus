import os from "node:os"
import path from "node:path"
import { Hono } from "hono"
import { readEngineSettings } from "./engine-config"

// Host-credential discovery for the web's provider-setup UI: which GCP projects
// the server's gcloud ADC can see, and which AWS profiles its ~/.aws files
// define. Both live on the host the engine runs on — the browser cannot see
// them — and both resolve to empty lists when the host has no such sign-in, so
// the UI's fields degrade to plain typing.

// The projects the host's ADC token can list, via Cloud Resource Manager. The
// token comes from the gcloud CLI rather than a Google SDK so the ingest
// service needs no GCP dependency — the same sign-in the engine's Vertex calls
// rely on is the one being enumerated.
export async function listGcpProjects() {
  const proc = Bun.spawn(["gcloud", "auth", "application-default", "print-access-token"], {
    stdout: "pipe",
    stderr: "ignore",
  })
  const token = (await new Response(proc.stdout).text()).trim()
  if ((await proc.exited) !== 0 || !token) return []
  const res = await fetch("https://cloudresourcemanager.googleapis.com/v1/projects?filter=lifecycleState:ACTIVE", {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!res.ok) return []
  const body = (await res.json()) as { projects?: { projectId: string }[] }
  return (body.projects ?? []).map((p) => p.projectId).sort()
}

// Compatibility for older web clients. Inference belongs to the native engine:
// this route only checks the already saved cloud options and forwards model IDs.
export function vertexProbeRoutes() {
  return new Hono().post("/probe-vertex", async c => {
    const input: unknown = await c.req.json().catch(() => undefined)
    const invalid = { ok: false, kind: "model", error: "Choose a Vertex publisher, project, location and model to test." }
    if (!input || typeof input !== "object" || Array.isArray(input)) return c.json(invalid, 400)
    const value = input as Record<string, unknown>
    if (Object.keys(value).some(key => !["project", "location", "publisher", "model"].includes(key)) ||
        typeof value.publisher !== "string" || !["google", "anthropic"].includes(value.publisher) ||
        ["project", "location", "model"].some(key => typeof value[key] !== "string" || !value[key].trim() || value[key].length > 256 || /[\s#\x00-\x1f\x7f]/.test(value[key].trim())))
      return c.json(invalid, 400)
    const providerID = value.publisher === "google" ? "google-vertex" : "google-vertex-anthropic"
    const unavailable = { ok: false, kind: "unavailable", error: "The native engine provider check could not complete. Check the engine and retry." }
    try {
      const settings = readEngineSettings()
      const options = settings.provider?.[providerID]?.options
      if (!options || typeof options !== "object" || Array.isArray(options) ||
          typeof (options as Record<string, unknown>).project !== "string" ||
          String((options as Record<string, unknown>).project).trim() !== String(value.project).trim() ||
          String((options as Record<string, unknown>).location ?? "global").trim() !== String(value.location).trim())
        return c.json({ ok: false, kind: "model", error: "Save this provider's project and location in Settings before testing it." }, 409)
      const engine = process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
      const password = process.env.OPENCODE_SERVER_PASSWORD
      const response = await fetch(`${engine}/api/dochaus/provider-test`, {
        method: "POST", signal: AbortSignal.any([c.req.raw.signal, AbortSignal.timeout(25_000)]),
        headers: { "Content-Type": "application/json", ...(password ? { Authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}) },
        body: JSON.stringify({ providerID, modelID: String(value.model).trim() }),
      })
      // Bound the managed result and copy only its public contract. Provider
      // diagnostics, headers and credentials never become legacy-route fields.
      const reader = response.body?.getReader()
      if (!reader) return c.json(unavailable, 503)
      const chunks: Uint8Array[] = []
      let length = 0
      while (true) {
        const chunk = await reader.read()
        if (chunk.done) break
        length += chunk.value.length
        if (length > 8192) { await reader.cancel(); return c.json(unavailable, 503) }
        chunks.push(chunk.value)
      }
      const bytes = new Uint8Array(length)
      let offset = 0
      for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length }
      const result: unknown = JSON.parse(new TextDecoder().decode(bytes))
      if (!result || typeof result !== "object" || Array.isArray(result)) return c.json(unavailable, 503)
      const body = result as Record<string, unknown>
      if (response.ok && body.ok === true) return Response.json({ ok: true }, { status: response.status })
      const kind = ["auth", "model", "unavailable"].includes(String(body.kind)) ? body.kind : "unavailable"
      const error = typeof body.error === "string" && body.error.length <= 1024 ? body.error : unavailable.error
      return Response.json({ ok: false, kind, error,
        ...(body.expired === true && kind === "auth" ? { expired: true } : {}),
        ...(kind === "model" && ["free-tier", "region"].includes(String(body.restriction)) ? { restriction: body.restriction } : {}),
      }, { status: response.status })
    } catch { return c.json(unavailable, 503) }
  })
}

// Profile names from ~/.aws/config ([profile x] plus the bare [default]) and
// ~/.aws/credentials ([x]), deduped.
export async function listAwsProfiles() {
  const read = (file: string) =>
    Bun.file(path.join(os.homedir(), ".aws", file))
      .text()
      .catch(() => "")
  const [config, credentials] = await Promise.all([read("config"), read("credentials")])
  const names = [
    ...[...config.matchAll(/^\[profile ([^\]]+)\]/gm)].map((m) => m[1]!.trim()),
    ...(/^\[default\]/m.test(config) ? ["default"] : []),
    ...[...credentials.matchAll(/^\[([^\]]+)\]/gm)].map((m) => m[1]!.trim()),
  ]
  return [...new Set(names)].sort()
}
