import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { dataPaths } from "../../engine/src/data-paths"

const repository = fileURLToPath(new URL("../../../", import.meta.url))

export function wordConfigurationFile(environment: Record<string, string | undefined> = process.env) {
  return path.join(dataPaths(repository, environment).runtime, "config", "word-addin.json")
}

export function wordOrigin(value: string) {
  const url = new URL(value)
  if (url.protocol !== "https:" || url.username || url.password || url.pathname !== "/" || url.search || url.hash)
    throw new Error("The Word add-in needs an HTTPS origin without credentials, a path, query or fragment.")
  return url.origin
}

export type WordHttpsConfiguration = {
  origin: string
  certificate?: string
  key?: string
  external: boolean
  host: string
  port: number
}

export function wordHttpsConfiguration(environment: Record<string, string | undefined> = process.env): WordHttpsConfiguration | undefined {
  const file = wordConfigurationFile(environment)
  let saved: Record<string, unknown> = {}
  if (existsSync(file)) {
    const value: unknown = JSON.parse(readFileSync(file, "utf8"))
    if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("The Word add-in settings are invalid. Run setup again.")
    saved = value as Record<string, unknown>
  }
  const raw = environment.DOCHAUS_WORD_ORIGIN ?? saved.origin
  if (!raw) return
  if (typeof raw !== "string") throw new Error("The Word add-in HTTPS address is invalid.")
  const origin = wordOrigin(raw)
  const certificate = environment.DOCHAUS_WORD_TLS_CERT ?? saved.certificate
  const key = environment.DOCHAUS_WORD_TLS_KEY ?? saved.key
  const external = environment.DOCHAUS_WORD_EXTERNAL_HTTPS === "1" || (!environment.DOCHAUS_WORD_ORIGIN && saved.external === true)
  if (!external && (!certificate || !key)) return
  if ((certificate !== undefined && typeof certificate !== "string") || (key !== undefined && typeof key !== "string"))
    throw new Error("The Word add-in certificate paths are invalid.")
  if (!external && (!existsSync(certificate as string) || !existsSync(key as string))) return
  return { origin, certificate: certificate as string | undefined, key: key as string | undefined, external,
    host: environment.DOCHAUS_WORD_HOST ?? (new URL(origin).hostname === "[::1]" ? "::1" : "127.0.0.1"), port: Number(new URL(origin).port || 443) }
}
