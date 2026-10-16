import type { ProxyOptions } from "vite"

// Vite's streaming HTTP proxy forwards bodies, query strings and directory
// headers unchanged. Strip the routing prefix only; native /api remains intact.
export function backendProxy(environment: Record<string, string | undefined> = process.env): Record<string, ProxyOptions> {
  return Object.fromEntries([
    ["engine", environment.DOC_HAUS_ENGINE_PORT ?? "4097"],
    ["ingest", environment.DOC_HAUS_INGEST_PORT ?? "4501"],
  ].map(([service, port]) => [`^/${service}(?:/|$)`, {
    target: `http://127.0.0.1:${port}`,
    changeOrigin: false,
    ws: true,
    rewrite: (url: string) => url.replace(new RegExp(`^/${service}(?=/|$)`), "") || "/",
  }]))
}
