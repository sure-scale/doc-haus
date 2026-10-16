export {}

const loopback = (hostname: string) => hostname === "0.0.0.0" ? "127.0.0.1"
  : hostname.includes(":") ? `[${hostname === "::" ? "::1" : hostname}]` : hostname
const targets = [
  `http://${loopback(process.env.INGEST_HOST ?? "127.0.0.1")}:${process.env.DOC_HAUS_INGEST_PORT ?? "4501"}/engine/config`,
  `http://${loopback(process.env.WEB_HOST ?? process.env.OPENCODE_HOSTNAME ?? "127.0.0.1")}:${process.env.DOC_HAUS_WEB_PORT ?? "5174"}/`,
]
const deadline = Date.now() + 60_000
for (const url of targets) {
  while (true) {
    const response = await fetch(url, { signal: AbortSignal.timeout(5000) }).catch(() => undefined)
    await response?.body?.cancel()
    if (response?.ok) break
    if (Date.now() >= deadline) throw new Error("The managed ingest and web services did not become ready within 60 seconds")
    await Bun.sleep(250)
  }
}
console.log("doc.haus: ingest and web are ready for native restart")
