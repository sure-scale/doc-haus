import { useEffect, useRef, useState } from "react"
import { getConfig } from "../api/opencode/config"
import { listMatters } from "../api/ingest/matters"
import { useToast } from "./Toast"

const HEALTHY_INTERVAL_MS = 10_000
const DOWN_INTERVAL_MS = 3_000
const PROBE_TIMEOUT_MS = 8_000
// One missed probe is often a slow response or a laptop waking; only a second
// consecutive miss is worth interrupting the lawyer for.
const MISSES_BEFORE_BANNER = 2

type Service = "engine" | "ingest"

// After boot, keep probing both backends. When one stops answering, a quiet
// banner says so and polling continues; it clears itself once the service is
// back, so the lawyer never has to reload to find out.
export default function ConnectionBanner() {
  const [down, setDown] = useState<Service[]>([])
  const misses = useRef<Record<Service, number>>({ engine: 0, ingest: 0 })
  const toast = useToast()

  useEffect(() => {
    let cancelled = false
    let timer: ReturnType<typeof setTimeout> | undefined
    let wasDown = false

    async function tick() {
      if (document.visibilityState === "hidden") return schedule(HEALTHY_INTERVAL_MS)
      const [engine, ingest] = await Promise.all([probe(getConfig), probe(listMatters)])
      if (cancelled) return
      misses.current = { engine: engine ? 0 : misses.current.engine + 1, ingest: ingest ? 0 : misses.current.ingest + 1 }
      const next = (["engine", "ingest"] as const).filter((s) => misses.current[s] >= MISSES_BEFORE_BANNER)
      setDown(next)
      if (wasDown && next.length === 0) toast("success", "Reconnected. Everything is up to date.")
      wasDown = next.length > 0
      schedule(engine && ingest ? HEALTHY_INTERVAL_MS : DOWN_INTERVAL_MS)
    }

    function schedule(ms: number) {
      if (!cancelled) timer = setTimeout(tick, ms)
    }

    schedule(HEALTHY_INTERVAL_MS)
    return () => {
      cancelled = true
      clearTimeout(timer)
    }
  }, [toast])

  if (down.length === 0) return null
  return (
    <div className="conn-banner" role="status" aria-live="polite">
      <span className="conn-banner-dot" aria-hidden />
      <span>
        <strong>Reconnecting…</strong> {describe(down)} Your work is saved; this clears on its own once it is back.
      </span>
    </div>
  )
}

function describe(down: Service[]) {
  if (down.length === 2) return "doc.haus can't reach its services right now."
  return down[0] === "engine" ? "doc.haus can't reach the drafting engine right now." : "doc.haus can't reach the document service right now."
}

async function probe(call: () => Promise<unknown>) {
  const timeout = new Promise<false>((resolve) => setTimeout(() => resolve(false), PROBE_TIMEOUT_MS))
  return Promise.race([call().then(() => true, () => false), timeout])
}
