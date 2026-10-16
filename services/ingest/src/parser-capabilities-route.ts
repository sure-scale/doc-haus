import { Hono } from "hono"
import { inspectParserCapabilities, type ParserCapabilitiesSnapshot } from "./parser-capabilities"

export function parserCapabilitiesRoutes(options: {
  diskDirectory?: string
  inspect?: () => Promise<ParserCapabilitiesSnapshot>
} = {}) {
  const app = new Hono()
  app.get("/parser-capabilities", async c => {
    c.header("Cache-Control", "no-store")
    try {
      return c.json(await (options.inspect ?? (() => inspectParserCapabilities({ diskDirectory: options.diskDirectory })))())
    } catch {
      return c.json({ error: "The ingestion host inventory is unavailable." }, 503)
    }
  })
  return app
}
