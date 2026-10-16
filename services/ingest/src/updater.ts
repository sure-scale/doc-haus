import path from "node:path"
import { Hono } from "hono"
import { RepoUpdater } from "../../../scripts/updater"

export function updateRoutes(updater: RepoUpdater, busy: () => Promise<boolean>, restart: () => Promise<void>, signal = () => { process.kill(Number(process.env.DOCHAUS_LAUNCHER_PID), "SIGUSR1") }) {
  const app = new Hono()
  let updating = false
  let mutations = 0
  app.use("*", async (c, next) => {
    if (["POST", "PUT", "PATCH", "DELETE"].includes(c.req.method) && c.req.path !== "/updates/apply") {
      if (updating) return c.json({ error: "Doc.Haus is updating. Please wait for it to restart." }, 503)
      mutations++
      try { await next() } finally { mutations-- }
      return
    }
    await next()
  })
  app.get("/updates", async c => c.json(await updater.check(c.req.query("refresh") === "1")))
  app.post("/updates/apply", async c => {
    // Require a non-simple request and a local browser origin, including when
    // ingest is exposed for other API consumers. CORS alone does not block POSTs.
    const origin = c.req.header("origin")
    if (c.req.header("x-dochaus-update") !== "1" || !origin || !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin))
      return c.json({ error: "Updates must be started from Doc.Haus on this computer." }, 403)
    if (updating) return c.json({ error: "An update is already in progress." }, 409)
    updating = true
    try {
      if (mutations || await busy()) throw new Error("Work is still in progress. Wait for uploads, drafts and answers to finish, then try again.")
      const input = await c.req.json()
      const status = await updater.check()
      if (!status.available || input.target !== status.target) throw new Error("The available update changed. Check for updates again.")
      await updater.requireApplicable(input.target)
      // Validate the managed owner before writing a request or signalling it.
      await restart()
      await Bun.write(path.join(updater.runtime, "update-request.json"), JSON.stringify({ target: input.target }))
      await updater.writeResult({ phase: "restarting" })
      // Return the accepted response before the launcher stops this service.
      setTimeout(() => {
        try { signal() } catch {
          updating = false
          void updater.writeResult({ phase: "failed", error: "The launcher stopped before updating. Start Doc.Haus again and retry." })
        }
      }, 500)
      return c.json({ phase: "restarting" }, 202)
    } catch (error) {
      updating = false
      return c.json({ error: error instanceof Error ? error.message : "Could not start the update." }, 409)
    }
  })
  return app
}

export async function managedUpdateOwner(root: string, runtime: string) {
  const owner = Number(process.env.DOCHAUS_LAUNCHER_PID)
  const pid = (await Bun.file(path.join(runtime, "engine.lock/pid")).text().catch(() => "")).trim()
  if (!Number.isSafeInteger(owner) || owner <= 1 || pid !== String(owner))
    throw new Error("Start Doc.Haus with start.sh to enable one-click updates.")
  process.kill(owner, 0)
  const child = Bun.spawn(["bash", "-c", '. scripts/start-runtime.sh; runtime_launcher "$1"', "update-owner", String(owner)], { cwd: root, stdout: "ignore", stderr: "ignore" })
  if (await child.exited !== 0) throw new Error("The app launcher could not be identified. Restart Doc.Haus with start.sh.")
}
