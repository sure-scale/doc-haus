import path from "node:path"
import { readFileSync, mkdirSync, renameSync } from "node:fs"
import { installDependencies } from "./install-dependencies"
import { dataPaths } from "../services/engine/src/data-paths"

export type Release = { version: number; name: string; notes: string }
export type UpdateStatus = {
  current: Release
  latest?: Release
  target?: string
  available: boolean
  checkedAt?: number
  error?: string
  phase?: "restarting" | "failed"
}

function release(text: string): Release {
  const value = JSON.parse(text)
  if (!Number.isSafeInteger(value.version) || value.version < 1 || typeof value.name !== "string" || typeof value.notes !== "string")
    throw new Error("The release file is invalid.")
  return { version: value.version, name: value.name, notes: value.notes }
}

export class RepoUpdater {
  readonly current: Release
  private status: UpdateStatus
  private checking?: Promise<UpdateStatus>

  constructor(readonly root: string, readonly runtime: string) {
    this.current = release(readFileSync(path.join(root, "release.json"), "utf8"))
    this.status = { current: this.current, available: false }
  }

  async git(...args: string[]) {
    const child = Bun.spawn(["git", ...args], {
      cwd: this.root, stdout: "pipe", stderr: "ignore",
      env: { ...process.env, GIT_TERMINAL_PROMPT: "0" },
    })
    const timer = setTimeout(() => child.kill(), 30_000)
    try {
      const [output, code] = await Promise.all([new Response(child.stdout).text(), child.exited])
      if (code !== 0) throw new Error("Git could not complete the update. Check your repository access and connection.")
      return output.trim()
    } finally {
      clearTimeout(timer)
    }
  }

  private async checkRelease(force = false): Promise<UpdateStatus> {
    if (!this.checking && (force || !this.status.checkedAt || Date.now() - this.status.checkedAt >= 60 * 60_000))
      this.checking = this.inspect().finally(() => { this.checking = undefined })
    return this.checking ? await this.checking : this.status
  }

  async check(force = false): Promise<UpdateStatus> {
    const status = await this.checkRelease(force)
    const operation = await Bun.file(path.join(this.runtime, "update-result.json")).json().catch(() => undefined)
    return operation?.phase === "failed" ? { ...status, phase: "failed", error: operation.error } : status
  }

  private async inspect(): Promise<UpdateStatus> {
    try {
      // Releases always come from the public channel. A task branch's upstream
      // (or detached HEAD) must never choose the code offered to this install.
      await this.git("fetch", "--no-tags", "--", "origin", "refs/heads/dochaus:refs/remotes/origin/dochaus")
      // FETCH_HEAD can be replaced by an unrelated concurrent feature fetch.
      const fetched = await this.git("rev-parse", "refs/remotes/origin/dochaus")
      const target = await this.git("log", "-1", "--format=%H", fetched, "--", "release.json")
      const latest = release(await this.git("show", `${target}:release.json`))
      this.status = { current: this.current, latest, target, available: latest.version > this.current.version, checkedAt: Date.now() }
    } catch (error) {
      // Do not echo git output: remote URLs may contain credentials.
      this.status = { current: this.current, available: false, checkedAt: Date.now(), error: error instanceof Error ? error.message : "Could not check for updates." }
    }
    return this.status
  }

  async requireApplicable(target: string) {
    if (!/^[a-f0-9]{40,64}$/.test(target)) throw new Error("Check for updates again before updating.")
    const changes = await this.git("status", "--porcelain", "--untracked-files=normal")
    if (changes)
      throw new Error("This installation has local changes. Ask your administrator to save them before updating.")
    // Revalidate the offered SHA against the release channel for every apply,
    // including CLI requests. Never replace a stale offer with a newer target.
    const canonical = await this.checkRelease(true)
    if (canonical.error) throw new Error(canonical.error)
    if (!canonical.available || target !== canonical.target)
      throw new Error("The available update changed. Check for updates again.")
    const latest = release(await this.git("show", `${target}:release.json`))
    if (latest.version <= this.current.version) throw new Error("This update is no longer available. Check for updates again.")
    await this.git("merge-base", "--is-ancestor", "HEAD", target).catch(() => {
      throw new Error("This installation has separate code changes. Ask your administrator to update it.")
    })
  }

  async writeResult(value: { phase: "restarting" | "failed"; error?: string }) {
    mkdirSync(this.runtime, { recursive: true })
    const file = path.join(this.runtime, "update-result.json")
    await Bun.write(`${file}.tmp`, JSON.stringify(value))
    renameSync(`${file}.tmp`, file)
  }

  async apply(target: string, install = installDependencies) {
    await this.requireApplicable(target)
    const previous = await this.git("rev-parse", "HEAD")
    await this.writeResult({ phase: "restarting" })
    try {
      // Equivalent to a fast-forward-only pull, pinned to the release the user saw.
      await this.git("merge", "--ff-only", target)
      for (const dir of [this.root, path.join(this.root, "services/ingest")]) {
        if (await install(dir) !== 0) throw new Error("Dependencies could not be installed. The previous version has been restored; check your connection and try again.")
      }
    } catch (error) {
      // --keep refuses to overwrite edits made during installation.
      await this.git("reset", "--keep", previous)
      for (const dir of [this.root, path.join(this.root, "services/ingest")]) await install(dir)
      const message = error instanceof Error ? error.message : "The update failed. Please try again."
      await this.writeResult({ phase: "failed", error: message })
      throw new Error(message)
    }
  }
}

if (import.meta.main) {
  const root = path.resolve(import.meta.dir, "..")
  const updater = new RepoUpdater(root, dataPaths(root).runtime)
  if (process.argv[2] === "--apply") {
    const request = await Bun.file(path.join(updater.runtime, "update-request.json")).json()
    await updater.apply(request.target).catch(async error => {
      await updater.writeResult({ phase: "failed", error: error.message })
      process.exitCode = 1
    })
  } else {
    do {
      console.log(JSON.stringify(await updater.check(true)))
      if (process.argv[2] !== "--watch") break
      await Bun.sleep(60 * 60_000)
    } while (true)
  }
}
