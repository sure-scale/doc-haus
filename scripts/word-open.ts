// Native URI handoff only. The Word add-in remains the canonical save path.
import { createHash, randomBytes } from "node:crypto"
import { constants } from "node:fs"
import { chmod, link, lstat, mkdir, open, unlink } from "node:fs/promises"
import path from "node:path"
import { readWordLink } from "../dochaus/lib/word-linked-docx"
import { type WordDocumentLink } from "../dochaus/lib/word-addin-contract"
import { readLink } from "../apps/word-addin/src/client"
import { windowsPrivateFileCommand } from "../dochaus/lib/windows-private-file"

const MAX_BYTES = 128 * 1024 * 1024
const invalid = "This doc.haus Word link is not valid. Choose Open in Word again."
const downloadFailed = "doc.haus could not download this Word document. Check the connection and Word setup, then choose Open in Word again."
const packageFailed = "The Word document did not match this doc.haus connection. Choose Open in Word again."
const cacheFailed = "doc.haus could not use its private Word checkout. Your existing files were kept."

export class WordOpenError extends Error {}

export function parseWordOpenUri(uri: string): { source: string; origin: string } {
  try {
    const match = /^doc\.haus:\/\/word\/open\?source=([^&#]+)$/.exec(uri)
    if (!match) throw 0
    const source = decodeURIComponent(match[1]!)
    if (encodeURIComponent(source) !== match[1]) throw 0
    const url = new URL(source)
    if (url.protocol !== "https:" || url.username || url.password || url.search || url.hash ||
      !/^\/word\/open\/[A-Za-z0-9_-]{43}$/.test(url.pathname) || source !== `${url.origin}${url.pathname}`) throw 0
    return { source, origin: url.origin }
  } catch { throw new WordOpenError(invalid) }
}

/** Use the save path's bounded package reader and the pane's link validation. */
export function wordLink(bytes: Uint8Array, origin: string): WordDocumentLink {
  try {
    const value = readLink(readWordLink(bytes), origin)
    if (!value || value.apiOrigin !== origin) throw 0
    return value
  } catch { throw new WordOpenError(packageFailed) }
}

type CommandOptions = { input?: string; env?: NodeJS.ProcessEnv; limit: number; timeout: number }
type CommandResult = { bytes: Uint8Array; status: number; error: string }
export type WordCommandRunner = (command: string[], options: CommandOptions) => Promise<CommandResult>

async function bounded(stream: ReadableStream<Uint8Array>, limit: number) {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  let size = 0
  while (true) {
    const next = await reader.read()
    if (next.done) break
    size += next.value.byteLength
    if (size > limit) { await reader.cancel(); throw new WordOpenError(downloadFailed) }
    chunks.push(next.value)
  }
  return Buffer.concat(chunks, size)
}

const run: WordCommandRunner = async (command, options) => {
  const child = Bun.spawn(command, { stdin: options.input === undefined ? "ignore" : "pipe", stdout: "pipe", stderr: "pipe",
    env: options.env, windowsHide: true, timeout: options.timeout })
  if (options.input !== undefined && typeof child.stdin !== "number" && child.stdin) {
    child.stdin.write(options.input)
    child.stdin.end()
  }
  try {
    const [bytes, error, status] = await Promise.all([bounded(child.stdout, options.limit), bounded(child.stderr, 1024), child.exited])
    return { bytes, error: new TextDecoder().decode(error), status }
  } catch (error) { child.kill(); throw error }
}

function systemExecutable(platform: NodeJS.Platform, name: "curl" | "powershell") {
  if (platform === "darwin" && name === "curl") return "/usr/bin/curl"
  if (platform === "win32") return path.win32.join(process.env.SystemRoot || "C:\\Windows", "System32",
    name === "curl" ? "curl.exe" : "WindowsPowerShell\\v1.0\\powershell.exe")
  throw new WordOpenError("Open this Word link using the installed doc.haus app on Mac or Windows.")
}

export async function downloadWordDocument(source: string, options: { platform?: NodeJS.Platform; run?: WordCommandRunner } = {}) {
  try {
    // Source is never shell text or a child-process argument. Ignore curlrc,
    // keep normal system TLS trust, and never follow a ticket redirect.
    const parsed = parseWordOpenUri(`doc.haus://word/open?source=${encodeURIComponent(source)}`)
    const result = await (options.run ?? run)([systemExecutable(options.platform ?? process.platform, "curl"), "--disable", "--config", "-",
      "--silent", "--fail", "--proto", "=https", "--proto-redir", "=https", "--max-redirs", "0", "--connect-timeout", "10", "--max-time", "60",
      "--max-filesize", String(MAX_BYTES), "--write-out", "%{stderr}%{http_code}", "--output", "-"],
    { input: `url = "${parsed.source}"\n`, limit: MAX_BYTES, timeout: 65_000 })
    if (["404", "410"].includes(result.error)) throw new WordOpenError("This Word link expired or was already used. Choose Open in Word again.")
    if (result.status !== 0 || result.error !== "200" || !result.bytes.length || result.bytes.length > MAX_BYTES) throw new WordOpenError(downloadFailed)
    return result.bytes
  } catch (error) { throw error instanceof WordOpenError ? error : new WordOpenError(downloadFailed) }
}

/** Dispatch feedback only; this never claims that Word loaded or saved. */
export async function reportWordOpenResult(source: string, state: "opened" | "failed", options: { platform?: NodeJS.Platform; run?: WordCommandRunner } = {}) {
  try {
    const parsed = parseWordOpenUri(`doc.haus://word/open?source=${encodeURIComponent(source)}`)
    const result = await (options.run ?? run)([systemExecutable(options.platform ?? process.platform, "curl"), "--disable", "--config", "-",
      "--silent", "--fail", "--proto", "=https", "--proto-redir", "=https", "--max-redirs", "0", "--connect-timeout", "3", "--max-time", "5",
      "--max-filesize", "4096", "--header", "Content-Type: application/json", "--data-binary", JSON.stringify({ state }),
      "--write-out", "%{stderr}%{http_code}", "--output", "-"],
    { input: `url = "${parsed.source}/result"\n`, limit: 4096, timeout: 6000 })
    return result.status === 0 && ["200", "204"].includes(result.error)
  } catch { return false }
}

const digest = (value: string) => createHash("sha256").update(value).digest("hex")
export function wordCheckoutIdentity(value: WordDocumentLink) {
  return digest(JSON.stringify([value.apiOrigin, value.matter, value.document, value.baseVersion, digest(value.token)]))
}

function filename(name: string) {
  let safe = name.replace(/[\u0000-\u001f\u007f<>:"/\\|?*]/g, "_").replace(/[. ]+$/g, "").slice(0, 120)
  if (!safe || /^(con|prn|aux|nul|com[0-9]|lpt[0-9])(?:\.|$)/i.test(safe)) safe = "Document"
  return /\.docx$/i.test(safe) ? safe : `${safe}.docx`
}

type NativeOptions = { platform?: NodeJS.Platform; run?: WordCommandRunner }
async function privateWindowsFile(file: string, created: boolean, options: NativeOptions) {
  if ((options.platform ?? process.platform) !== "win32") return
  const { command, env } = windowsPrivateFileCommand(file, created)
  const result = await (options.run ?? run)(command, { env, limit: 4096, timeout: 5000 })
  if (result.status !== 0) throw 0
}

async function privateDirectory(directory: string, options: NativeOptions) {
  const created = await mkdir(directory, { mode: 0o700 }).then(() => true).catch(error => { if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error; return false })
  const info = await lstat(directory)
  if (!info.isDirectory() || info.isSymbolicLink() || (process.getuid && info.uid !== process.getuid())) throw 0
  if ((options.platform ?? process.platform) !== "win32") await chmod(directory, 0o700)
  await privateWindowsFile(directory, created, options)
}

async function existing(file: string, options: NativeOptions) {
  try {
    const info = await lstat(file)
    if (!info.isFile() || info.isSymbolicLink() || (process.getuid && info.uid !== process.getuid())) throw 0
    await privateWindowsFile(file, false, options)
    return true
  } catch (error) { if ((error as NodeJS.ErrnoException)?.code === "ENOENT") return false; throw error }
}

async function checkout(runtime: string, bytes: Uint8Array, value: WordDocumentLink, options: NativeOptions) {
  try {
    await mkdir(runtime, { recursive: true, mode: 0o700 })
    const word = path.join(runtime, "word"), cache = path.join(word, "checkout"), directory = path.join(cache, wordCheckoutIdentity(value))
    for (const parent of [word, cache, directory]) await privateDirectory(parent, options)
    const file = path.join(directory, filename(value.documentName))
    if (await existing(file, options)) return { path: file, reused: true }
    const temporary = path.join(directory, `.incoming-${randomBytes(12).toString("hex")}`)
    const handle = await open(temporary, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | (constants.O_NOFOLLOW || 0), 0o600)
    try {
      // An empty file exists here, but no document bytes are written until its
      // Windows ACL is private. POSIX creation already set mode 0600.
      await privateWindowsFile(temporary, true, options)
      await handle.writeFile(bytes)
      await handle.sync()
      await handle.close()
      // Atomic no-replace publication: overlapping opens can never truncate
      // an existing checkout, including a file Word has already changed.
      try { await link(temporary, file) } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST" || !await existing(file, options)) throw error
        return { path: file, reused: true }
      }
      return { path: file, reused: false }
    } finally { await handle.close().catch(() => {}); await unlink(temporary).catch(() => {}) }
  } catch { throw new WordOpenError(cacheFailed) }
}

export async function openNativeWord(file: string, options: { platform?: NodeJS.Platform; run?: WordCommandRunner } = {}) {
  const platform = options.platform ?? process.platform
  const command = platform === "darwin" ? ["/usr/bin/open", "-b", "com.microsoft.Word", file] :
    [systemExecutable(platform, "powershell"), "-NoProfile", "-NonInteractive", "-Command", "Start-Process -FilePath $env:DOCHAUS_WORD_FILE -ErrorAction Stop"]
  try {
    const result = await (options.run ?? run)(command, { env: { ...process.env, DOCHAUS_WORD_FILE: file }, limit: 4096, timeout: 10_000 })
    if (result.status !== 0) throw 0
  } catch { throw new WordOpenError("doc.haus kept the Word checkout but could not open Word. Check that Microsoft Word is installed, then choose Open in Word again.") }
}

export async function openWordDocument(options: NativeOptions & { uri: string; runtime: string; download?: (source: string) => Promise<Uint8Array>; open?: (file: string) => Promise<void>; report?: (source: string, state: "opened" | "failed") => Promise<unknown> }) {
  const { source, origin } = parseWordOpenUri(options.uri)
  if (!path.isAbsolute(options.runtime) || options.runtime.includes("\0")) throw new WordOpenError(cacheFailed)
  const report = async (state: "opened" | "failed") => {
    try { await (options.report ?? ((source, state) => reportWordOpenResult(source, state, options)))(source, state) } catch {}
  }
  try {
    const bytes = await (options.download ?? (source => downloadWordDocument(source, options)))(source)
    const value = wordLink(bytes, origin)
    const result = await checkout(options.runtime, bytes, value, options)
    await (options.open ?? (file => openNativeWord(file, options)))(result.path)
    await report("opened")
    return result
  } catch (error) {
    await report("failed")
    throw error instanceof WordOpenError ? error : new WordOpenError("doc.haus could not open this Word document. Your existing checkout was kept.")
  }
}

if (import.meta.main) {
  try {
    const args = process.argv.slice(2)
    if (args.length !== 3 || args[0] !== "--runtime") throw new WordOpenError(invalid)
    await openWordDocument({ runtime: args[1]!, uri: args[2]! })
  } catch (error) {
    // Native failures may include the ticket, grant, file path or document text.
    console.error(error instanceof WordOpenError ? error.message : "doc.haus could not open this Word document. Choose Open in Word again.")
    process.exitCode = 1
  }
}
