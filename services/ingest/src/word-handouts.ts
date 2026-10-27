import { spawnSync } from "node:child_process"
import { createHash, randomBytes } from "node:crypto"
import { chmodSync, closeSync, constants, fchmodSync, fstatSync, linkSync, lstatSync, mkdirSync, openSync, readdirSync, readFileSync, renameSync, rmSync, unlinkSync, writeFileSync, type Stats } from "node:fs"
import path from "node:path"
import type { WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { MAX_WORD_UPLOAD_BYTES, readWordLink } from "../../../dochaus/lib/word-linked-docx"
import { windowsPrivateFileCommand } from "../../../dochaus/lib/windows-private-file"
import { readLink } from "../../../apps/word-addin/src/client"

export type WordHandoutIdentity = Pick<WordDocumentLink, "apiOrigin" | "origin" | "matter" | "document" | "baseVersion">
export const wordGrantDigest = (token: string) => createHash("sha256").update(token).digest("hex")
export const MAX_WORD_HANDOUTS = 32
const unavailable = () => new Error("Word handout storage is unavailable.")

function windowsPrivacy(file: string, created: boolean) {
  if (process.platform !== "win32") return
  const { command, env } = windowsPrivateFileCommand(file, created)
  const result = spawnSync(command[0]!, command.slice(1), { env, stdio: "ignore", timeout: 5000, windowsHide: true })
  if (result.status !== 0) throw unavailable()
}

function owned(stat: Stats) {
  if (process.getuid && stat.uid !== process.getuid()) throw unavailable()
}

function documentDirectory(matterDir: string, document: string) {
  if (!/^[a-f0-9-]{36}$/.test(document)) throw unavailable()
  return path.join(matterDir, ".dochaus", "word-handouts", document)
}

function directory(matterDir: string, document: string) {
  const metadata = path.join(matterDir, ".dochaus"), handouts = path.join(metadata, "word-handouts")
  const directory = documentDirectory(matterDir, document)
  // The matter's metadata may have other consumers; only this transport cache
  // receives a restrictive ACL. Never follow a metadata/cache symlink.
  mkdirSync(metadata, { recursive: true })
  if (!lstatSync(metadata).isDirectory() || lstatSync(metadata).isSymbolicLink()) throw unavailable()
  for (const folder of [handouts, directory]) {
    let created = false
    try { mkdirSync(folder, { mode: 0o700 }); created = true }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw unavailable() }
    const stat = lstatSync(folder)
    if (!stat.isDirectory() || stat.isSymbolicLink()) throw unavailable()
    owned(stat)
    // Existing Windows ACLs are checked, never silently rewritten. A newly
    // created folder is protected before any child or document bytes exist.
    windowsPrivacy(folder, created)
    if (process.platform !== "win32") chmodSync(folder, 0o700)
  }
  ignoreSnapshots(handouts)
  return directory
}

function file(matterDir: string, identity: WordHandoutIdentity) {
  const base = wordGrantDigest(identity.baseVersion)
  const key = wordGrantDigest(JSON.stringify([identity.apiOrigin, identity.origin, identity.matter, identity.document, identity.baseVersion]))
  return path.join(directory(matterDir, identity.document), `${base}.${key}.docx`)
}

function privateFile(filename: string) {
  const stat = lstatSync(filename)
  if (!stat.isFile() || stat.isSymbolicLink()) throw unavailable()
  owned(stat)
  windowsPrivacy(filename, false)
  return stat
}

function ignoreSnapshots(folder: string) {
  const filename = path.join(folder, ".gitignore")
  const verify = () => {
    if (privateFile(filename).size !== 2) throw unavailable()
    const descriptor = openSync(filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0))
    try { if (readFileSync(descriptor, "utf8") !== "*\n") throw unavailable() }
    finally { closeSync(descriptor) }
  }
  try { verify(); return }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw unavailable() }
  const temporary = path.join(folder, `.${randomBytes(8).toString("hex")}.gitignore.tmp`)
  let descriptor: number | undefined
  try {
    descriptor = openSync(temporary, "wx", 0o600)
    windowsPrivacy(temporary, true)
    writeFileSync(descriptor, "*\n")
    closeSync(descriptor); descriptor = undefined
    // Atomic publication cannot replace an existing file or symlink.
    try { linkSync(temporary, filename) }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error }
    verify()
  } catch { throw unavailable() }
  finally {
    if (descriptor !== undefined) closeSync(descriptor)
    try { unlinkSync(temporary) } catch { /* Already removed, or never created. */ }
  }
}

/** A private transport copy preserves the exact capability across restart.
 * It is never the matter's document, a Word working file, or a second writer. */
export function readWordHandout(matterDir: string, identity: WordHandoutIdentity, live: (tokenHash: string) => boolean) {
  const filename = file(matterDir, identity)
  let descriptor: number | undefined
  try {
    if (privateFile(filename).size > MAX_WORD_UPLOAD_BYTES) return
    descriptor = openSync(filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0))
    if (process.platform !== "win32") fchmodSync(descriptor, 0o600)
  } catch (error) {
    if (descriptor !== undefined) closeSync(descriptor)
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return
    throw unavailable()
  }
  try {
    const stat = fstatSync(descriptor)
    if (!stat.isFile() || stat.size > MAX_WORD_UPLOAD_BYTES) return
    const bytes = readFileSync(descriptor)
    const link = readLink(readWordLink(bytes), identity.apiOrigin)
    if (!link || !Object.entries(identity).every(([key, value]) => link[key as keyof WordDocumentLink] === value) ||
      !/^[A-Za-z0-9_-]{43}$/.test(link.token) || !live(wordGrantDigest(link.token))) return
    return { bytes, link }
  } catch { return }
  finally { closeSync(descriptor) }
}

/** Keep only the current base and a bounded number of browser-origin handouts.
 * An old ticket safely expires if its transport copy was pruned. */
function prune(filename: string) {
  const folder = path.dirname(filename), prefix = path.basename(filename).split(".")[0] + "."
  const current: { file: string; time: number }[] = []
  for (const entry of readdirSync(folder)) {
    const candidate = path.join(folder, entry)
    if (candidate === filename) continue
    if (!entry.startsWith(prefix) || !/^[a-f0-9]{64}\.[a-f0-9]{64}\.docx$/.test(entry)) {
      // No recursive traversal: an unexpected directory fails closed.
      unlinkSync(candidate)
    } else current.push({ file: candidate, time: lstatSync(candidate).mtimeMs })
  }
  current.sort((a, b) => b.time - a.time)
  for (const entry of current.slice(MAX_WORD_HANDOUTS - 1)) unlinkSync(entry.file)
}

export function writeWordHandout(matterDir: string, identity: WordHandoutIdentity, bytes: Uint8Array) {
  if (bytes.byteLength > MAX_WORD_UPLOAD_BYTES) throw unavailable()
  const filename = file(matterDir, identity), temporary = `${filename}.${randomBytes(8).toString("hex")}.tmp`
  let descriptor: number | undefined
  try {
    try { privateFile(filename) }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error }
    // Secure an EMPTY file before writing a linked document/live capability.
    descriptor = openSync(temporary, "wx", 0o600)
    windowsPrivacy(temporary, true)
    writeFileSync(descriptor, bytes)
    closeSync(descriptor); descriptor = undefined
    renameSync(temporary, filename)
    prune(filename)
  } catch { throw unavailable() }
  finally {
    if (descriptor !== undefined) closeSync(descriptor)
    try { unlinkSync(temporary) } catch { /* Already atomically installed, or never created. */ }
  }
}

/** The existing delete/redaction/startup purge owns transport copies too. */
export function purgeWordHandouts(matterDir: string, document: string, currentBase?: string) {
  const folder = documentDirectory(matterDir, document)
  let entries: string[]
  try {
    // Purge never creates storage or follows a parent symlink outside a matter.
    for (const parent of [path.dirname(path.dirname(folder)), path.dirname(folder), folder]) {
      const stat = lstatSync(parent)
      if (!stat.isDirectory() || stat.isSymbolicLink()) throw unavailable()
      owned(stat)
    }
    entries = readdirSync(folder)
  }
  catch (error) { if ((error as NodeJS.ErrnoException).code === "ENOENT") return 0; throw error }
  const prefix = currentBase ? wordGrantDigest(currentBase) + "." : undefined
  const kept = prefix ? entries.filter(entry => entry.startsWith(prefix) && /^[a-f0-9]{64}\.[a-f0-9]{64}\.docx$/.test(entry)) : []
  if (!kept.length) rmSync(folder, { recursive: true, force: true })
  else for (const entry of entries) if (!kept.includes(entry)) rmSync(path.join(folder, entry), { recursive: true, force: true })
  return entries.length - kept.length
}
