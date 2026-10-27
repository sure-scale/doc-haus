import { afterAll, expect, test } from "bun:test"
import { spawnSync } from "node:child_process"
import { randomBytes, randomUUID } from "node:crypto"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import type { WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { linkedWordDocx } from "../../../dochaus/lib/word-linked-docx"
import { windowsPrivateFileCommand } from "../../../dochaus/lib/windows-private-file"
import { openDb } from "./db"
import { captureVersion, detachVersionedDocument, purgeVersionBytes, sweepRemovedVersionBytes } from "./document-versions"
import { MAX_WORD_HANDOUTS, readWordHandout, wordGrantDigest, writeWordHandout } from "./word-handouts"

const roots: string[] = []
const plain = readFileSync(path.join(import.meta.dir, "fixtures/synthetic-word-accepted.docx"))
const gitAvailable = spawnSync("git", ["--version"], { stdio: "ignore" }).status === 0
afterAll(() => { for (const root of roots) rmSync(root, { recursive: true, force: true }) })
function fixture() {
  const matter = mkdtempSync(path.join(tmpdir(), "dochaus-word-handouts-"))
  roots.push(matter)
  const link: WordDocumentLink = { version: 1, apiOrigin: "https://localhost:4612", origin: "http://localhost:5174",
    matter: randomUUID(), matterName: "Synthetic matter", document: randomUUID(), documentName: "Synthetic advice.docx",
    baseVersion: randomUUID(), token: randomBytes(32).toString("base64url") }
  return { matter, link, bytes: linkedWordDocx(plain, link), folder: path.join(matter, ".dochaus/word-handouts", link.document) }
}
const live = (link: WordDocumentLink) => (hash: string) => hash === wordGrantDigest(link.token)

test("private matter transport survives a new reader only with the same live identity", () => {
  const { matter, link, bytes, folder } = fixture()
  writeWordHandout(matter, link, bytes)
  const cached = readWordHandout(matter, link, live(link))
  expect(Boolean(cached && Buffer.compare(cached.bytes, bytes) === 0 && cached.link.token === link.token)).toBe(true)
  expect(readWordHandout(matter, link, () => false) === undefined).toBe(true)
  expect(readWordHandout(matter, { ...link, baseVersion: randomUUID() }, live(link)) === undefined).toBe(true)
  const files = readdirSync(folder)
  expect(files.length === 1 && /^[a-f0-9]{64}\.[a-f0-9]{64}\.docx$/.test(files[0]!)).toBe(true)
  if (process.platform !== "win32") {
    expect(statSync(folder).mode & 0o777).toBe(0o700)
    expect(statSync(path.join(folder, files[0]!)).mode & 0o777).toBe(0o600)
  }
}, 30_000)

test("new handouts bound current origins and remove every stale base without touching working bytes", () => {
  const { matter, link, bytes, folder } = fixture()
  const working = path.join(matter, link.documentName)
  writeFileSync(working, plain)
  writeWordHandout(matter, link, bytes)
  const base = wordGrantDigest(link.baseVersion)
  // Existing valid private transport entries from earlier origin handouts.
  // Filling them directly keeps this retention test fast on Windows as well.
  for (let index = 0; index < MAX_WORD_HANDOUTS + 2; index++) {
    writeFileSync(path.join(folder, `${base}.${wordGrantDigest(String(index))}.docx`), bytes, { mode: 0o600 })
  }
  writeWordHandout(matter, link, bytes)
  expect(readdirSync(folder).length).toBe(MAX_WORD_HANDOUTS)
  const next = { ...link, baseVersion: randomUUID(), token: randomBytes(32).toString("base64url") }
  writeWordHandout(matter, next, linkedWordDocx(plain, next))
  expect(readdirSync(folder).length).toBe(1)
  expect(readWordHandout(matter, link, live(link)) === undefined).toBe(true)
  expect(Boolean(readWordHandout(matter, next, live(next)))).toBe(true)
  expect(Buffer.compare(readFileSync(working), plain) === 0).toBe(true)
}, 60_000)

test("redaction purges old transport while repeated startup sweeps preserve a fresh current identity until deletion", async () => {
  const { matter, link, folder } = fixture()
  const db = openDb(matter)
  try {
    const first = captureVersion(db, matter, { name: link.documentName, bytes: plain, version: { origin: "upload" } })
    const current = { ...link, document: first.document_id, baseVersion: first.id }
    const cache = path.join(matter, ".dochaus/word-handouts", current.document)
    writeWordHandout(matter, current, linkedWordDocx(plain, current))
    // A redaction uses the same central purge after its registry write commits.
    const redacted = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Neutral synthetic advice: [REDACTED].")] }] }))
    const revision = captureVersion(db, matter, { name: link.documentName, bytes: redacted, prior: plain, version: { origin: "redaction" } })
    purgeVersionBytes(matter, current.document)
    expect(existsSync(cache)).toBe(false)
    const fresh = { ...current, baseVersion: revision.id, token: randomBytes(32).toString("base64url") }
    const freshBytes = linkedWordDocx(redacted, fresh)
    writeWordHandout(matter, fresh, freshBytes)
    for (let restart = 0; restart < 2; restart++) {
      sweepRemovedVersionBytes(matter)
      const retained = readWordHandout(matter, fresh, live(fresh))
      expect(Boolean(retained && retained.link.token === fresh.token && Buffer.compare(retained.bytes, freshBytes) === 0)).toBe(true)
      expect(readWordHandout(matter, current, live(current)) === undefined).toBe(true)
    }
    // A stop after the delete transaction, including an already removed
    // revisions directory, must still be recovered by the startup sweep.
    detachVersionedDocument(db, matter, link.documentName)
    rmSync(path.join(matter, ".dochaus/revisions", current.document), { recursive: true, force: true })
    expect(sweepRemovedVersionBytes(matter)).toBeGreaterThan(0)
    expect(existsSync(cache)).toBe(false)
    expect(existsSync(folder)).toBe(false)
  } finally { db.close() }
  rmSync(matter, { recursive: true, force: true })
  expect(existsSync(matter)).toBe(false)
}, 60_000)

test("purge refuses a metadata parent symlink and leaves bytes outside the matter untouched", () => {
  const { matter, link } = fixture(), external = fixture().matter
  mkdirSync(path.join(matter, ".dochaus"))
  const target = path.join(external, link.document)
  mkdirSync(target)
  const canary = path.join(target, "outside.docx")
  writeFileSync(canary, plain)
  symlinkSync(external, path.join(matter, ".dochaus/word-handouts"), process.platform === "win32" ? "junction" : "dir")
  expect(() => purgeVersionBytes(matter, link.document)).toThrow("storage is unavailable")
  expect(Buffer.compare(readFileSync(canary), plain) === 0).toBe(true)
})

test.skipIf(!gitAvailable)("transport handouts are ignored in a custom Git matter and an existing ignore rule is never overwritten", () => {
  const { matter, link, bytes, folder } = fixture()
  const git = (args: string[]) => spawnSync("git", args, { cwd: matter, stdio: "ignore",
    env: { ...process.env, GIT_CONFIG_NOSYSTEM: "1", GIT_CONFIG_GLOBAL: process.platform === "win32" ? "NUL" : "/dev/null" } }).status
  expect(git(["init", "--quiet"])).toBe(0)
  const future = path.join(folder, `${wordGrantDigest(link.baseVersion)}.${wordGrantDigest("future")}.docx`)
  expect(git(["check-ignore", "--no-index", "--quiet", "--", future])).toBe(1)
  writeWordHandout(matter, link, bytes)
  const file = path.join(folder, readdirSync(folder)[0]!)
  expect(git(["check-ignore", "--no-index", "--quiet", "--", file])).toBe(0)
  const ignore = path.join(path.dirname(folder), ".gitignore")
  writeFileSync(ignore, "!*.docx\n")
  expect(() => writeWordHandout(matter, link, bytes)).toThrow("storage is unavailable")
  expect(readFileSync(ignore, "utf8")).toBe("!*.docx\n")
}, 30_000)

function acl(file: string, created: boolean) {
  const { command, env } = windowsPrivateFileCommand(file, created)
  return spawnSync(command[0]!, command.slice(1), { env, stdio: "ignore", timeout: 5000, windowsHide: true }).status
}
function allowEveryone(file: string, directory = false) {
  const { command, env } = windowsPrivateFileCommand(file, false)
  // A realistic custom/shared storage ACL. All source is fixed; the path is
  // environment data, including spaces and PowerShell punctuation.
  const script = ["$ErrorActionPreference='Stop'", "$p=$env:DOCHAUS_WORD_CACHE_PATH", "$acl=Get-Acl -LiteralPath $p",
    "$everyone=[System.Security.Principal.SecurityIdentifier]::new('S-1-1-0')",
    directory
      ? "$rule=[System.Security.AccessControl.FileSystemAccessRule]::new($everyone,'ReadAndExecute','ContainerInherit,ObjectInherit','None','Allow')"
      : "$rule=[System.Security.AccessControl.FileSystemAccessRule]::new($everyone,'ReadAndExecute','Allow')",
    "$acl.AddAccessRule($rule)", "Set-Acl -LiteralPath $p -AclObject $acl"].join("\n")
  const result = spawnSync(command[0]!, [...command.slice(1, -1), script], { env, stdio: "ignore", timeout: 5000, windowsHide: true })
  expect(result.status).toBe(0)
}

test.skipIf(process.platform !== "win32")("Windows protects new empty cache storage and refuses inherited shared directories or exposed existing files", () => {
  const { matter, link, bytes, folder } = fixture()
  // A shared matter/custom runtime does not make a NEW transport folder shared.
  allowEveryone(matter, true)
  writeWordHandout(matter, link, bytes)
  const file = path.join(folder, readdirSync(folder)[0]!)
  expect(acl(path.dirname(folder), false)).toBe(0)
  expect(acl(folder, false)).toBe(0)
  expect(acl(file, false)).toBe(0)
  expect(Boolean(readWordHandout(matter, link, live(link)))).toBe(true)
  allowEveryone(file)
  expect(() => readWordHandout(matter, link, live(link))).toThrow("storage is unavailable")
  expect(() => writeWordHandout(matter, link, bytes)).toThrow("storage is unavailable")
  expect(acl(file, false) === 0).toBe(false) // No silent ACL repair of existing bytes.
  expect(readdirSync(folder).length).toBe(1)

  const shared = fixture()
  mkdirSync(shared.folder, { recursive: true })
  allowEveryone(path.dirname(shared.folder), true)
  expect(() => writeWordHandout(shared.matter, shared.link, shared.bytes)).toThrow("storage is unavailable")
  expect(readdirSync(shared.folder)).toEqual([]) // No payload or temporary file was written.
  expect(acl(path.dirname(shared.folder), false) === 0).toBe(false)
}, 60_000)
