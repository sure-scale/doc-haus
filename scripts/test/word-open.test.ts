import { afterEach, expect, test } from "bun:test"
import { mkdtemp, readFile, readdir, rm, stat, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { linkedWordDocx } from "../../dochaus/lib/word-linked-docx"
import { type WordDocumentLink } from "../../dochaus/lib/word-addin-contract"
import { downloadWordDocument, openNativeWord, openWordDocument, parseWordOpenUri, wordCheckoutIdentity, wordLink, reportWordOpenResult, type WordCommandRunner } from "../word-open"

const origin = "https://localhost:4612"
const source = `${origin}/word/open/${"a".repeat(43)}`
const uri = `doc.haus://word/open?source=${encodeURIComponent(source)}`
const link: WordDocumentLink = { version: 1, apiOrigin: origin, origin: "http://localhost:5274", matter: "synthetic-matter", matterName: "Synthetic matter",
  document: "synthetic-document", documentName: "Sample advice.docx", baseVersion: "synthetic-version", token: "synthetic-capability" }
const plain = await Bun.file(path.resolve(import.meta.dir, "../../services/ingest/src/fixtures/synthetic-word-accepted.docx")).bytes()
const handout = linkedWordDocx(plain, link)
const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })
async function runtime() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-word-open-"))
  roots.push(root)
  return path.join(root, "runtime")
}

const openDocument = (options: Parameters<typeof openWordDocument>[0]) => openWordDocument({ report: async () => {}, ...options })

test("Word URI accepts a canonical local or remote scoped ticket and rejects alternative actions, parameters and URL tricks before I/O", async () => {
  expect(parseWordOpenUri(uri)).toEqual({ source, origin })
  const remote = `https://documents.example.test:8443/word/open/${"_".repeat(43)}`
  expect(parseWordOpenUri(`doc.haus://word/open?source=${encodeURIComponent(remote)}`).source).toBe(remote)
  let downloaded = false
  const invalid = ["doc.haus://launch", uri + "#fragment", uri + "&source=x", uri + "&other=x", uri.replace("/word/", "/WORD/"),
    uri.replace("doc.haus://", "https://"), uri.replace("source=", "SOURCE="), uri.replace("source=", "source=%"),
    ...[source.replace("https:", "http:"), source + "?query=x", source + "#hash", source.replace("localhost", "user:password@localhost"),
      source.replace("/word/open/", "/word/other/"), source + "a", source.replace("/word/open/", "/word/../word/open/"),
      source.replace("localhost", "localhost\\attacker.test"), source + "\n"].map(value => `doc.haus://word/open?source=${encodeURIComponent(value)}`)]
  for (const value of invalid) await expect(openDocument({ uri: value, runtime: "/unused", download: async () => { downloaded = true; return handout }, open: async () => {} })).rejects.toThrow("not valid")
  expect(downloaded).toBe(false)
})

test("native download uses system trust, private stdin URL, a bounded deadline and refuses redirects and non-200 content", async () => {
  const calls: { command: string[]; options: Parameters<WordCommandRunner>[1] }[] = []
  const runner: WordCommandRunner = async (command, options) => { calls.push({ command, options }); return { bytes: handout, status: 0, error: "200" } }
  expect(await downloadWordDocument(source, { platform: "darwin", run: runner })).toEqual(handout)
  const call = calls[0]!
  expect(call.command[0]).toBe("/usr/bin/curl")
  expect(call.command[1]).toBe("--disable")
  expect(call.command).not.toContain(source)
  expect(call.options.input).toBe(`url = "${source}"\n`)
  expect(call.options.limit).toBe(128 * 1024 * 1024)
  expect(call.options.timeout).toBeLessThanOrEqual(65_000)
  expect(call.command).not.toContain("--insecure")
  expect(call.command).not.toContain("-k")
  expect(call.command).not.toContain("--location")
  expect(call.command).toContain("=https")
  for (const [status, http] of [[0, "302"], [0, "307"], [22, "403"], [60, "000"], [28, "200"]] as const) {
    await expect(downloadWordDocument(source, { platform: "darwin", run: async () => ({ bytes: handout, status, error: http }) })).rejects.toThrow("could not download")
  }
  await expect(downloadWordDocument(source, { platform: "win32", run: async (command) => {
    expect(command[0]).toMatch(/Windows\\System32\\curl\.exe$/i)
    return { bytes: handout, status: 22, error: "410" }
  } })).rejects.toThrow("expired or was already used")
})

test("native handout must be a valid linked DOCX for exactly the source HTTPS instance", () => {
  expect(wordLink(handout, origin)).toEqual(link)
  for (const bytes of [plain, new TextEncoder().encode("not a DOCX"), linkedWordDocx(plain, { ...link, apiOrigin: "https://elsewhere.example.test" }),
    linkedWordDocx(plain, { ...link, origin: "http://localhost:5274/private" })]) expect(() => wordLink(bytes, origin)).toThrow("did not match")
})

test("first open creates a private unchanged DOCX; repeated opens preserve local edits and modification time", async () => {
  const directory = await runtime()
  const opened: string[] = []
  const options = { uri, runtime: directory, download: async () => handout, open: async (file: string) => { opened.push(file) } }
  const first = await openDocument(options)
  expect(first.reused).toBe(false)
  expect(await readFile(first.path)).toEqual(Buffer.from(handout))
  expect(path.basename(first.path)).toBe("Sample advice.docx")
  expect(first.path).not.toContain(link.token)
  expect(first.path).not.toContain(link.document)
  if (process.platform !== "win32") {
    expect((await stat(first.path)).mode & 0o777).toBe(0o600)
    expect((await stat(path.dirname(first.path))).mode & 0o777).toBe(0o700)
    expect((await stat(path.dirname(path.dirname(first.path)))).mode & 0o777).toBe(0o700)
  }
  const edits = linkedWordDocx(plain, { ...link, matterName: "Synthetic locally changed settings" })
  await writeFile(first.path, edits)
  const before = await stat(first.path)
  expect(await openDocument(options)).toEqual({ path: first.path, reused: true })
  expect(await readFile(first.path)).toEqual(Buffer.from(edits))
  expect((await stat(first.path)).mtimeMs).toBe(before.mtimeMs)
  expect(opened).toEqual([first.path, first.path])
  expect(await readdir(path.dirname(first.path))).toEqual(["Sample advice.docx"])
})

test("concurrent requests atomically publish one checkout and never overwrite it", async () => {
  const directory = await runtime()
  const results = await Promise.all(Array.from({ length: 4 }, () => openDocument({ uri, runtime: directory, download: async () => handout, open: async () => {} })))
  expect(new Set(results.map(result => result.path)).size).toBe(1)
  expect(results.filter(result => !result.reused)).toHaveLength(1)
  expect(await readFile(results[0]!.path)).toEqual(Buffer.from(handout))
  expect(await readdir(path.dirname(results[0]!.path))).toEqual(["Sample advice.docx"])
})

test("instance, matter, document, version and grant identities retain separate checkouts without exposing capabilities", async () => {
  const directory = await runtime()
  const identities = [link, { ...link, apiOrigin: "https://remote.example.test" }, { ...link, matter: "other-matter" },
    { ...link, document: "other-document" }, { ...link, baseVersion: "later-version" }, { ...link, token: "replacement-capability" }]
  expect(new Set(identities.map(wordCheckoutIdentity)).size).toBe(identities.length)
  const paths = []
  for (const value of identities) {
    const nextURI = `doc.haus://word/open?source=${encodeURIComponent(value.apiOrigin + "/word/open/" + "b".repeat(43))}`
    paths.push((await openDocument({ uri: nextURI, runtime: directory, download: async () => linkedWordDocx(plain, value), open: async () => {} })).path)
  }
  expect(new Set(paths).size).toBe(identities.length)
  for (const file of paths) expect(path.basename(path.dirname(file))).toMatch(/^[a-f0-9]{64}$/)
})

test("unsafe names stay inside the private cache and symlink targets are refused without touching other files", async () => {
  const directory = await runtime()
  const dangerous = { ...link, documentName: "../../a|b?.docx" }
  const result = await openDocument({ uri, runtime: directory, download: async () => linkedWordDocx(plain, dangerous), open: async () => {} })
  expect(path.dirname(result.path)).toBe(path.join(directory, "word", "checkout", wordCheckoutIdentity(dangerous)))
  expect(path.basename(result.path)).toBe(".._.._a_b_.docx")
  if (process.platform === "win32") return // Creating a Windows symlink requires separate OS privileges.
  const target = path.join(path.dirname(directory), "preserved.docx")
  await writeFile(target, plain)
  await rm(result.path)
  await symlink(target, result.path)
  let opened = false
  await expect(openDocument({ uri, runtime: directory, download: async () => linkedWordDocx(plain, dangerous), open: async () => { opened = true } })).rejects.toThrow("existing files were kept")
  expect(opened).toBe(false)
  expect(await readFile(target)).toEqual(Buffer.from(plain))
})

test("native Word open passes the local path as data, and a failed opener keeps the checkout", async () => {
  const file = path.join(os.tmpdir(), "word $(echo should-not-run) & quote's.docx")
  const runner: WordCommandRunner = async (command, options) => {
    if (command[0] === "/usr/bin/open") expect(command).toEqual(["/usr/bin/open", "-b", "com.microsoft.Word", file])
    else {
      expect(command.at(-1)).toBe("Start-Process -FilePath $env:DOCHAUS_WORD_FILE -ErrorAction Stop")
      expect(command).not.toContain(file)
      expect(options.env?.DOCHAUS_WORD_FILE).toBe(file)
    }
    return { bytes: new Uint8Array(), status: 0, error: "" }
  }
  await openNativeWord(file, { platform: "darwin", run: runner })
  await openNativeWord(file, { platform: "win32", run: runner })
  const directory = await runtime()
  let saved = ""
  await expect(openDocument({ uri, runtime: directory, download: async () => handout, open: async file => {
    saved = file
    await openNativeWord(file, { platform: "darwin", run: async () => ({ bytes: new Uint8Array(), status: 1, error: "private native exception" }) })
  } })).rejects.toThrow("kept the Word checkout")
  expect(await readFile(saved)).toEqual(Buffer.from(handout))
})

test("helper CLI rejects malformed input without creating runtime data or echoing supplied secrets", async () => {
  const directory = await runtime()
  const supplied = "doc.haus://word/open?source=private-fake-ticket&token=private-fake-credential"
  const child = Bun.spawn([process.execPath, path.resolve(import.meta.dir, "../word-open.ts"), "--runtime", directory, supplied], { stdout: "pipe", stderr: "pipe" })
  const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
  expect(status).toBe(1)
  expect(stdout).toBe("")
  expect(stderr).toContain("not valid")
  expect(stderr).not.toContain("private-fake")
  expect(await Bun.file(path.join(directory, "word")).exists()).toBe(false)
})

test("Windows cache ACLs are restricted before document bytes persist; reuse checks but never replaces existing ACLs", async () => {
  const directory = await runtime()
  const checks: { file: string; created: string }[] = []
  const runner: WordCommandRunner = async (command, options) => {
    const file = options.env!.DOCHAUS_WORD_CACHE_PATH!
    const created = options.env!.DOCHAUS_WORD_CACHE_NEW!
    checks.push({ file, created })
    expect(command[0]).toMatch(/Windows\\System32\\WindowsPowerShell\\v1\.0\\powershell\.exe$/i)
    expect(command).not.toContain(file)
    expect(command.at(-1)).toContain("SetAccessRuleProtection($true,$false)")
    if (created === "1" && path.basename(file).startsWith(".incoming-")) expect((await stat(file)).size).toBe(0)
    return { bytes: new Uint8Array(), status: 0, error: "" }
  }
  const options = { uri, runtime: directory, platform: "win32" as const, run: runner, download: async () => handout, open: async () => {} }
  const first = await openDocument(options)
  expect(checks.filter(check => check.created === "1")).toHaveLength(4)
  checks.length = 0
  expect((await openDocument(options)).reused).toBe(true)
  expect(checks).toHaveLength(4)
  expect(checks.every(check => check.created === "0")).toBe(true)
  expect(checks.at(-1)!.file).toBe(first.path)
  await expect(openDocument({ ...options, run: async () => ({ bytes: new Uint8Array(), status: 1, error: "" }) })).rejects.toThrow("existing files were kept")
  expect(await readFile(first.path)).toEqual(Buffer.from(handout))
})

test("Word dispatch reports only its bounded fixed result and cannot turn callback failure into an open failure", async () => {
  const directory = await runtime()
  const states: string[] = []
  const options = { uri, runtime: directory, download: async () => handout, open: async () => {}, report: async (sourceURL: string, state: "opened" | "failed") => {
    expect(sourceURL).toBe(source)
    states.push(state)
    throw new Error("synthetic callback outage")
  } }
  expect((await openDocument(options)).reused).toBe(false)
  await expect(openDocument({ ...options, open: async () => { throw new Error("synthetic native failure") } })).rejects.toThrow("checkout was kept")
  expect(states).toEqual(["opened", "failed"])
  expect(await reportWordOpenResult(source, "opened", { platform: "darwin", run: async (command, options) => {
    expect(options.input).toBe(`url = "${source}/result"\n`)
    expect(options.timeout).toBeLessThanOrEqual(6000)
    expect(options.limit).toBe(4096)
    expect(command).toContain('{"state":"opened"}')
    expect(command.join(" ")).not.toContain(source)
    expect(command).not.toContain("--location")
    expect(command).not.toContain("--insecure")
    return { bytes: new Uint8Array(), status: 0, error: "204" }
  } })).toBe(true)
  expect(await reportWordOpenResult(source, "failed", { platform: "darwin", run: async () => { throw new Error("synthetic transport failure") } })).toBe(false)
})
