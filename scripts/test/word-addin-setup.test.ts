import { afterEach, expect, test } from "bun:test"
import { X509Certificate, createPrivateKey, createPublicKey } from "node:crypto"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { prepareLocalCertificate, syncWordAddinRegistration, validateWordCertificate } from "../setup-word-addin"
import { renderManifest } from "../../apps/word-addin/src/manifest"
import { WORD_ADDIN_ID } from "../../dochaus/lib/word-addin-contract"

const directories: string[] = []
afterEach(() => { for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true }) })

test("local TLS key matches a loopback-only server leaf and setup reuses valid material", async () => {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-tls-"))
  directories.push(directory)
  const paths = await prepareLocalCertificate(directory)
  const certificate = new X509Certificate(readFileSync(paths.certificate))
  expect(certificate.ca).toBe(false)
  expect(certificate.checkHost("localhost")).toBe("localhost")
  expect(certificate.checkIP("127.0.0.1")).toBe("127.0.0.1")
  expect(certificate.checkIP("::1")).toBe("::1")
  expect(certificate.checkHost("example.org")).toBeUndefined()
  expect(certificate.publicKey.export({ type: "spki", format: "pem" })).toBe(
    createPublicKey(createPrivateKey(readFileSync(paths.key))).export({ type: "spki", format: "pem" }),
  )
  if (process.platform !== "win32") expect(statSync(paths.key).mode & 0o777).toBe(0o600)
  const before = readFileSync(paths.certificate, "utf8")
  await prepareLocalCertificate(directory)
  expect(readFileSync(paths.certificate, "utf8")).toBe(before)
  expect(() => validateWordCertificate(paths.certificate, paths.key, "example.org")).toThrow()
  // A partial previous setup can be repaired without trusting corrupt material.
  writeFileSync(paths.key, "incomplete key")
  await prepareLocalCertificate(directory)
  expect(() => validateWordCertificate(paths.certificate, paths.key)).not.toThrow()
}, 15000)

test("Windows startup switches an existing doc.haus registration to this configured installation", async () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-sync-win-"))
  directories.push(root)
  const config = path.join(root, "runtime/config")
  const addin = path.join(config, "word-addin")
  const tls = await prepareLocalCertificate(addin)
  writeFileSync(path.join(config, "word-addin.json"), JSON.stringify({ origin: "https://localhost:4511", external: false, ...tls }))
  const old = path.join(root, "old.xml")
  const local = path.join(root, "local")
  mkdirSync(path.join(local, "doc.haus"), { recursive: true })
  writeFileSync(path.join(local, "doc.haus/launch-profile.json"), JSON.stringify({ repository: path.resolve(import.meta.dir, "../.."), environment: { DOCHAUS_DATA_ROOT: root } }))
  writeFileSync(old, renderManifest("https://localhost:4611"))
  writeFileSync(path.join(addin, "manifest.xml"), renderManifest("https://localhost:4511"))
  const calls: string[] = []
  const result = await syncWordAddinRegistration({ platform: "win32", home: root, environment: { DOCHAUS_DATA_ROOT: root, LOCALAPPDATA: local },
    command: async (args) => {
      const script = args.at(-1) ?? ""
      calls.push(script)
      return { code: 0, stdout: script.includes("Get-ItemPropertyValue") ? old : "" }
    } })
  expect(result).toBe(true)
  expect(calls).toHaveLength(3)
  expect(calls[1]).toContain("Cert:\\CurrentUser\\Root")
  expect(calls[2]).toContain("New-ItemProperty")
  expect(calls.join(" ")).not.toContain("addstore")
}, 15000)

test("startup refreshes an existing trusted registration for the active origin without changing trust", async () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-sync-"))
  directories.push(root)
  const config = path.join(root, "runtime/config")
  const addin = path.join(config, "word-addin")
  const registered = path.join(root, "Library/Containers/com.microsoft.Word/Data/Documents/wef", `doc-haus-${WORD_ADDIN_ID}.xml`)
  mkdirSync(path.dirname(registered), { recursive: true })
  const metadata = path.join(root, "Applications/doc.haus.app/Contents/Resources/installation.json")
  mkdirSync(path.dirname(metadata), { recursive: true })
  writeFileSync(metadata, JSON.stringify({ repository: path.resolve(import.meta.dir, "../.."), runtime: path.join(root, "runtime") }))
  const tls = await prepareLocalCertificate(addin)
  writeFileSync(path.join(config, "word-addin.json"), JSON.stringify({ origin: "https://localhost:4511", external: false, ...tls }))
  const manifest = path.join(addin, "manifest.xml")
  writeFileSync(manifest, renderManifest("https://localhost:4511"))
  writeFileSync(registered, renderManifest("https://localhost:4611"))
  const calls: string[][] = []
  const options = { platform: "darwin" as const, home: root, environment: { DOCHAUS_DATA_ROOT: root, DOCHAUS_WORD_ORIGIN: "https://localhost:4711" },
    command: async (args: string[]) => { calls.push(args); return { code: 0, stdout: "" } } }
  expect(await syncWordAddinRegistration(options)).toBe(true)
  expect(readFileSync(registered, "utf8")).toBe(renderManifest("https://localhost:4711"))
  expect(readFileSync(manifest, "utf8")).toBe(renderManifest("https://localhost:4711"))
  expect(calls.map(call => call.slice(0, 2))).toEqual([["/usr/bin/codesign", "--verify"], ["security", "verify-cert"]])
  writeFileSync(registered, "unrelated add-in")
  expect(await syncWordAddinRegistration(options)).toBe(false)
  expect(readFileSync(registered, "utf8")).toBe("unrelated add-in")
  writeFileSync(registered, renderManifest("https://localhost:4611"))
  await expect(syncWordAddinRegistration({ ...options, command: async (args) => ({ code: args[0] === "security" ? 1 : 0, stdout: "" }) })).rejects.toThrow("not trusted")
  expect(readFileSync(registered, "utf8")).toBe(renderManifest("https://localhost:4611"))
  // A second configured runtime cannot take the first installation's Word registration.
  const other = path.join(root, "other")
  const otherConfig = path.join(other, "runtime/config"), otherAddin = path.join(otherConfig, "word-addin")
  mkdirSync(otherAddin, { recursive: true })
  writeFileSync(path.join(otherConfig, "word-addin.json"), JSON.stringify({ origin: "https://localhost:4711", external: false, ...tls }))
  writeFileSync(path.join(otherAddin, "manifest.xml"), renderManifest("https://localhost:4711"))
  expect(await syncWordAddinRegistration({ ...options, environment: { DOCHAUS_DATA_ROOT: other } })).toBe(false)
  expect(readFileSync(registered, "utf8")).toBe(renderManifest("https://localhost:4611"))
}, 15000)
