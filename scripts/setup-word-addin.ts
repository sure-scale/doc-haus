import { X509Certificate, createPrivateKey, createPublicKey } from "node:crypto"
import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { createInterface } from "node:readline/promises"
import { generate } from "selfsigned"
import { renderManifest } from "../apps/word-addin/src/manifest"
import { WORD_ADDIN_ID } from "../dochaus/lib/word-addin-contract"
import { wordConfigurationFile, wordOrigin } from "../services/ingest/src/word-configuration"

export function validateWordCertificate(certificate: string, key: string, hostname = "localhost") {
  const cert = new X509Certificate(readFileSync(certificate))
  const publicKey = createPublicKey(createPrivateKey(readFileSync(key))).export({ type: "spki", format: "pem" })
  const host = hostname.replace(/^\[|\]$/g, "")
  const matches = host.includes(":") || /^\d+\.\d+\.\d+\.\d+$/.test(host) ? cert.checkIP(host) : cert.checkHost(host)
  if (cert.ca || !matches || cert.publicKey.export({ type: "spki", format: "pem" }) !== publicKey ||
      Date.parse(cert.validFrom) > Date.now() || Date.parse(cert.validTo) <= Date.now() + 7 * 86400_000)
    throw new Error("The Word TLS certificate must be an unexpired server leaf for the configured host, with its matching private key. Renew it and run setup again.")
  return cert
}

export async function prepareLocalCertificate(directory: string) {
  mkdirSync(directory, { recursive: true, mode: 0o700 })
  const certificate = path.join(directory, "localhost.crt")
  const key = path.join(directory, "localhost.key")
  if (existsSync(certificate) && existsSync(key)) {
    try {
      const cert = validateWordCertificate(certificate, key)
      if (cert.checkIP("127.0.0.1") && cert.checkIP("::1"))
      return { certificate, key }
    } catch { /* A partial or expired local pair is renewed by explicit setup. */ }
  }
  const pem = await generate([{ name: "commonName", value: "localhost" }], {
    algorithm: "sha256", keyType: "rsa", keySize: 2048,
    notBeforeDate: new Date(Date.now() - 5 * 60_000), notAfterDate: new Date(Date.now() + 365 * 86400_000),
    extensions: [
      { name: "basicConstraints", cA: false, critical: true },
      { name: "keyUsage", digitalSignature: true, keyEncipherment: true, critical: true },
      { name: "extKeyUsage", serverAuth: true },
      { name: "subjectAltName", altNames: [{ type: 2, value: "localhost" }, { type: 7, ip: "127.0.0.1" }, { type: 7, ip: "::1" }] },
    ],
  })
  writeFileSync(key, pem.private, { mode: 0o600 })
  chmodSync(key, 0o600)
  writeFileSync(certificate, pem.cert, { mode: 0o600 })
  return { certificate, key }
}

async function run(command: string[], env = process.env) {
  const child = Bun.spawn(command, { env, stdin: "inherit", stdout: "inherit", stderr: "inherit" })
  if (await child.exited !== 0) throw new Error("Word add-in setup did not complete. The operating system command above failed.")
}

async function wordInstalled() {
  if (process.platform === "darwin") return ["/Applications/Microsoft Word.app", path.join(os.homedir(), "Applications/Microsoft Word.app")].some(existsSync)
  if (process.platform !== "win32") return false
  const child = Bun.spawn(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
    "if (Test-Path 'Registry::HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE') { exit 0 }; if (Test-Path 'Registry::HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE') { exit 0 }; exit 1"], { stdout: "ignore", stderr: "ignore" })
  return await child.exited === 0
}

export async function setupWordAddin(args = process.argv.slice(2)) {
  if (!["darwin", "win32"].includes(process.platform)) {
    console.log("Open in Word requires desktop Microsoft Word on macOS or Windows. See docs/word-addin.md for firm-hosted installations.")
    return
  }
  if (args.includes("--offer")) {
    if (!process.stdin.isTTY || !await wordInstalled()) return
    const terminal = createInterface({ input: process.stdin, output: process.stdout })
    const answer = await terminal.question("Set up Open in Word? This installs the doc.haus add-in and trusts a localhost certificate for this installation. [y/N] ")
    terminal.close()
    if (!/^y(es)?$/i.test(answer.trim())) return
  }
  const file = wordConfigurationFile()
  const directory = path.join(path.dirname(file), "word-addin")
  mkdirSync(directory, { recursive: true, mode: 0o700 })
  const saved = existsSync(file) ? JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown> : {}
  const origin = wordOrigin(process.env.DOCHAUS_WORD_ORIGIN || (typeof saved.origin === "string" ? saved.origin : "https://localhost:4511"))
  const external = process.env.DOCHAUS_WORD_EXTERNAL_HTTPS === "1" || (!process.env.DOCHAUS_WORD_ORIGIN && saved.external === true)
  let settings: Record<string, unknown> = { origin, external }
  if (!external) {
    if (!["localhost", "127.0.0.1", "[::1]"].includes(new URL(origin).hostname))
      throw new Error("For a firm HTTPS address, set DOCHAUS_WORD_EXTERNAL_HTTPS=1. Local certificate setup supports localhost only.")
    if (Boolean(process.env.DOCHAUS_WORD_TLS_CERT) !== Boolean(process.env.DOCHAUS_WORD_TLS_KEY))
      throw new Error("Provide both DOCHAUS_WORD_TLS_CERT and DOCHAUS_WORD_TLS_KEY, or neither to create a local certificate.")
    const provided = process.env.DOCHAUS_WORD_TLS_CERT && process.env.DOCHAUS_WORD_TLS_KEY
    const tls = provided ? { certificate: path.resolve(process.env.DOCHAUS_WORD_TLS_CERT!), key: path.resolve(process.env.DOCHAUS_WORD_TLS_KEY!) } :
      await prepareLocalCertificate(directory)
    if (!existsSync(tls.certificate) || !existsSync(tls.key)) throw new Error("The configured Word TLS files do not exist.")
    validateWordCertificate(tls.certificate, tls.key, new URL(origin).hostname)
    // Explicit setup is the only path that changes OS trust. The server never does.
    console.log("Trusting the installation's localhost certificate in your user account. The operating system may ask you to confirm.")
    if (process.platform === "darwin")
      await run(["security", "add-trusted-cert", "-r", "trustRoot", "-k", path.join(os.homedir(), "Library/Keychains/login.keychain-db"), tls.certificate])
    else await run(["certutil.exe", "-user", "-addstore", "Root", tls.certificate])
    settings = { ...settings, ...tls }
  }
  const manifest = path.join(directory, "manifest.xml")
  writeFileSync(manifest, renderManifest(origin), { mode: 0o600 })
  if (process.platform === "darwin") {
    const wef = path.join(os.homedir(), "Library/Containers/com.microsoft.Word/Data/Documents/wef")
    mkdirSync(wef, { recursive: true })
    writeFileSync(path.join(wef, `doc-haus-${WORD_ADDIN_ID}.xml`), renderManifest(origin))
  } else {
    await run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
      "$ErrorActionPreference = 'Stop'; $p = 'HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer'; New-Item -Path $p -Force | Out-Null; New-ItemProperty -Path $p -Name $env:DOCHAUS_WORD_MANIFEST_ID -Value $env:DOCHAUS_WORD_MANIFEST_FILE -PropertyType String -Force | Out-Null"],
      { ...process.env, DOCHAUS_WORD_MANIFEST_ID: WORD_ADDIN_ID, DOCHAUS_WORD_MANIFEST_FILE: manifest })
  }
  writeFileSync(file, JSON.stringify(settings, null, 2) + "\n", { mode: 0o600 })
  console.log(`Open in Word is configured for ${origin}. Restart doc.haus and Word, then use Open in Word on a document.`)
}

if (import.meta.main) {
  try { await setupWordAddin() }
  catch (error) { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1 }
}
