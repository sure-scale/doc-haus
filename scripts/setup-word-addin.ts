import { X509Certificate, createPrivateKey, createPublicKey } from "node:crypto"
import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { createInterface } from "node:readline/promises"
import { generate } from "selfsigned"
import { renderManifest } from "../apps/word-addin/src/manifest"
import { WORD_ADDIN_ID } from "../dochaus/lib/word-addin-contract"
import { wordConfigurationFile, wordOrigin } from "../services/ingest/src/word-configuration"
import { refreshDesktopLauncher } from "./desktop-launcher"

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

export type WordSetupOptions = {
  environment?: Record<string, string | undefined>
  platform?: NodeJS.Platform
  home?: string
  nativeConfirmation?: boolean
  command?: (command: string[], environment: Record<string, string | undefined>) => Promise<{ code: number; stdout: string }>
}

async function command(args: string[], options: WordSetupOptions = {}, extra: Record<string, string> = {}) {
  const environment = { ...(options.environment ?? process.env), ...extra }
  if (options.command) return options.command(args, environment)
  const child = Bun.spawn(args, { env: environment, stdin: "ignore", stdout: "pipe", stderr: "ignore" })
  const [code, stdout] = await Promise.all([child.exited, new Response(child.stdout).text()])
  return { code, stdout }
}

async function run(args: string[], options: WordSetupOptions, extra?: Record<string, string>) {
  if ((await command(args, options, extra)).code !== 0)
    throw new Error("Word setup was not completed. Allow the system confirmation, then try again.")
}

export async function wordInstalled(options: WordSetupOptions = {}) {
  const platform = options.platform ?? process.platform
  if (platform === "darwin") return ["/Applications/Microsoft Word.app", path.join(options.home ?? os.homedir(), "Applications/Microsoft Word.app")].some(existsSync)
  if (platform !== "win32") return false
  return (await command(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
    "if (Test-Path 'Registry::HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE') { exit 0 }; if (Test-Path 'Registry::HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE') { exit 0 }; exit 1"], options)).code === 0
}

export async function wordCertificateTrusted(certificate: string, hostname: string, options: WordSetupOptions = {}) {
  try {
    const platform = options.platform ?? process.platform
    if (platform === "darwin") return (await command(["security", "verify-cert", "-c", certificate, "-p", "ssl", "-s", hostname.replace(/^\[|\]$/g, "")], options)).code === 0
    if (platform === "win32") return (await command(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
      "if (Test-Path ('Cert:\\CurrentUser\\Root\\' + $env:DOCHAUS_WORD_CERT_THUMBPRINT)) { exit 0 }; if (Test-Path ('Cert:\\LocalMachine\\Root\\' + $env:DOCHAUS_WORD_CERT_THUMBPRINT)) { exit 0 }; exit 1"], options,
      { DOCHAUS_WORD_CERT_THUMBPRINT: new X509Certificate(readFileSync(certificate)).fingerprint.replaceAll(":", "") })).code === 0
  } catch { /* An unreadable certificate or unavailable trust tool is not trusted. */ }
  return false
}

function registrationFile(home: string) {
  return path.join(home, "Library/Containers/com.microsoft.Word/Data/Documents/wef", `doc-haus-${WORD_ADDIN_ID}.xml`)
}

export async function wordAddinRegistered(origin: string, manifest: string, options: WordSetupOptions = {}) {
  try {
    const platform = options.platform ?? process.platform
    let registered = registrationFile(options.home ?? os.homedir())
    if (platform === "win32") {
      const result = await command(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
        "$p = 'HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer'; try { (Get-ItemPropertyValue -Path $p -Name $env:DOCHAUS_WORD_MANIFEST_ID -ErrorAction Stop) } catch { exit 1 }"], options, { DOCHAUS_WORD_MANIFEST_ID: WORD_ADDIN_ID })
      if (result.code !== 0 || path.resolve(result.stdout.trim()) !== path.resolve(manifest)) return false
      registered = manifest
    } else if (platform !== "darwin") return false
    return readFileSync(registered, "utf8") === renderManifest(origin)
  } catch { return false }
}

export async function setupWordAddin(args = process.argv.slice(2), options: WordSetupOptions = {}) {
  const platform = options.platform ?? process.platform
  const environment = options.environment ?? process.env
  const home = options.home ?? os.homedir()
  if (!["darwin", "win32"].includes(platform)) {
    console.log("Open in Word requires desktop Microsoft Word on macOS or Windows. See docs/word-addin.md for firm-hosted installations.")
    return
  }
  if (args.includes("--offer")) {
    if (!process.stdin.isTTY || !await wordInstalled(options)) return
    const terminal = createInterface({ input: process.stdin, output: process.stdout })
    const answer = await terminal.question("Set up Open in Word? This installs the doc.haus add-in and trusts a localhost certificate for this installation. [y/N] ")
    terminal.close()
    if (!/^y(es)?$/i.test(answer.trim())) return
  }
  const file = wordConfigurationFile(environment)
  const directory = path.join(path.dirname(file), "word-addin")
  mkdirSync(directory, { recursive: true, mode: 0o700 })
  let saved: Record<string, unknown> = {}
  try {
    const value: unknown = JSON.parse(readFileSync(file, "utf8"))
    if (value && typeof value === "object" && !Array.isArray(value)) saved = value as Record<string, unknown>
  } catch { /* Explicit setup also repairs incomplete configuration. */ }
  let savedOrigin: string | undefined
  try { if (typeof saved.origin === "string") savedOrigin = wordOrigin(saved.origin) } catch { /* Repair malformed local settings. */ }
  const origin = wordOrigin(environment.DOCHAUS_WORD_ORIGIN || savedOrigin || "https://localhost:4511")
  const external = environment.DOCHAUS_WORD_EXTERNAL_HTTPS === "1" || (!environment.DOCHAUS_WORD_ORIGIN && saved.external === true)
  let settings: Record<string, unknown> = { origin, external }
  if (!external) {
    if (!["localhost", "127.0.0.1", "[::1]"].includes(new URL(origin).hostname))
      throw new Error("For a firm HTTPS address, set DOCHAUS_WORD_EXTERNAL_HTTPS=1. Local certificate setup supports localhost only.")
    if (Boolean(environment.DOCHAUS_WORD_TLS_CERT) !== Boolean(environment.DOCHAUS_WORD_TLS_KEY))
      throw new Error("Provide both DOCHAUS_WORD_TLS_CERT and DOCHAUS_WORD_TLS_KEY, or neither to create a local certificate.")
    const provided = environment.DOCHAUS_WORD_TLS_CERT && environment.DOCHAUS_WORD_TLS_KEY
    const tls = provided ? { certificate: path.resolve(environment.DOCHAUS_WORD_TLS_CERT!), key: path.resolve(environment.DOCHAUS_WORD_TLS_KEY!) } :
      await prepareLocalCertificate(directory)
    if (!existsSync(tls.certificate) || !existsSync(tls.key)) throw new Error("The configured Word TLS files do not exist.")
    validateWordCertificate(tls.certificate, tls.key, new URL(origin).hostname)
    // Only an explicit setup action can change trust. A native confirmation is
    // shown for a browser-initiated request, before the OS trust-store command.
    if (!await wordCertificateTrusted(tls.certificate, new URL(origin).hostname, options)) {
      if (options.nativeConfirmation) {
        if (platform === "darwin") await run(["osascript", "-e",
          'display dialog "Allow doc.haus to trust its local Microsoft Word connection? This adds this installation’s localhost certificate to your account. Only continue if you chose Set up Word in doc.haus." with title "Set up Microsoft Word" buttons {"Cancel", "Allow"} default button "Allow" cancel button "Cancel" with icon caution'], options)
        else await run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
          "Add-Type -AssemblyName System.Windows.Forms; if ([System.Windows.Forms.MessageBox]::Show('Allow doc.haus to trust its local Microsoft Word connection? This adds this installation''s localhost certificate to your account. Only continue if you chose Set up Word in doc.haus.', 'Set up Microsoft Word', 'OKCancel', 'Warning') -ne 'OK') { exit 1 }"], options)
      }
      if (platform === "darwin")
        await run(["security", "add-trusted-cert", "-r", "trustRoot", "-k", path.join(home, "Library/Keychains/login.keychain-db"), tls.certificate], options)
      else await run(["certutil.exe", "-user", "-addstore", "Root", tls.certificate], options)
      if (!await wordCertificateTrusted(tls.certificate, new URL(origin).hostname, options))
        throw new Error("The local Word certificate is not trusted yet. Complete the system confirmation, then try again.")
    }
    settings = { ...settings, ...tls }
  }
  const manifest = path.join(directory, "manifest.xml")
  const registrationChanged = !await wordAddinRegistered(origin, manifest, options)
  writeFileSync(manifest, renderManifest(origin), { mode: 0o600 })
  if (registrationChanged && platform === "darwin") {
    const registered = registrationFile(home)
    mkdirSync(path.dirname(registered), { recursive: true })
    writeFileSync(registered, renderManifest(origin))
  } else if (registrationChanged) {
    await run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
      "$ErrorActionPreference = 'Stop'; $p = 'HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer'; New-Item -Path $p -Force | Out-Null; New-ItemProperty -Path $p -Name $env:DOCHAUS_WORD_MANIFEST_ID -Value $env:DOCHAUS_WORD_MANIFEST_FILE -PropertyType String -Force | Out-Null"],
      options, { DOCHAUS_WORD_MANIFEST_ID: WORD_ADDIN_ID, DOCHAUS_WORD_MANIFEST_FILE: manifest })
  }
  const temporary = `${file}.${crypto.randomUUID()}.tmp`
  try {
    writeFileSync(temporary, JSON.stringify(settings, null, 2) + "\n", { mode: 0o600, flag: "wx" })
    renameSync(temporary, file)
  } finally { rmSync(temporary, { force: true }) }
  await refreshDesktopLauncher(options)
  return { origin, registrationChanged }
}

if (import.meta.main) {
  try {
    const result = await setupWordAddin()
    if (result) console.log(`Open in Word is configured for ${result.origin}. Restart doc.haus${result.registrationChanged ? " and Word" : ""}, then use Open in Word on a document.`)
  }
  catch (error) { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1 }
}
