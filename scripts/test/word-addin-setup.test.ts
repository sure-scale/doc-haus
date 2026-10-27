import { afterEach, expect, test } from "bun:test"
import { X509Certificate, createPrivateKey, createPublicKey } from "node:crypto"
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { prepareLocalCertificate, validateWordCertificate } from "../setup-word-addin"

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
