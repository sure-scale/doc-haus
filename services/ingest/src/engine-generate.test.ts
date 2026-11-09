import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdtempSync, rmSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { DOCUMENT_AMENDMENT_LEGAL_OUTPUT } from "../../../dochaus/lib/document-amendment-legal-contract"
import { generateFor } from "./engine-generate"

const original = globalThis.fetch
const cleanups: string[] = []
afterEach(() => { globalThis.fetch = original; for (const directory of cleanups.splice(0)) rmSync(directory,{recursive:true,force:true}) })

test("a lost structured response retains unknown billing instead of becoming an unrecorded free request", async () => {
  const directory = mkdtempSync(path.join(os.tmpdir(),"structured-generation-")); cleanups.push(directory)
  globalThis.fetch = Object.assign(async () => { throw new Error("Synthetic lost local response") },{preconnect:original.preconnect}) as typeof fetch
  await expect(generateFor("prepare-changes",directory,"Synthetic legal job",undefined,DOCUMENT_AMENDMENT_LEGAL_OUTPUT)).rejects.toThrow("Synthetic lost local response")
  const db = new Database(path.join(directory,".dochaus","telemetry.db"),{readonly:true})
  try { expect(db.query("SELECT count(*) AS rows, cost_usd FROM generations").get()).toEqual({rows:1,cost_usd:null}) }
  finally {db.close()}
})
