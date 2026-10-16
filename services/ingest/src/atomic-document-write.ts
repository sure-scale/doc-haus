import { randomUUID } from "node:crypto"
import { existsSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"

type SynchronousResult = string | number | boolean | bigint | symbol | null | void | { then?: never }

// Replace document bytes and their registry mutation as one compensated unit.
// The registry callback should use a database transaction. If it throws, the
// prior bytes are restored (or a newly created file is removed) before the
// error reaches the web route.
export function commitDocumentWrite<T extends SynchronousResult>(docPath: string, buffer: Buffer, updateRegistry: () => T) {
  const temporary = path.join(path.dirname(docPath), `.${path.basename(docPath)}.${process.pid}.${randomUUID()}.tmp`)
  const backup = path.join(path.dirname(docPath), `.${path.basename(docPath)}.${process.pid}.${randomUUID()}.bak`)
  const existed = existsSync(docPath)
  writeFileSync(temporary, buffer, { flag: "wx" })
  try {
    if (existed) renameSync(docPath, backup)
    try {
      renameSync(temporary, docPath)
    } catch (error) {
      if (existed) renameSync(backup, docPath)
      throw error
    }
    const result = (() => {
      try {
        const value = updateRegistry()
        if (value && typeof value === "object" && "then" in value && typeof value.then === "function")
          throw new Error("commitDocumentWrite requires a synchronous registry callback.")
        return value
      } catch (error) {
        rmSync(docPath, { force: true })
        if (existed) renameSync(backup, docPath)
        throw error
      }
    })()
    if (existed) removeSettledBackup(backup)
    return result
  } finally {
    rmSync(temporary, { force: true })
  }
}

// Source ingestion also publishes a derived source-corpus snapshot before its
// registry transaction can settle. Keep the prior bytes available until that
// asynchronous publication either commits the whole unit or fails, then apply
// the same compensation as the synchronous deliverable path.
export async function commitDocumentWriteAsync<T>(
  docPath: string,
  buffer: Buffer,
  updateRegistry: () => Promise<T>,
  options: { backupPath?: string } = {},
) {
  const temporary = path.join(path.dirname(docPath), `.${path.basename(docPath)}.${process.pid}.${randomUUID()}.tmp`)
  const backup = options.backupPath ?? path.join(path.dirname(docPath), `.${path.basename(docPath)}.${process.pid}.${randomUUID()}.bak`)
  const existed = existsSync(docPath)
  writeFileSync(temporary, buffer, { flag: "wx" })
  try {
    if (existed) renameSync(docPath, backup)
    try {
      renameSync(temporary, docPath)
    } catch (error) {
      if (existed) renameSync(backup, docPath)
      throw error
    }
    const result = await (async () => {
      try {
        return await updateRegistry()
      } catch (error) {
        rmSync(docPath, { force: true })
        if (existed) renameSync(backup, docPath)
        throw error
      }
    })()
    if (existed) removeSettledBackup(backup)
    return result
  } finally {
    rmSync(temporary, { force: true })
  }
}

// Delete document bytes only after the asynchronous registry/projection unit
// settles. Until then the prior file is a private backup which is restored on
// any failure.
export async function commitDocumentDeleteAsync<T>(
  docPath: string,
  updateRegistry: () => Promise<T>,
  options: { backupPath?: string } = {},
) {
  const backup = options.backupPath ?? path.join(path.dirname(docPath), `.${path.basename(docPath)}.${process.pid}.${randomUUID()}.bak`)
  const existed = existsSync(docPath)
  if (existed) renameSync(docPath, backup)
  const result = await (async () => {
    try {
      return await updateRegistry()
    } catch (error) {
      if (existed) renameSync(backup, docPath)
      throw error
    }
  })()
  if (existed) removeSettledBackup(backup)
  return result
}

function removeSettledBackup(backup: string) {
  try {
    rmSync(backup, { force: true })
  } catch {
    // The current document and registry have already settled. Retaining a
    // private backup is safer than rolling back only the filesystem afterward.
  }
}
