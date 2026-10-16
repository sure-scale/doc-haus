import {
  closeSync,
  constants,
  fchmodSync,
  fstatSync,
  openSync,
  readFileSync,
  realpathSync,
  statSync,
  writeSync,
} from "node:fs"
import path from "node:path"
import { dlopen, FFIType, ptr } from "bun:ffi"

export function readRegularFile(file: string): Buffer | undefined {
  return readRegularFileSnapshot(file)?.content
}

export function readRegularFileSnapshot(
  file: string,
  knownIdentity?: string,
): { identity: string; content?: Buffer } | undefined {
  let handle: number
  try {
    handle = openSync(
      file,
      constants.O_RDONLY |
        constants.O_NONBLOCK |
        constants.O_NOFOLLOW,
    )
  } catch (error) {
    if (
      error instanceof Error &&
      "code" in error &&
      error.code === "ENOENT"
    ) return
    throw new Error(
      `Refusing to read non-regular or unsafe file ${file}.`,
      { cause: error },
    )
  }
  try {
    const stat = fstatSync(handle, { bigint: true })
    if (!stat.isFile())
      throw new Error(
        `Refusing to read non-regular file ${file}.`,
      )
    const identity = `${stat.dev}:${stat.ino}:${stat.size}:${stat.mtimeNs}:${stat.ctimeNs}`
    if (knownIdentity === identity) return { identity }
    return { identity, content: readFileSync(handle) }
  } finally {
    closeSync(handle)
  }
}

// Append (create-if-missing) bytes to a regular file inside the matter via
// descriptor-anchored parent/name I/O. The parent is opened O_DIRECTORY|O_NOFOLLOW
// and verified to still resolve inside the real matter root through the opened
// descriptor (`/proc/self/fd` or `/dev/fd`), with its fstat dev/ino matching the
// visible parent. The terminal name is resolved ONLY inside that anchored parent:
// first a non-creating libc openat (O_WRONLY|O_APPEND|O_NONBLOCK|O_NOFOLLOW)
// appends a pre-existing regular file in place; only if that fails is the name
// created exactly once with O_CREAT|O_EXCL (mode 0600), so a race-created or
// unsafe existing name fails closed. A freshly created descriptor is fchmod'd to
// 0600 before any second open, because Darwin's variadic openat mode argument is
// not reliably carried through FFI and previously produced nondeterministic mode
// bits (e.g. 0040) that then made the reopen return EACCES. FIFOs/specials/
// symlinks at the terminal name are rejected and left untouched. The anchored
// name dev/ino is verified before and after the write via a second non-creating
// openat with the same append flags, and the opened parent is re-validated inside
// the matter before and after the write. Every fd and the dlopen handle are
// closed on every path. Fails closed on platforms without a usable libc openat.
export function appendContainedRegularFile(
  directory: string,
  artifact: string,
  bytes: Uint8Array | string,
): void {
  const buffer = typeof bytes === "string" ? Buffer.from(bytes, "utf8") : Buffer.from(bytes)
  const { parent, basename } = resolveContainedParentAndBasename(directory, artifact)
  const operations = openContainedDirectoryOperations()
  if (!operations)
    throw new Error("Anchored matter appends are unavailable on this platform.")
  let parentHandle: number
  try {
    parentHandle = openSync(
      parent,
      constants.O_RDONLY | constants.O_DIRECTORY | constants.O_NOFOLLOW,
    )
  } catch (error) {
    operations.close()
    throw new Error(
      `Refusing to append to ${artifact}: the parent directory is missing or unsafe.`,
      { cause: error },
    )
  }
  try {
    if (!openedParentContainedWithin(directory, parent, parentHandle))
      throw new Error(
        `Refusing to append to ${artifact}: the parent resolves outside the matter.`,
      )
    const name = Buffer.from(`${basename}\0`)
    // Non-creating open first: a pre-existing regular file is appended in place,
    // and an existing unsafe name (FIFO/symlink) is rejected by fstat below.
    const appendFlags =
      constants.O_WRONLY | constants.O_APPEND | constants.O_NONBLOCK | constants.O_NOFOLLOW
    let writeHandle = operations.symbols.openat(parentHandle, ptr(name), appendFlags, 0)
    let created = false
    if (writeHandle < 0) {
      // Create exactly once with O_EXCL: a race-created or unsafe name fails closed.
      writeHandle = operations.symbols.openat(
        parentHandle,
        ptr(name),
        appendFlags | constants.O_CREAT | constants.O_EXCL,
        0o600,
      )
      if (writeHandle < 0)
        throw new Error(`Refusing to append to non-regular or unsafe file ${artifact}.`)
      created = true
    }
    try {
      // Normalize a freshly created descriptor before any second open; never chmod
      // an existing file (its mode is the caller's own, established on first create).
      if (created) fchmodSync(writeHandle, 0o600)
      const writeStat = fstatSync(writeHandle, { bigint: true })
      if (!writeStat.isFile())
        throw new Error(`Refusing to append to non-regular file ${artifact}.`)
      if (!anchoredNameMatchesHandle(operations, parentHandle, name, appendFlags, writeStat))
        throw new Error(
          `Refusing to append to ${artifact}: the terminal name resolves to a different file.`,
        )
      if (!openedParentContainedWithin(directory, parent, parentHandle))
        throw new Error(
          `Refusing to append to ${artifact}: the parent moved out of the matter before the write.`,
        )
      let offset = 0
      while (offset < buffer.length) {
        const written = writeSync(writeHandle, buffer, offset, buffer.length - offset)
        if (written <= 0) throw new Error(`Failed to append bytes to ${artifact}.`)
        offset += written
      }
      if (!openedParentContainedWithin(directory, parent, parentHandle))
        throw new Error(
          `Refusing to append to ${artifact}: the parent moved out of the matter after the write.`,
        )
      if (!anchoredNameMatchesHandle(operations, parentHandle, name, appendFlags, writeStat))
        throw new Error(
          `Refusing to append to ${artifact}: the terminal name changed during the append.`,
        )
    } finally {
      closeSync(writeHandle)
    }
  } finally {
    closeSync(parentHandle)
    operations.close()
  }
}

function resolveContainedParentAndBasename(
  directory: string,
  artifact: string,
): { parent: string; basename: string } {
  if (path.isAbsolute(artifact))
    throw new Error(`Refusing to append to an absolute artifact path ${artifact}.`)
  const forwardSlashed = artifact.replace(/\\/g, "/")
  if (/(^|\/)\.\.(\/|$)/.test(forwardSlashed) || forwardSlashed.startsWith("/"))
    throw new Error(`Refusing to append to an artifact that escapes the matter ${artifact}.`)
  const target = path.join(directory, artifact)
  const relative = path.relative(directory, target)
  if (relative === "" || relative === ".")
    throw new Error(`Refusing to append to the matter root for artifact ${artifact}.`)
  if (relative.startsWith("..") || path.isAbsolute(relative))
    throw new Error(`Refusing to append outside the matter for artifact ${artifact}.`)
  return { parent: path.dirname(target), basename: path.basename(target) }
}

function openedParentContainedWithin(directory: string, parent: string, handle: number): boolean {
  try {
    const rootReal = realpathSync.native(directory)
    const visibleReal = realpathSync.native(parent)
    if (visibleReal === rootReal) return false
    if (!visibleReal.startsWith(rootReal + path.sep)) return false
    const descriptor = process.platform === "linux" ? `/proc/self/fd/${handle}` : `/dev/fd/${handle}`
    const anchoredReal = realpathSync.native(descriptor)
    if (anchoredReal === rootReal) return false
    if (!anchoredReal.startsWith(rootReal + path.sep)) return false
    const anchoredStat = fstatSync(handle, { bigint: true })
    const visibleStat = statSync(parent, { bigint: true })
    return anchoredStat.dev === visibleStat.dev && anchoredStat.ino === visibleStat.ino
  } catch {
    return false
  }
}

function anchoredNameMatchesHandle(
  operations: NonNullable<ReturnType<typeof openContainedDirectoryOperations>>,
  parentHandle: number,
  name: Buffer,
  flags: number,
  expected: ReturnType<typeof fstatSync>,
): boolean {
  const handle = operations.symbols.openat(parentHandle, ptr(name), flags, 0)
  if (handle < 0) return false
  try {
    const current = fstatSync(handle, { bigint: true })
    return (
      current.isFile() &&
      current.dev === expected.dev &&
      current.ino === expected.ino
    )
  } finally {
    closeSync(handle)
  }
}

function openContainedDirectoryOperations() {
  if (process.platform !== "darwin" && process.platform !== "linux") return
  return dlopen(process.platform === "darwin" ? "/usr/lib/libSystem.B.dylib" : "libc.so.6", {
    openat: {
      args: [FFIType.i32, FFIType.ptr, FFIType.i32, FFIType.i32],
      returns: FFIType.i32,
    },
  })
}
