import { expect, test } from "bun:test"
import {
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  canonicalArtifactIdentity,
  withArtifactMutationLock,
  withArtifactMutationLockAsync,
} from "./mutation-lock"

test("canonicalizes an absent artifact through its nearest existing ancestor", () => {
  const root = mkdtempSync(path.join(tmpdir(), "dochaus-mutation-lock-"))
  try {
    const real = path.join(root, "real")
    const alias = path.join(root, "alias")
    const matter = path.join(alias, "matter")
    mkdirSync(path.join(real, "matter"), { recursive: true })
    symlinkSync(real, alias, "dir")

    const before = canonicalArtifactIdentity(matter, ".dochaus/drafting-run-request.json")
    mkdirSync(path.join(matter, ".dochaus"), { recursive: true })

    expect(canonicalArtifactIdentity(matter, ".dochaus/drafting-run-request.json")).toBe(before)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

test.each(["control-root", "lock-root"] as const)(
  "rejects an outside package %s before creating a lock or mutating its target",
  (mode) => {
    const directory = mkdtempSync(
      path.join(
        tmpdir(),
        `dochaus-mutation-lock-${mode}-`,
      ),
    )
    const outside = mkdtempSync(
      path.join(
        tmpdir(),
        `dochaus-mutation-lock-${mode}-outside-`,
      ),
    )
    try {
      if (mode === "control-root")
        symlinkSync(
          outside,
          path.join(directory, ".dochaus"),
        )
      else {
        mkdirSync(
          path.join(directory, ".dochaus"),
        )
        symlinkSync(
          outside,
          path.join(
            directory,
            ".dochaus",
            "artifact-mutation-locks",
          ),
        )
      }
      const target = path.join(
        directory,
        "instrument.md",
      )
      const prior = Buffer.from("prior bytes\n")
      writeFileSync(target, prior)

      expect(() =>
        withArtifactMutationLock(
          directory,
          "instrument.md",
          () => {
            writeFileSync(target, "foreign bytes\n")
          },
        )
      ).toThrow(/outside|unsafe|matter/i)
      expect(readdirSync(outside)).toEqual([])
      expect(readFileSync(target)).toEqual(prior)
    } finally {
      rmSync(directory, {
        recursive: true,
        force: true,
      })
      rmSync(outside, {
        recursive: true,
        force: true,
      })
    }
  },
)

test.each(["traversal", "absolute", "symlink"] as const)(
  "rejects an outside ordinary synchronous %s target before its callback executes",
  (mode) => {
    const directory = mkdtempSync(
      path.join(tmpdir(), `dochaus-mutation-lock-target-${mode}-`),
    )
    const outside = mkdtempSync(
      path.join(tmpdir(), `dochaus-mutation-lock-target-${mode}-outside-`),
    )
    try {
      const outsideTarget = path.join(outside, "instrument.md")
      const prior = Buffer.from("exact outside bytes\n")
      writeFileSync(outsideTarget, prior)
      const artifact = mode === "traversal"
        ? path.relative(directory, outsideTarget)
        : mode === "absolute"
          ? outsideTarget
          : "instrument.md"
      if (mode === "symlink")
        symlinkSync(outsideTarget, path.join(directory, artifact))
      const outsideState = readdirSync(outside)
      let executed = false

      expect(() =>
        withArtifactMutationLock(directory, artifact, () => {
          executed = true
          writeFileSync(outsideTarget, "mutated outside bytes\n")
        })
      ).toThrow(/outside|unsafe|matter/i)
      expect(executed).toBe(false)
      expect(readFileSync(outsideTarget)).toEqual(prior)
      expect(readdirSync(outside)).toEqual(outsideState)
    } finally {
      rmSync(directory, { recursive: true, force: true })
      rmSync(outside, { recursive: true, force: true })
    }
  },
)

test.each(["traversal", "absolute", "symlink"] as const)(
  "rejects an outside ordinary asynchronous %s target before its callback executes",
  async (mode) => {
    const directory = mkdtempSync(
      path.join(tmpdir(), `dochaus-mutation-lock-target-async-${mode}-`),
    )
    const outside = mkdtempSync(
      path.join(tmpdir(), `dochaus-mutation-lock-target-async-${mode}-outside-`),
    )
    try {
      const outsideTarget = path.join(outside, "instrument.md")
      const prior = Buffer.from("exact outside bytes\n")
      writeFileSync(outsideTarget, prior)
      const artifact = mode === "traversal"
        ? path.relative(directory, outsideTarget)
        : mode === "absolute"
          ? outsideTarget
          : "instrument.md"
      if (mode === "symlink")
        symlinkSync(outsideTarget, path.join(directory, artifact))
      const outsideState = readdirSync(outside)
      let executed = false

      await expect(
        withArtifactMutationLockAsync(directory, artifact, async () => {
          executed = true
          await Bun.write(outsideTarget, "mutated outside bytes\n")
        }),
      ).rejects.toThrow(/outside|unsafe|matter/i)
      expect(executed).toBe(false)
      expect(readFileSync(outsideTarget)).toEqual(prior)
      expect(readdirSync(outside)).toEqual(outsideState)
    } finally {
      rmSync(directory, { recursive: true, force: true })
      rmSync(outside, { recursive: true, force: true })
    }
  },
)
