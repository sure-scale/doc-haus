import { expect, test } from "bun:test"
import path from "node:path"
import { dataPaths } from "../../engine/src/data-paths"

const repository = path.resolve(import.meta.dir, "../../..")
const helper = path.join(repository, "scripts/data-paths.sh")
test("launcher and direct services agree on data-root layout and explicit overrides", async () => {
  for (const environment of [
    {},
    { DOCHAUS_DATA_ROOT: "/tmp/firm data" },
    { DOCHAUS_DATA_ROOT: "data" },
    { DOCHAUS_DATA_ROOT: "data", WORKSPACE_ROOT: "/tmp/existing-matters", DOCHAUS_RUNTIME_DIR: "custom-runtime" },
  ]) {
    const expected = dataPaths(repository, environment)
    const child = Bun.spawn(
      [
        "bash",
        "-c",
        'source "$1"; configure_data_paths "$2"; printf "%s\\n%s\\n" "$WORKSPACE_ROOT" "$DOCHAUS_RUNTIME_DIR"',
        "paths",
        helper,
        repository,
      ],
      {
        env: { PATH: process.env.PATH, ...environment },
        stdout: "pipe",
        stderr: "pipe",
      },
    )
    expect(await child.exited).toBe(0)
    expect(
      (await new Response(child.stdout).text())
        .trim()
        .split("\n")
        .map((value) => path.resolve(value)),
    ).toEqual([expected.workspace, expected.runtime])
  }
})

test("unconfigured installations retain existing paths and standalone ingest fallback", () => {
  expect(dataPaths(repository, {})).toEqual({
    workspace: path.join(repository, "workspace"),
    runtime: path.join(repository, ".dochaus-runtime"),
  })
  expect(dataPaths(repository, {}, "/tmp/old-ingest/workspace").workspace).toBe("/tmp/old-ingest/workspace")
})
