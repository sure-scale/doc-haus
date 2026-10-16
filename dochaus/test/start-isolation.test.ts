import { describe, expect, test } from "bun:test"

const startScript = () => Bun.file("../start.sh").text()
const engineEntry = () => Bun.file("../services/engine/src/server.ts").text()

describe("start.sh isolated OpenCode runtime", () => {
  test("always isolates the web engine from host OpenCode context", async () => {
    const script = await startScript()

    expect(script).toContain("XDG_CONFIG_HOME=\"$DOCHAUS_RUNTIME_DIR/config\"")
    expect(script).toContain("XDG_STATE_HOME=\"$DOCHAUS_RUNTIME_DIR/state\"")
    expect(script).toContain("XDG_CACHE_HOME=\"$DOCHAUS_RUNTIME_DIR/cache\"")
    expect(script).toContain("export XDG_DATA_HOME=\"$DOCHAUS_RUNTIME_DIR/data\"")
    expect(script).toContain("unset OPENCODE_CONFIG OPENCODE_CONFIG_CONTENT OPENCODE_PERMISSION")
    expect(script).not.toContain("DOCHAUS_ISOLATED")

    // v2 does not honour the old OPENCODE_DISABLE_* or sealed-config env flags.
    // The engine composes explicit discovery policy in-process instead, so the
    // boundary is asserted where it is enforced.
    const engine = await engineEntry()
    expect(engine).toContain("process.env.OPENCODE_CONFIG_DIR = legal")
    expect(engine).toContain("layer: legalConfigLayer(legal)")
    expect(engine).toContain("config: { directory: path.join(runtime, \"config\"), project: false }")
    expect(engine).toContain("Config.node.replace(legalConfig)")
    expect(engine).toContain("InstructionDiscovery.node.replace(InstructionDiscovery.configured({ global: false, project: false }))")
    expect(engine).toContain("Remote configuration discovery is disabled in the legal engine")
  })

  test("leaves provider retry timing and attempts to OpenCode", async () => {
    const script = await startScript()

    expect(script).not.toMatch(/^OPENCODE_PROVIDER_MAX_ATTEMPTS=/m)
    expect(script).not.toContain("export OPENCODE_PROVIDER_MAX_ATTEMPTS")
  })
})
