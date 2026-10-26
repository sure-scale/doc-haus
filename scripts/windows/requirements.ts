import { existsSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { createInterface } from "node:readline"

// Windows counterpart of requirements.sh: checks the optional host tools doc.haus
// calls and offers to install whatever is missing or too old with winget. Every
// install is asked first and Enter means yes. Without a terminal nothing is
// installed unless DOCHAUS_INSTALL_TOOLS=yes; DOCHAUS_INSTALL_TOOLS=no never
// installs and DOCHAUS_INSTALL_TOOLS=ask asks again about tools declined before.

const TOOLS: { id: string; label: string; commands: string[]; winget?: string }[] = [
  { id: "deno", label: "Deno 2+ (runs the shipped Python tool)", commands: ["deno"], winget: "DenoLand.Deno" },
  // Ingest runs markitdown through uvx when it is not installed directly, and uv
  // fetches its own Python, so Windows needs no separate Python or pipx.
  { id: "markitdown", label: "markitdown via uv (PDF text fallback)", commands: ["markitdown", "uvx"], winget: "astral-sh.uv" },
  { id: "poppler", label: "Poppler (renders PDF pages for OCR)", commands: ["pdftoppm"], winget: "oschwartz10612.Poppler" },
  { id: "tesseract", label: "Tesseract (OCR for scanned PDFs)", commands: ["tesseract"], winget: "tesseract-ocr.tesseract" },
  { id: "libreoffice", label: "LibreOffice Writer (PDF to DOCX conversion)", commands: ["soffice"], winget: "TheDocumentFoundation.LibreOffice" },
  // Windows ships Calibri and Carlito is its metric-compatible stand-in; winget has no package for either.
  { id: "fonts", label: "Calibri or Carlito font (DOCX rendering)", commands: [] },
  { id: "git", label: "Git (doc.haus updates)", commands: ["git"], winget: "Git.Git" },
]

// Windows environment names are case-insensitive, but a plain env object is not.
function envValue(env: NodeJS.ProcessEnv, name: string) {
  const key = Object.keys(env).find(key => key.toLowerCase() === name.toLowerCase())
  return key === undefined ? undefined : env[key]
}

// Expand %NAME% like Windows does; unknown names stay as written.
function expandVariables(value: string, env: NodeJS.ProcessEnv) {
  return value.replace(/%([^%]+)%/g, (whole, name: string) => envValue(env, name) ?? whole)
}

// `reg query <key> /v Path` prints the value as "    Path    REG_EXPAND_SZ    C:\a;C:\b".
// Windows expands only REG_EXPAND_SZ values.
export function parseRegistryPath(output: string, env: NodeJS.ProcessEnv): string {
  const match = /^[ \t]+Path[ \t]+(REG_SZ|REG_EXPAND_SZ)(?:[ \t]+(.*))?$/im.exec(output)
  const value = match?.[2]?.trim() ?? ""
  return match?.[1].toUpperCase() === "REG_EXPAND_SZ" ? expandVariables(value, env) : value
}

function registryPath(key: string, env: NodeJS.ProcessEnv) {
  const result = Bun.spawnSync(["reg", "query", key, "/v", "Path"], { stdout: "pipe", stderr: "ignore" })
  return result.exitCode === 0 ? parseRegistryPath(result.stdout.toString(), env) : ""
}

// Join PATH lists in order, keeping the first copy of each folder (compared
// case-insensitively and without a trailing backslash).
export function joinWindowsPath(...lists: string[]): string {
  const seen = new Set<string>(), folders: string[] = []
  for (const folder of lists.flatMap(list => list.split(";")).map(folder => folder.trim())) {
    const key = folder.replace(/[\\/]+$/, "").toLowerCase()
    if (!key || seen.has(key)) continue
    seen.add(key)
    folders.push(folder)
  }
  return folders.join(";")
}

// winget does not update the PATH of a shell opened before it ran, so re-read the
// registry and append the folders installers use. The current PATH stays first.
export function windowsToolPath(env: NodeJS.ProcessEnv = process.env): string {
  const local = envValue(env, "LOCALAPPDATA") ?? "", home = envValue(env, "USERPROFILE") ?? "", programs = envValue(env, "ProgramFiles") ?? ""
  // winget unpacks Poppler into a versioned folder and adds its bin to the user PATH.
  const packages = path.join(local, "Microsoft", "WinGet", "Packages")
  const poppler = existsSync(packages)
    ? [...new Bun.Glob("oschwartz10612.Poppler_*/poppler-*/Library/bin").scanSync({ cwd: packages, onlyFiles: false })].map(folder => path.join(packages, folder))
    : []
  const known = [
    path.join(programs, "LibreOffice", "program"),
    path.join(programs, "Tesseract-OCR"),
    ...poppler,
    path.join(local, "Microsoft", "WinGet", "Links"),
    path.join(envValue(env, "DENO_INSTALL") ?? path.join(home, ".deno"), "bin"),
    path.join(home, ".local", "bin"),
    path.join(programs, "Git", "cmd"),
    path.join(local, "Programs", "Git", "cmd"),
  ].filter(folder => existsSync(folder))
  return joinWindowsPath(
    envValue(env, "PATH") ?? "",
    registryPath("HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment", env),
    registryPath("HKCU\\Environment", env),
    ...known,
  )
}

// Why the `deno --version` output needs an update, or undefined for Deno 2+.
export function denoProblem(output: string): string | undefined {
  const version = /^deno (\d+(?:\.\d+)*)/m.exec(output)?.[1]
  return version && Number(version.split(".")[0]) >= 2 ? undefined : `found ${version ?? "an unknown version"}, needs 2 or newer`
}

function fontInstalled(env: NodeJS.ProcessEnv) {
  const folders = [path.join(envValue(env, "WINDIR") ?? "C:\\Windows", "Fonts"), path.join(envValue(env, "LOCALAPPDATA") ?? "", "Microsoft", "Windows", "Fonts")]
  return folders.some(folder => existsSync(folder) && readdirSync(folder).some(name => /^(calibri|carlito)/i.test(name)))
}

// Map each tool that is missing or too old to why it needs installing.
function toolProblems(toolPath: string) {
  const problems = new Map<string, string>()
  for (const tool of TOOLS) {
    if (tool.id === "fonts") {
      if (!fontInstalled(process.env)) problems.set(tool.id, "not found")
      continue
    }
    const found = tool.commands.map(command => Bun.which(command, { PATH: toolPath })).find(Boolean)
    if (!found) problems.set(tool.id, "not found")
    else if (tool.id === "deno") {
      let output = ""
      try {
        output = Bun.spawnSync([found, "--version"], { stdout: "pipe", stderr: "ignore" }).stdout.toString()
      } catch {}
      const problem = denoProblem(output)
      if (problem) problems.set(tool.id, problem)
    }
  }
  return problems
}

// Ask a yes/no question that defaults to yes; closed input means no.
function confirmInstall(question: string) {
  const prompt = createInterface({ input: process.stdin, output: process.stderr, terminal: false })
  return new Promise<boolean>(resolve => {
    prompt.once("close", () => resolve(false))
    prompt.question(`${question} [Y/n] `, answer => {
      resolve(/^(y|yes)?$/i.test(answer.trim()))
      prompt.close()
    })
  })
}

// Print each install command before running it, so the user sees what changes.
async function runInstall(command: string[]) {
  console.error(`+ ${[path.parse(command[0]).name, ...command.slice(1)].join(" ")}`)
  try {
    return await Bun.spawn(command, { stdin: "inherit", stdout: "inherit", stderr: "inherit" }).exited
  } catch {
    return 1
  }
}

// Offer to install missing optional tools. A declined set is remembered in
// runtimeDir/declined-tools and not asked about again until another tool goes
// missing. Failures are reported, never fatal: the stack starts without them.
export async function ensureWindowsHostTools(runtimeDir: string, options: { interactive?: boolean } = {}): Promise<{ missing: string[] }> {
  const declinedFile = path.join(runtimeDir, "declined-tools"), choice = process.env.DOCHAUS_INSTALL_TOOLS ?? ""
  const toolPath = windowsToolPath()
  const problems = toolProblems(toolPath), missing = [...problems.keys()]
  if (!missing.length) {
    rmSync(declinedFile, { force: true })
    return { missing }
  }
  const declined = existsSync(declinedFile) ? readFileSync(declinedFile, "utf8").split(/\s+/) : []
  if (missing.every(id => declined.includes(id)) && !choice) {
    console.error(`doc.haus: optional tools not installed: ${missing.join(" ")} (set DOCHAUS_INSTALL_TOOLS=ask to be asked again)`)
    return { missing }
  }
  console.error("doc.haus: some optional tools are missing or out of date:")
  for (const tool of TOOLS) if (problems.has(tool.id)) console.error(`  - ${tool.label}: ${problems.get(tool.id)}`)
  const winget = Bun.which("winget", { PATH: toolPath })
  if (!winget || !TOOLS.some(tool => tool.winget && problems.has(tool.id))) {
    console.error(`doc.haus: ${winget ? "winget has no package for these" : "winget (App Installer) was not found"}; install them manually. The stack starts without them.`)
    return { missing }
  }
  if (choice !== "yes" && (choice === "no" || !(options.interactive ?? process.stdin.isTTY))) {
    console.error("doc.haus: starting without them. Run start.cmd in a terminal to install them, or set DOCHAUS_INSTALL_TOOLS=yes.")
    return { missing }
  }
  if (choice !== "yes" && !(await confirmInstall("doc.haus: install them with winget now?"))) {
    writeFileSync(declinedFile, `${missing.join(" ")}\n`)
    console.error("doc.haus: starting without them; you will not be asked again for these tools (set DOCHAUS_INSTALL_TOOLS=ask to change that).")
    return { missing }
  }
  rmSync(declinedFile, { force: true })
  for (const tool of TOOLS) {
    if (!problems.has(tool.id)) continue
    // An outdated Deno updates itself; everything else installs with winget.
    const deno = tool.id === "deno" ? Bun.which("deno", { PATH: toolPath }) : null
    const command = deno
      ? [deno, "upgrade"]
      : tool.winget ? [winget, "install", "--id", tool.winget, "--exact", "--source", "winget", "--accept-package-agreements", "--accept-source-agreements"] : undefined
    if (command && (await runInstall(command)) !== 0) console.error(`doc.haus: ${tool.id} did not ${deno ? "update" : "install"}.`)
  }
  const still = [...toolProblems(windowsToolPath()).keys()]
  if (still.length) console.error(`doc.haus: still missing: ${still.join(" ")}; the stack starts without them.`)
  return { missing: still }
}
