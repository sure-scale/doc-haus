import { expect, test } from "bun:test"
import { denoProblem, joinWindowsPath, parseRegistryPath } from "./requirements"

// Windows hands names in its own casing; the lookup must not depend on it.
const env = { USERPROFILE: "C:\\Users\\Ana", systemroot: "C:\\Windows" }

test("registry PATH values come from reg query output, and only REG_EXPAND_SZ is expanded", () => {
  const user = "\r\nHKEY_CURRENT_USER\\Environment\r\n    Path    REG_EXPAND_SZ    %USERPROFILE%\\.bun\\bin;C:\\Program Files\\Git\\cmd;%NOT_SET%\\bin\r\n\r\n"
  expect(parseRegistryPath(user, env)).toBe("C:\\Users\\Ana\\.bun\\bin;C:\\Program Files\\Git\\cmd;%NOT_SET%\\bin")
  const machine = "\r\nHKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment\r\n    PATH    REG_EXPAND_SZ    %SystemRoot%\\system32;%SYSTEMROOT%\r\n\r\n"
  expect(parseRegistryPath(machine, env)).toBe("C:\\Windows\\system32;C:\\Windows")
  expect(parseRegistryPath("\r\nHKEY_CURRENT_USER\\Environment\r\n    Path    REG_SZ    %USERPROFILE%\\tools\r\n\r\n", env)).toBe("%USERPROFILE%\\tools")
  expect(parseRegistryPath("\r\nHKEY_CURRENT_USER\\Environment\r\n    Path    REG_SZ    \r\n\r\n", env)).toBe("")
  expect(parseRegistryPath("", env)).toBe("")
})

test("the tool PATH keeps the current PATH first and each folder once", () => {
  expect(joinWindowsPath(
    "C:\\Users\\Ana\\.bun\\bin;C:\\Windows\\system32;;",
    "C:\\WINDOWS\\System32\\;C:\\Program Files\\Git\\cmd",
    "",
    "C:\\Program Files\\LibreOffice\\program",
    "c:\\program files\\git\\cmd\\",
  )).toBe("C:\\Users\\Ana\\.bun\\bin;C:\\Windows\\system32;C:\\Program Files\\Git\\cmd;C:\\Program Files\\LibreOffice\\program")
})

test("Deno older than 2 or with an unreadable version needs an update", () => {
  expect(denoProblem("deno 2.8.2 (stable, release, x86_64-pc-windows-msvc)\r\nv8 13.7.152.6\r\ntypescript 5.8.3\r\n")).toBeUndefined()
  expect(denoProblem("deno 1.46.3 (stable, release, x86_64-pc-windows-msvc)\r\n")).toBe("found 1.46.3, needs 2 or newer")
  expect(denoProblem("")).toBe("found an unknown version, needs 2 or newer")
})
