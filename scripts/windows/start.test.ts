import { expect, test } from "bun:test"
import { listeningPids, parseEnvFile } from "./start"

test(".env.local is read with dotenv quoting and comments, and Windows paths stay literal", () => {
  const file = [
    "# provider credentials",
    "GOOGLE_VERTEX_PROJECT=demo-project",
    "export OPENAI_API_KEY = sk-test # rotated monthly",
    "GOOGLE_APPLICATION_CREDENTIALS=C:\\Users\\Ana\\keys\\adc.json",
    "QUOTED=\"C:\\Program Files\\doc #1\"  # trailing",
    "SINGLE='a\\nb'",
    "LINES=\"a\\nb\"",
    "EMPTY=",
    "not a variable",
    "",
  ].join("\r\n")
  expect(parseEnvFile(file)).toEqual({
    GOOGLE_VERTEX_PROJECT: "demo-project",
    OPENAI_API_KEY: "sk-test",
    GOOGLE_APPLICATION_CREDENTIALS: "C:\\Users\\Ana\\keys\\adc.json",
    QUOTED: "C:\\Program Files\\doc #1",
    SINGLE: "a\\nb",
    LINES: "a\nb",
    EMPTY: "",
  })
})

// Shape of `netstat -ano` on Windows 11; the state column is localised.
const netstat = `
Active Connections

  Proto  Local Address          Foreign Address        State           PID
  TCP    0.0.0.0:135            0.0.0.0:0              LISTENING       1032
  TCP    127.0.0.1:4097         0.0.0.0:0              ABHÖREN         5120
  TCP    127.0.0.1:14097        0.0.0.0:0              LISTENING       7000
  TCP    127.0.0.1:52100        127.0.0.1:4501         TIME_WAIT       0
  TCP    [::1]:5174             [::]:0                 LISTENING       6400
  UDP    0.0.0.0:4501           *:*                                    900
`.replaceAll("\n", "\r\n")

test("only TCP listeners on the exact port count as occupied, on IPv4 and IPv6", () => {
  expect(listeningPids(netstat, "4097")).toEqual(["5120"])
  expect(listeningPids(netstat, "5174")).toEqual(["6400"])
  expect(listeningPids(netstat, "4501")).toEqual([])
  expect(listeningPids("", "4097")).toEqual([])
})
