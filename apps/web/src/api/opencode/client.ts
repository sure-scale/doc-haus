import { OpenCode } from "@opencode/client"
import { OPENCODE_URL } from "../../config"

export const clientDirectories = new WeakMap<object, string>()

export function location(client: Client) {
  const directory = clientDirectories.get(client)
  return directory ? { directory } : undefined
}

// One OpenCode client per matter. The matter's directory is sent as the
// x-opencode-directory header on every request, so every session and every
// tool (including search-document) is scoped to that matter's files + DB.
export function matterClient(directory: string) {
  const client = OpenCode.make({ baseUrl: OPENCODE_URL, headers: { "x-opencode-directory": directory } })
  clientDirectories.set(client, directory)
  return client
}

export type Client = ReturnType<typeof matterClient>

// Settings use the sealed runtime configuration and native credential store.
// Matter-independent catalog and integration calls omit a directory.
export function settingsClient() {
  return OpenCode.make({ baseUrl: OPENCODE_URL })
}

export function boundedRequest<T>(promise: Promise<T>, milliseconds: number, label: string) {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${label} timed out.`)), milliseconds)
    promise.then(
      (value) => {
        clearTimeout(timer)
        resolve(value)
      },
      (error) => {
        clearTimeout(timer)
        reject(error)
      },
    )
  })
}
