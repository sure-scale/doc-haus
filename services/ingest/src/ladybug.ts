import { Connection, Database } from "@ladybugdb/core"
import { existsSync, mkdirSync, realpathSync, statSync } from "node:fs"
import path from "node:path"

const GRAPH_DIRECTORY = path.join(".dochaus", "ladybug")
const GRAPH_DATABASE = "matter.lbdb"

export function matterGraphPath(matterDir: string) {
  if (!path.isAbsolute(matterDir)) throw new Error("Matter graph access requires an absolute matter directory.")
  if (!existsSync(matterDir) || !statSync(matterDir).isDirectory())
    throw new Error("Matter graph access requires an existing matter directory.")
  return path.join(realpathSync(matterDir), GRAPH_DIRECTORY, GRAPH_DATABASE)
}

export async function openMatterGraph(matterDir: string, options: { readOnly?: boolean } = {}) {
  const databasePath = matterGraphPath(matterDir)
  if (options.readOnly && !existsSync(databasePath)) throw new Error("The matter graph sidecar does not exist.")
  if (!options.readOnly) mkdirSync(path.dirname(databasePath), { recursive: true })

  const database = new Database(databasePath, 0, true, options.readOnly ?? false)
  const connection = new Connection(database)
  try {
    await database.init()
    await connection.init()
  } catch (error) {
    await connection.close()
    await database.close()
    throw error
  }

  return {
    databasePath,
    connection,
    close: () => connection.close().finally(() => database.close()),
  }
}
