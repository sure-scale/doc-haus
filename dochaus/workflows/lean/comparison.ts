import { readFileSync, realpathSync } from "node:fs"
import path from "node:path"
import { plannedDocument, type DocID } from "./plan"
import { matterSourceIDs } from "./quotes"
import { leanFile } from "./state"

/** A comparison source is expressly selected by the request/outline. Template
 * reuse alone never makes a source the redline baseline. */
export function comparisonBaseline(directory: string, document: DocID) {
  const source = plannedDocument(directory, document)?.comparison
  if (!source) return { name: "the first draft", bytes: readFileSync(leanFile(directory, `${document}/body.docx`)) }
  if (path.basename(source) !== source || !/\.docx$/i.test(source) || !matterSourceIDs(directory).includes(source))
    throw new Error(`The requested redline baseline "${source}" must name an indexed matter DOCX document.`)
  const root = realpathSync(directory)
  const file = realpathSync(path.join(root, source))
  const relative = path.relative(root, file)
  if (relative.startsWith("..") || path.isAbsolute(relative)) throw new Error("The redline baseline must remain inside the matter.")
  return { name: source, bytes: readFileSync(file) }
}

/** Original bytes affect rendering even when canonical extracted text is unchanged. */
export function comparisonBasis(directory: string, document: DocID) {
  return plannedDocument(directory, document)?.comparison ? comparisonBaseline(directory, document).bytes.toString("base64") : ""
}

/** The redline's baseline in the words the notes may repeat to the lawyer. */
export function comparisonDescription(directory: string, document: DocID) {
  const original = plannedDocument(directory, document)?.comparison
  return original ? `Redline compares the supplied original "${original}" with the final document.` : "Redline shows the changes made in review to the first draft."
}
