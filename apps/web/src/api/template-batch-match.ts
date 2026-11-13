import {
  readSuggestion,
  suggestionColumns,
  type BatchField,
} from "../../../../dochaus/lib/template-batch"
import { getMatter } from "./ingest/matters"
import { location, matterClient } from "./opencode/client"
import { getConfig, modelSpec } from "./opencode/config"
import { abortSession, deleteSession, listMessageInfos } from "./opencode/sessions"

// One small job per spreadsheet: the fields the rules left unmatched, the free
// column headings and the first rows' values. Code keeps only a supplied
// column, at most one per field; the lawyer still confirms in the Match step.
export const MAX_SUGGEST_ROWS = 3
export const MAX_SUGGEST_CHARS = 80

export type SuggestInput = {
  fields: BatchField[]
  columns: string[]
  rows: string[][]
  mapping: Record<string, string | null>
}

const clip = (text: string) => {
  const flat = text.replace(/\s+/g, " ").trim()
  return flat.length <= MAX_SUGGEST_CHARS ? flat : `${flat.slice(0, MAX_SUGGEST_CHARS - 1).trimEnd()}…`
}

// The field keys are the template's exact placeholder text, so the reply can
// name a field without repeating its label. Headings and cells are data.
export function matchPrompt(fields: BatchField[], free: string[], rows: string[][]) {
  return [
    "<request>Match each template field to one supplied spreadsheet column, or leave it unmatched. Read the headings and the example values together.</request>",
    `<fields>${JSON.stringify(fields.map((field) => ({ key: field.key, label: field.label, ...(field.hint ? { hint: field.hint } : {}) })))}</fields>`,
    `<columns>${JSON.stringify(free)}</columns>`,
    `<example_rows>${JSON.stringify(rows.map((row) => row.map(clip)))}</example_rows>`,
    "<format>Reply with JSON only: {\"matches\":{\"<field key>\":\"<column heading exactly as given>\"}}. Omit a field, or give it null, when the headings and values do not support a match.</format>",
  ].join("\n")
}

// One bounded native job on the configured fast model, as compareColumn runs.
// The id is minted before admission so cleanup removes the session even when
// the create response is lost; an unsettled turn is interrupted first.
export async function suggestMatches(matterId: string, input: SuggestInput, signal: AbortSignal): Promise<Record<string, string | null> | undefined> {
  const { keys, free } = suggestionColumns(input.fields, input.columns, input.mapping)
  if (!keys.length || !free.length) return undefined
  signal.throwIfAborted()
  const config = await getConfig()
  signal.throwIfAborted()
  const model = modelSpec({ model: config.small_model ?? config.model })
  if (!model) return undefined
  const matter = await getMatter(matterId)
  signal.throwIfAborted()
  const fields = input.fields.filter((field) => keys.includes(field.key))
  const client = matterClient(matter.dir)
  const sessionID = `ses_${crypto.randomUUID().replaceAll("-", "")}`
  let settled = false
  try {
    await client.session.create({
      id: sessionID,
      title: "doc.haus template column matching",
      agent: "template-match",
      model: { providerID: model.providerID, id: model.modelID },
      location: location(client),
      metadata: { dochaus_role: "template-match" },
    }, { signal })
    await client.session.prompt({
      sessionID,
      text: matchPrompt(fields, free, input.rows.slice(0, MAX_SUGGEST_ROWS)),
      delivery: "queue",
    }, { signal })
    await client.session.wait({ sessionID }, { signal })
    settled = true
    const messages = await listMessageInfos(client, sessionID, signal)
    const terminal = messages.at(-1)
    if (terminal?.type !== "idle" || terminal.outcome !== "succeeded") return undefined
    const reply = messages.findLast((message) => message.type === "assistant" && message.content.some((part) => part.type === "text" && part.text.trim()))
    const text = reply?.type === "assistant" ? reply.content.flatMap((part) => part.type === "text" ? [part.text] : []).join("") : ""
    return readSuggestion(text, keys, free)
  } finally {
    if (!settled) await abortSession(client, sessionID).catch(() => {})
    await deleteSession(client, sessionID).catch(() => {})
  }
}
