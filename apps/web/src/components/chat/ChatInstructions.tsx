import type { Turn } from "./turn-model"

// Instructions from this chat (#364): a drafting run reads only request.md, so
// an instruction or exclusion the lawyer typed in this chat before Start would
// never reach it. Before a full run starts, the lawyer's own earlier messages
// (never an assistant's) are offered, pre-filled and included, for them to
// edit or leave out; ingest adds what they keep to request.md.

/** Ingest's limit on the carried text, as on the request itself. */
const LIMIT = 20_000

export type ChatInstructionsValue = { text: string; include: boolean; count: number }

/** The lawyer's own messages among `before`, oldest first, one paragraph each;
 * canned workflow prompts are not the lawyer's words. The oldest give way when
 * they would not fit the limit. */
export function earlierInstructions(before: Turn[], workflowPrompts: string[] = []) {
  const canned = new Set(workflowPrompts.map((prompt) => prompt.trim()))
  const own = before.filter((turn) => turn.role === "user").map((turn) => turn.text.trim()).filter((text) => text && !canned.has(text))
  const kept: string[] = []
  let length = 0
  for (const text of [...own].reverse()) {
    length += text.length + (kept.length ? 2 : 0)
    if (length > LIMIT) break
    kept.unshift(text)
  }
  if (!kept.length && own.length) kept.push(own.at(-1)!.slice(-LIMIT))
  return { text: kept.join("\n\n"), count: kept.length }
}

/** What the start request carries: the kept text, or nothing. */
export function carriedInstructions(value: ChatInstructionsValue | undefined) {
  const text = value?.include ? value.text.trim() : ""
  return text ? text : undefined
}

export default function ChatInstructions({
  value,
  onChange,
  disabled,
}: {
  value: ChatInstructionsValue
  onChange: (change: Partial<Pick<ChatInstructionsValue, "text" | "include">>) => void
  disabled?: boolean
}) {
  return (
    <fieldset className="chat-instructions">
      <legend className="question-label">Instructions from this chat</legend>
      <label className="chat-instructions-include">
        <input
          type="checkbox"
          checked={value.include}
          disabled={disabled}
          onChange={(event) => onChange({ include: event.target.checked })}
        />{" "}
        Include your {value.count === 1 ? "earlier message" : `${value.count} earlier messages`} in this run
      </label>
      <details className="chat-instructions-edit">
        <summary>Review or edit</summary>
        <textarea
          className="question-answer"
          aria-label="Instructions from this chat"
          value={value.text}
          maxLength={LIMIT}
          disabled={disabled}
          onChange={(event) => onChange({ text: event.target.value })}
        />
      </details>
    </fieldset>
  )
}
