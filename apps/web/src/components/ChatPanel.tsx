import { useEffect, useMemo, useRef, useState } from "react"
import type { Event, Part } from "@opencode-ai/sdk"
import {
  createSession,
  getMessages,
  matterClient,
  sendPrompt,
  subscribeEvents,
  type Citation,
  type Client,
} from "../api/opencode"
import CitationView from "./CitationView"
import Markdown from "./Markdown"
import ModelSelector from "./ModelSelector"
import WorkflowLauncher from "./WorkflowLauncher"

type Turn = { role: "user" | "assistant"; text: string; citations: Citation[] }

// Quiet starter prompts so a fresh matter is not a blank box — mirrors how Harvey
// and Legora seat the lawyer with ready questions about the documents in scope.
const STARTERS = [
  "What are the key obligations of each party?",
  "What termination rights does each party have?",
  "Flag any unusual or one-sided clauses.",
]

// Reduce a message's parts to its visible text + any search-document citations.
function contentOf(parts: Part[]) {
  const text = parts
    .filter((p): p is Extract<Part, { type: "text" }> => p.type === "text")
    .map((p) => p.text)
    .join("")
  const citations = parts
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => p.tool === "search-document" && p.state.status === "completed")
    .flatMap((p) => ((p.state as { metadata?: { citations?: Citation[] } }).metadata?.citations ?? []))
  return { text, citations }
}

// Same, over the live part map keyed by id, for one streaming message.
function readMessage(parts: Map<string, Part>, messageID: string) {
  return contentOf([...parts.values()].filter((p) => p.messageID === messageID))
}

export default function ChatPanel({
  directory,
  sessionID,
  agent,
  available,
  onAgentChange,
  onLaunchWorkflow,
}: {
  directory: string
  sessionID?: string
  agent: string
  available: Set<string>
  onAgentChange: (name: string) => void
  onLaunchWorkflow: (name: string) => void
}) {
  const client = useMemo<Client>(() => matterClient(directory), [directory])
  const [turns, setTurns] = useState<Turn[]>([])
  const [input, setInput] = useState("")
  const [busy, setBusy] = useState(false)
  const [, bump] = useState(0)

  const sessionRef = useRef<string>("")
  const partsRef = useRef<Map<string, Part>>(new Map())
  const rolesRef = useRef<Map<string, string>>(new Map())
  const assistantRef = useRef<string>("")
  const logRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const controller = new AbortController()
    if (sessionID) {
      sessionRef.current = sessionID
      getMessages(client, sessionID).then((msgs) =>
        setTurns(
          msgs
            .map((m) => ({ role: m.info.role, ...contentOf(m.parts) }))
            .filter((t) => t.text || t.citations.length),
        ),
      )
    }
    // No session until the first send (see onSend) — mounting the panel must not
    // mint an empty throwaway session that would clutter the conversation list.
    subscribeEvents(client, onEvent, controller.signal).catch(() => {})
    return () => controller.abort()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [client, sessionID])

  useEffect(() => {
    logRef.current?.scrollTo(0, logRef.current.scrollHeight)
  })

  function onEvent(event: Event) {
    if (event.type === "message.updated") {
      const info = event.properties.info
      if (info.sessionID !== sessionRef.current) return
      rolesRef.current.set(info.id, info.role)
      if (info.role === "assistant") assistantRef.current = info.id
      bump((n) => n + 1)
      return
    }
    if (event.type === "message.part.updated") {
      const part = event.properties.part
      if (part.sessionID !== sessionRef.current) return
      partsRef.current.set(part.id, part)
      bump((n) => n + 1)
      return
    }
    if (event.type === "session.idle" && event.properties.sessionID === sessionRef.current) {
      finalize()
    }
  }

  function finalize() {
    const id = assistantRef.current
    if (id) {
      const { text, citations } = readMessage(partsRef.current, id)
      if (text || citations.length) setTurns((prev) => [...prev, { role: "assistant", text, citations }])
    }
    partsRef.current.clear()
    assistantRef.current = ""
    setBusy(false)
  }

  async function onSend() {
    const text = input.trim()
    if (!text || busy) return
    setTurns((prev) => [...prev, { role: "user", text, citations: [] }])
    setInput("")
    setBusy(true)
    partsRef.current.clear()
    assistantRef.current = ""
    // Create the session lazily, titled from this first message so it reads as a
    // distinct conversation in the rail rather than an interchangeable "Q&A".
    if (!sessionRef.current) sessionRef.current = (await createSession(client, text.slice(0, 60))).id
    await sendPrompt(client, sessionRef.current, agent, text)
  }

  const live = assistantRef.current ? readMessage(partsRef.current, assistantRef.current) : null

  return (
    <div className="card">
      <div className="row" style={{ justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
        <h2 style={{ margin: 0 }}>Ask the matter</h2>
        <ModelSelector available={available} value={agent} onChange={onAgentChange} />
      </div>
      <div className="chat-log" ref={logRef}>
        {turns.length === 0 && !busy && (
          <div className="chat-empty">
            <p className="muted">Ask a question about this matter's documents. Every answer cites the source section.</p>
            <div className="starters">
              {STARTERS.map((s) => (
                <button key={s} className="starter" onClick={() => setInput(s)}>
                  {s}
                </button>
              ))}
            </div>
          </div>
        )}
        {turns.map((t, i) => (
          <div key={i} className={`msg ${t.role}`}>
            {t.role === "assistant" ? <Markdown>{t.text}</Markdown> : t.text}
            <CitationView citations={t.citations} />
          </div>
        ))}
        {busy && (
          <div className="msg assistant">
            {live?.text ? <Markdown>{live.text}</Markdown> : <span className="muted">Thinking...</span>}
            {live && <CitationView citations={live.citations} />}
          </div>
        )}
      </div>
      <div className="composer-tools">
        <WorkflowLauncher available={available} onLaunch={onLaunchWorkflow} />
      </div>
      <div className="composer">
        <textarea
          placeholder="e.g. What termination rights does each party have?"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) onSend()
          }}
        />
        <button className="primary" onClick={onSend} disabled={busy || !input.trim()}>
          Send
        </button>
      </div>
      <p className="muted" style={{ marginTop: 6, fontSize: 12 }}>
        Cmd/Ctrl + Enter to send. Answers cite [Document § section] from indexed documents.
      </p>
    </div>
  )
}
