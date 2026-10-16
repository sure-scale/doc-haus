import { useEffect, useRef, useState } from "react"
import { useSearchParams } from "react-router-dom"
import { deleteSkill, importSkill, listSkills, setSkillEnabled, type Skill } from "../api/ingest/library"
import { useToast } from "../components/Toast"
import Dropzone from "../components/Dropzone"
import { humanizeError } from "../api/ingest/errors"
import ChatPanel from "../components/ChatPanel"
import LibraryLayout from "../components/LibraryLayout"
import { DocMarkdown } from "../components/markdown/Markdown"
import { SKILL_BUILDER } from "../agents"
import ModalHeader from "../components/ModalHeader"

// The firm's skill library: reference knowledge (clause standards, checklists,
// drafting guidance) the specialist agents load on demand. Repo-shipped skills
// are listed read-only; custom ones live in WORKSPACE_ROOT/.skills, owned by the
// ingest service. Beside the list the Skill Builder chat composes and maintains
// skills conversationally — it runs in the library directory (there is no matter
// to scope to), pinned to the skill-builder agent. A dropzone imports an existing
// firm document (.md keeps its frontmatter; .docx/.pdf/.txt arrive as raw text
// for the builder to refine).
export default function Skills() {
  const [params, setParams] = useSearchParams()
  const session = params.get("session") ?? undefined
  const [dir, setDir] = useState<string>()
  const [skills, setSkills] = useState<Skill[]>([])
  const [loading, setLoading] = useState(true)
  const [busy, setBusy] = useState(false)
  const [confirmName, setConfirmName] = useState<string | null>(null)
  const [viewing, setViewing] = useState<Skill>()
  const createdRef = useRef<string | undefined>(undefined)
  const toast = useToast()

  function refresh() {
    return listSkills().then((res) => {
      setDir(res.dir)
      setSkills(res.skills)
    })
  }

  useEffect(() => {
    refresh()
      .catch((e) => toast("error", humanizeError(e)))
      .finally(() => setLoading(false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (!confirmName) return
    const dismiss = () => setConfirmName(null)
    window.addEventListener("click", dismiss)
    return () => window.removeEventListener("click", dismiss)
  }, [confirmName])

  async function onFiles(files: File[]) {
    const importable = files.filter((f) => /\.(md|txt|docx|pdf)$/i.test(f.name))
    if (!importable.length) {
      toast("error", "Skills import .md, .txt, .docx, or .pdf files.")
      return
    }
    setBusy(true)
    for (const file of importable) {
      try {
        const skill = await importSkill(file)
        setSkills((prev) => [...prev.filter((s) => s.name !== skill.name), skill])
        toast("success", `Imported skill "${skill.name}". Ask the skill builder to refine it.`)
      } catch (e) {
        toast("error", humanizeError(e, `Could not import ${file.name}.`))
      }
    }
    setBusy(false)
  }

  async function onDelete(s: Skill) {
    setConfirmName(null)
    try {
      await deleteSkill(s.name)
      setSkills((prev) => prev.filter((x) => x.name !== s.name))
      toast("success", `Deleted skill "${s.name}".`)
    } catch (e) {
      toast("error", humanizeError(e, `Could not delete ${s.name}.`))
    }
  }

  // Flip the row straight away so the switch feels instant; revert on failure.
  async function onToggle(s: Skill) {
    const enabled = !s.enabled
    setSkills((prev) => prev.map((x) => (x.name === s.name ? { ...x, enabled } : x)))
    try {
      await setSkillEnabled(s.name, enabled)
    } catch (e) {
      setSkills((prev) => prev.map((x) => (x.name === s.name ? { ...x, enabled: s.enabled } : x)))
      toast("error", humanizeError(e, `Could not update ${s.name}.`))
    }
  }

  const visible = [...skills].sort((a, b) => Number(b.builtin) - Number(a.builtin) || a.name.localeCompare(b.name))

  return (
    <>
      <LibraryLayout
        title="Skills"
        description="Reusable guidance and checklists for your assistants."
        libraryTitle="Available skills"
        count={loading ? undefined : skills.length}
        active={Boolean(session)}
        upload={
          <Dropzone
            accept=".md,.txt,.docx,.pdf"
            busy={busy}
            onFiles={onFiles}
            title={busy ? "Importing…" : "Import a skill"}
            hint="Drop a guide (.md, .txt, .docx, .pdf) or browse files."
          />
        }
        assistant={
          dir && (
            <ChatPanel
              key={session ?? "new"}
              directory={dir}
              sessionID={session}
              created={createdRef.current === session}
              agent={SKILL_BUILDER.name}
              available={new Set([SKILL_BUILDER.name])}
              pinned={SKILL_BUILDER}
              title="Create or edit a skill"
              emptyHint="Describe your firm’s guidance, or ask about your skills."
              starters={[
                "What skills do we have?",
                "Create a skill with our positions on limitation of liability.",
                "Refine the skill I just imported.",
              ]}
              composerPlaceholder="e.g. Create a skill for our standard confidentiality positions"
              onAgentChange={() => {}}
              onSessionCreated={(sid) => {
                createdRef.current = sid
                setParams({ session: sid }, { replace: true })
              }}
              onViewDocument={() => {}}
              onDocumentsChanged={refresh}
            />
          )
        }
      >
        {loading ? (
          <div className="skeleton-list">
            <div className="skeleton-row" />
            <div className="skeleton-row" />
            <div className="skeleton-row" />
          </div>
        ) : visible.length === 0 ? (
          <p className="library-empty">Import a guide or describe a skill to get started.</p>
        ) : (
          <ul className="matter-list">
            {visible.map((s) => (
              <li
                key={s.name}
                className="list-row list-row-clickable"
                title="View this skill"
                onClick={() => setViewing(s)}
              >
                <div className="list-row-main">
                  <span className="list-row-title">
                    {s.name}
                    {s.builtin && <span className="badge badge-quiet">Built-in</span>}
                  </span>
                  <span className="list-row-meta">
                    <span className="list-row-desc">{s.description || "No description yet"}</span>
                  </span>
                </div>
                <div className="list-row-actions">
                  <button
                    role="switch"
                    aria-checked={s.enabled}
                    aria-label={s.name}
                    className={s.enabled ? "switch on" : "switch"}
                    title={s.enabled ? "Disable skill — agents stop loading it" : "Enable skill"}
                    onClick={(e) => {
                      e.stopPropagation()
                      onToggle(s)
                    }}
                  />
                  <button
                    className="icon-btn"
                    title="View skill"
                    onClick={(e) => {
                      e.stopPropagation()
                      setViewing(s)
                    }}
                  >
                    View
                  </button>
                  {!s.builtin &&
                    (confirmName === s.name ? (
                      <button
                        className="icon-btn danger"
                        title="Confirm delete"
                        onClick={(e) => {
                          e.stopPropagation()
                          onDelete(s)
                        }}
                      >
                        Confirm
                      </button>
                    ) : (
                      <button
                        className="icon-btn"
                        title="Delete skill"
                        onClick={(e) => {
                          e.stopPropagation()
                          setConfirmName(s.name)
                        }}
                      >
                        Delete
                      </button>
                    ))}
                </div>
              </li>
            ))}
          </ul>
        )}
      </LibraryLayout>

      {viewing && (
        <div className="viewer-overlay" onClick={() => setViewing(undefined)}>
          <div className="picker-panel doc-panel" onClick={(e) => e.stopPropagation()}>
            <ModalHeader
              title={`${viewing.name}${viewing.builtin ? " (built-in)" : ""}`}
              onClose={() => setViewing(undefined)}
            />
            <div className="picker-body">
              {viewing.description && <p className="muted">{viewing.description}</p>}
              <DocMarkdown>{viewing.content}</DocMarkdown>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
