import { Link } from "react-router-dom"
import JurisdictionSelect from "../JurisdictionSelect"
import type { Jurisdiction, MatterDetail as Detail } from "../../api/ingest/matters"
import { POSTURES, type Posture } from "../../api/ingest/preferences"

/** The title-area fields one save changes; "" clears the matter's Posture. */
export type MatterFields = Partial<Pick<Detail, "title" | "reference" | "jurisdictions" | "playbook" | "client">> & { posture?: Posture | "" }

export default function MatterHeader({
  matter,
  jurisdictions,
  editingRef,
  setEditingRef,
  draftRef,
  setDraftRef,
  commitRef,
  editingClient,
  setEditingClient,
  draftClient,
  setDraftClient,
  commitClient,
  renaming,
  setRenaming,
  draftTitle,
  setDraftTitle,
  commitRename,
  patch,
}: {
  matter: Detail
  jurisdictions: Jurisdiction[]
  editingRef: boolean
  setEditingRef: (editing: boolean) => void
  draftRef: string
  setDraftRef: (draft: string) => void
  commitRef: () => Promise<void>
  editingClient: boolean
  setEditingClient: (editing: boolean) => void
  draftClient: string
  setDraftClient: (draft: string) => void
  commitClient: () => Promise<void>
  renaming: boolean
  setRenaming: (renaming: boolean) => void
  draftTitle: string
  setDraftTitle: (draft: string) => void
  commitRename: () => Promise<void>
  patch: (fields: MatterFields) => Promise<void>
}) {
  const refField = editingRef ? (
    <input
      className="matter-ref-input"
      autoFocus
      placeholder="Matter number"
      aria-label="Matter number"
      value={draftRef}
      onChange={(e) => setDraftRef(e.target.value)}
      onBlur={commitRef}
      onKeyDown={(e) => {
        if (e.key === "Enter") commitRef()
        if (e.key === "Escape") setEditingRef(false)
      }}
    />
  ) : (
    <button
      type="button"
      className={matter.reference ? "matter-ref matter-field" : "matter-meta-add"}
      title={matter.reference ? "Edit matter number" : "Add a matter number"}
      aria-label={matter.reference ? `Edit matter number: ${matter.reference}` : "Add a matter number"}
      onClick={() => {
        setDraftRef(matter.reference ?? "")
        setEditingRef(true)
      }}
    >
      {matter.reference || "+ Matter number"}
    </button>
  )

  const clientField = editingClient ? (
    <input
      className="matter-ref-input"
      autoFocus
      placeholder="Whom you act for"
      aria-label="Client you act for"
      maxLength={200}
      value={draftClient}
      onChange={(e) => setDraftClient(e.target.value)}
      onBlur={commitClient}
      onKeyDown={(e) => {
        if (e.key === "Enter") commitClient()
        if (e.key === "Escape") setEditingClient(false)
      }}
    />
  ) : (
    <button
      type="button"
      className={matter.client ? "matter-ref matter-field" : "matter-meta-add"}
      title={matter.client ? "Edit the client you act for" : "Add the client you act for"}
      aria-label={matter.client ? `Edit client: ${matter.client}` : "Add the client you act for"}
      onClick={() => {
        setDraftClient(matter.client ?? "")
        setEditingClient(true)
      }}
    >
      {matter.client ? `Client: ${matter.client}` : "+ Client"}
    </button>
  )

  // How hard drafts lean toward the client here: one choice, saved as soon as it is picked.
  const postureField = (
    <select
      className={`matter-posture${matter.posture ? "" : " is-default"}`}
      aria-label="Posture for this matter"
      title="How hard drafts lean toward your client on this matter"
      value={matter.posture ?? ""}
      onChange={(e) => patch({ posture: e.target.value as Posture | "" })}
    >
      <option value="">Posture: firm default</option>
      {POSTURES.map((row) => (
        <option key={row.value} value={row.value}>
          Posture: {row.label}
        </option>
      ))}
    </select>
  )

  // A permanent route back to the matter list, then one quiet line of facts: the matter number, the
  // client, the Posture and the jurisdictions, each a ghost "+ …" button (the Posture its firm
  // default) until set, then its chip.
  return (
    <header className="page-head matter-page-head">
      <div className="matter-head">
        <nav className="matter-breadcrumb" aria-label="Breadcrumb">
          <ol>
            <li><Link to="/">Matters</Link></li>
            <li className="matter-breadcrumb-current" aria-current="page">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true">
                <path d="m9 5 7 7-7 7" />
              </svg>
              <h1>
                {renaming ? (
                  <input
                    className="matter-title-input"
                    autoFocus
                    aria-label="Matter title"
                    value={draftTitle}
                    onChange={(e) => setDraftTitle(e.target.value)}
                    onBlur={commitRename}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") commitRename()
                      if (e.key === "Escape") setRenaming(false)
                    }}
                  />
                ) : (
                  <button
                    type="button"
                    className="matter-title"
                    title="Rename matter"
                    aria-label={`Rename matter: ${matter.title}`}
                    onClick={() => {
                      setDraftTitle(matter.title)
                      setRenaming(true)
                    }}
                  >
                    {matter.title}
                    <svg className="matter-title-pencil" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                      <path d="M12 20h9" />
                      <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z" />
                    </svg>
                  </button>
                )}
              </h1>
            </li>
          </ol>
        </nav>

        <div className="matter-meta">
          {refField}
          <span className="matter-meta-sep" aria-hidden>
            ·
          </span>
          {clientField}
          <span className="matter-meta-sep" aria-hidden>
            ·
          </span>
          {postureField}
          <span className="matter-meta-sep" aria-hidden>
            ·
          </span>
          <JurisdictionSelect
            jurisdictions={jurisdictions}
            selected={matter.jurisdictions ?? []}
            onChange={(codes) => patch({ jurisdictions: codes })}
            trigger={
              (matter.jurisdictions?.length ?? 0) === 0 ? (
                <span className="matter-meta-add">+ Jurisdiction</span>
              ) : (
                <span className="matter-jx-badges">
                  {matter.jurisdictions?.map((code) => (
                    <span key={code} className="matter-badge">
                      {jurisdictions.find((j) => j.code === code)?.name ?? code}
                    </span>
                  ))}
                </span>
              )
            }
          />
        </div>
      </div>

    </header>
  )
}

