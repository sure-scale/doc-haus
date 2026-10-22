import { useContext, useRef, useState, type ReactNode } from "react"
import { Closing } from "./useExit"
import type { CompareOptions, CompareParams } from "../api/ingest/compare"
import { uploadDocument, type Document } from "../api/ingest/documents"
import { humanizeError } from "../api/ingest/errors"
import ModalHeader from "./ModalHeader"
import { useToast } from "./Toast"
import backdropClose from "./backdropClose"

const DOCX = /\.docx$/i

export const fileStem = (name: string) => name.replace(/\.[^.]+$/, "")

// The matter's other Word files, sources and deliverables alike, newest first.
export function compareCandidates(documents: Document[], name: string) {
  return documents.filter((d) => d.name !== name && DOCX.test(d.name)).sort((a, b) => b.created_at - a.created_at)
}

export type CompareChoice = {
  name: string
  other: string
  // Which side this file is on: "left" is the earlier version.
  thisSide: "left" | "right"
  thisLabel: string
  otherLabel: string
  options: CompareOptions
}

// Labels follow their files across a swap; a cleared label is left to the
// route, which names the side after its file.
export function choiceParams(choice: CompareChoice): CompareParams {
  const mine = { name: choice.name, label: choice.thisLabel.trim() }
  const theirs = { name: choice.other, label: choice.otherLabel.trim() }
  const [left, right] = choice.thisSide === "left" ? [mine, theirs] : [theirs, mine]
  return {
    left: left.name,
    right: right.name,
    ...(left.label ? { leftLabel: left.label } : {}),
    ...(right.label ? { rightLabel: right.label } : {}),
    ...choice.options,
  }
}

const OPTIONS: { key: keyof CompareOptions; label: string }[] = [
  { key: "ignoreFormatting", label: "Ignore formatting-only changes" },
  { key: "detectMoves", label: "Detect moved text" },
  { key: "headersFooters", label: "Include headers and footers" },
]

// Pick the other version of a Word file and how to compare them. This file
// stays fixed on one side; the other comes from the matter or a new upload,
// which always keeps both files rather than replacing one.
export default function CompareDialog({
  matterId,
  name,
  documents,
  onUploaded,
  onCompare,
  onClose,
}: {
  matterId: string
  name: string
  documents: Document[]
  // A file was added to the matter; reload its document list.
  onUploaded: () => void
  onCompare: (params: CompareParams) => void
  onClose: () => void
}) {
  const closing = useContext(Closing)
  const [thisSide, setThisSide] = useState<"left" | "right">("left")
  const [other, setOther] = useState("")
  const [thisLabel, setThisLabel] = useState(fileStem(name))
  const [otherLabel, setOtherLabel] = useState("")
  const [options, setOptions] = useState<CompareOptions>({ ignoreFormatting: true, detectMoves: true, headersFooters: true })
  const [upload, setUpload] = useState<"uploading" | "reading">()
  const [error, setError] = useState<string>()
  const fileRef = useRef<HTMLInputElement>(null)
  const toast = useToast()

  const candidates = compareCandidates(documents, name).map((d) => d.name)
  // A fresh upload is picked before the matter's list has reloaded.
  if (other && !candidates.includes(other)) candidates.unshift(other)

  // An upload in flight lands in the matter either way; stay open until it
  // does, so the new file is picked here rather than added unseen.
  function close() {
    if (!upload) onClose()
  }

  function pick(next: string) {
    setOther(next)
    setOtherLabel(fileStem(next))
  }

  async function add(file: File | undefined) {
    if (!file) return
    setError(undefined)
    if (!DOCX.test(file.name)) {
      setError("Choose a Word (.docx) file to compare.")
      return
    }
    setUpload("uploading")
    try {
      const result = await uploadDocument(matterId, file, { existing: "keep-both", onSent: () => setUpload("reading") })
      onUploaded()
      pick(result.name)
      if (result.injection)
        toast("warning", `${result.name} is ready, with a warning — possible prompt injection: ${[
          ...new Set(result.injection.findings.map((f) => f.detail)),
        ].join("; ")}. The assistant will treat this content as untrusted.`)
    } catch (cause) {
      setError(humanizeError(cause, `Could not add ${file.name}.`))
    } finally {
      setUpload(undefined)
      if (fileRef.current) fileRef.current.value = ""
    }
  }

  const thisFile = (
    <div className="compare-file" title={name}>
      <span className="compare-file-name">{name}</span>
      <span className="compare-file-tag">This file</span>
    </div>
  )
  const otherFile = (
    <div className="compare-pick">
      <select aria-label="Other version" value={other} onChange={(event) => pick(event.target.value)} disabled={upload !== undefined} autoFocus>
        <option disabled value="">
          {candidates.length ? "Choose a file…" : "No other Word files yet"}
        </option>
        {candidates.map((candidate) => (
          <option key={candidate} value={candidate}>
            {candidate}
          </option>
        ))}
      </select>
      <button type="button" onClick={() => fileRef.current?.click()} disabled={upload !== undefined}>
        {upload === "reading" ? "Reading…" : upload === "uploading" ? "Uploading…" : "Upload a file…"}
      </button>
      <input ref={fileRef} type="file" accept=".docx" hidden onChange={(event) => add(event.target.files?.[0])} />
    </div>
  )
  const mine = { file: thisFile, label: thisLabel, setLabel: setThisLabel, placeholder: fileStem(name) }
  const theirs = { file: otherFile, label: otherLabel, setLabel: setOtherLabel, placeholder: other ? fileStem(other) : "", disabled: !other }
  const [left, right] = thisSide === "left" ? [mine, theirs] : [theirs, mine]

  return (
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => close())}>
      <div className="picker-panel compare-dialog" role="dialog" aria-modal="true" aria-labelledby="compare-dialog-title" onClick={(e) => e.stopPropagation()}>
        <ModalHeader title={<span id="compare-dialog-title">Compare versions</span>} onClose={close} />
        <form
          className="picker-body compare-form"
          onSubmit={(event) => {
            event.preventDefault()
            if (!other || upload) return
            onCompare(choiceParams({ name, other, thisSide, thisLabel, otherLabel, options }))
          }}
        >
          <CompareSide id="compare-left" role="Earlier version" {...left} />
          <button
            type="button"
            className="compare-swap"
            onClick={() => setThisSide((side) => (side === "left" ? "right" : "left"))}
            title="Swap the earlier and later versions"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="m7 4-4 4 4 4M3 8h14M17 20l4-4-4-4M21 16H7" />
            </svg>
            Swap
          </button>
          <CompareSide id="compare-right" role="Later version" {...right} />
          <p className="compare-hint">The later version's label is the author name on its changes in the Word file.</p>
          <fieldset className="compare-options">
            <legend>Options</legend>
            {OPTIONS.map((option) => (
              <label key={option.key}>
                <input
                  type="checkbox"
                  checked={options[option.key]}
                  onChange={(event) => setOptions((current) => ({ ...current, [option.key]: event.target.checked }))}
                />
                {option.label}
              </label>
            ))}
          </fieldset>
          {error && (
            <p className="compare-error" role="alert">
              {error}
            </p>
          )}
          <div className="compare-actions">
            <button type="button" onClick={close} disabled={upload !== undefined}>
              Cancel
            </button>
            <button type="submit" className="primary" disabled={!other || upload !== undefined}>
              Compare
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}

function CompareSide({ id, role, file, label, setLabel, placeholder, disabled }: {
  id: string
  role: string
  file: ReactNode
  label: string
  setLabel: (value: string) => void
  placeholder: string
  disabled?: boolean
}) {
  return (
    <div className="compare-side">
      <span className="compare-side-role">{role}</span>
      {file}
      <label className="compare-label" htmlFor={`${id}-label`}>
        Label
      </label>
      <input id={`${id}-label`} value={label} placeholder={placeholder} maxLength={120} disabled={disabled} onChange={(event) => setLabel(event.target.value)} />
    </div>
  )
}

// Two arrows passing each other: compare one version with another.
export function CompareIcon({ size = 15 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="m16 3 4 4-4 4" />
      <path d="M20 7H4" />
      <path d="m8 21-4-4 4-4" />
      <path d="M4 17h16" />
    </svg>
  )
}
