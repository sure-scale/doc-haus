import { useContext, useEffect, useRef, useState, type ReactNode } from "react"
import { Closing } from "./useExit"
import type { CompareOptions, CompareParams } from "../api/ingest/compare"
import { duplicateMessage, DuplicateUploadError, uploadDocument, type Document } from "../api/ingest/documents"
import { humanizeError } from "../api/ingest/errors"
import { fetchDocumentVersions, type DocumentVersions } from "../api/ingest/versions"
import { previousVersion, versionOption, versionSideLabel, versionUnavailable } from "./DocumentVersions"
import ModalHeader from "./ModalHeader"
import { useToast } from "./Toast"
import backdropClose from "./backdropClose"

const DOCX = /\.docx$/i

export const fileStem = (name: string) => name.replace(/\.[^.]+$/, "")

// The matter's other Word files, sources and deliverables alike, newest first.
export function compareCandidates(documents: Document[], name: string) {
  return documents.filter((d) => d.name !== name && DOCX.test(d.name)).sort((a, b) => b.created_at - a.created_at)
}

// A select value naming a kept version of this file rather than a matter file:
// a file name never contains "/".
const VERSION = "/version/"

export type CompareChoice = {
  name: string
  other: string
  // Kept versions (#268): this file's side, and the other side when it is a
  // version of this file (other is then this file's name). Absent means the
  // working file.
  thisVersion?: string
  otherVersion?: string
  // Which side this file is on: "left" is the earlier version.
  thisSide: "left" | "right"
  thisLabel: string
  otherLabel: string
  options: CompareOptions
}

// Labels follow their files across a swap; a cleared label is left to the
// route, which names the side after its file.
export function choiceParams(choice: CompareChoice): CompareParams {
  const mine = { name: choice.name, version: choice.thisVersion, label: choice.thisLabel.trim() }
  const theirs = { name: choice.other, version: choice.otherVersion, label: choice.otherLabel.trim() }
  const [left, right] = choice.thisSide === "left" ? [mine, theirs] : [theirs, mine]
  return {
    left: left.name,
    right: right.name,
    ...(left.version ? { leftVersion: left.version } : {}),
    ...(right.version ? { rightVersion: right.version } : {}),
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
// stays fixed on one side; the other comes from its kept versions, the matter
// or a new upload, which always keeps both files rather than replacing one.
// With kept versions it starts on the previous version, this file the later
// side, and either side can be any kept version.
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
  const [notice, setNotice] = useState<string>()
  const [kept, setKept] = useState<DocumentVersions>()
  const [thisVersion, setThisVersion] = useState("")
  // Set once the lawyer picks or uploads, so a late version list changes nothing.
  const picked = useRef(false)
  const fileRef = useRef<HTMLInputElement>(null)
  const toast = useToast()

  // Kept versions load once. Without any, or if they cannot be read, the
  // dialog offers the matter's files as before.
  useEffect(() => {
    let cancelled = false
    fetchDocumentVersions(matterId, name).then((data) => {
      if (cancelled) return
      setKept(data)
      const previous = previousVersion(data)
      const current = data.versions.find((version) => version.current)
      if (!previous || versionUnavailable(previous) || picked.current) return
      setOther(VERSION + previous.id)
      setOtherLabel(versionSideLabel(name, previous.number))
      setThisSide("right")
      if (data.working_version && current) setThisLabel((label) => (label === fileStem(name) ? versionSideLabel(name, current.number) : label))
    }, () => {})
    return () => {
      cancelled = true
    }
  }, [matterId, name])

  const versions = kept?.versions ?? []
  const events = kept?.events ?? []
  const byId = (id: string) => versions.find((version) => version.id === id)
  // The version this side holds: the one picked, or the working file's.
  const thisHeld = thisVersion || kept?.working_version || ""
  const otherVersions = versions.filter((version) => version.id !== thisHeld)
  const otherVersion = other.startsWith(VERSION) ? other.slice(VERSION.length) : undefined
  const otherName = otherVersion ? name : other
  const candidates = compareCandidates(documents, name).map((d) => d.name)
  // A fresh upload is picked before the matter's list has reloaded.
  if (otherName && !otherVersion && !candidates.includes(otherName)) candidates.unshift(otherName)

  // An upload in flight lands in the matter either way; stay open until it
  // does, so the new file is picked here rather than added unseen.
  function close() {
    if (!upload) onClose()
  }

  function pick(next: string) {
    picked.current = true
    setNotice(undefined)
    setOther(next)
    const version = next.startsWith(VERSION) ? byId(next.slice(VERSION.length)) : undefined
    setOtherLabel(version ? versionSideLabel(name, version.number) : fileStem(next))
  }

  function pickThis(id: string) {
    picked.current = true
    setThisVersion(id)
    const version = byId(id || kept?.working_version || "")
    setThisLabel(version ? versionSideLabel(name, version.number) : fileStem(name))
    // A side cannot hold the version the other side holds.
    if (other === VERSION + (id || kept?.working_version)) {
      setOther("")
      setOtherLabel("")
    }
  }

  async function add(file: File | undefined) {
    if (!file) return
    setError(undefined)
    if (!DOCX.test(file.name)) {
      setError("Choose a Word (.docx) file to compare.")
      return
    }
    picked.current = true
    setUpload("uploading")
    try {
      const result = await uploadDocument(matterId, file, { existing: "keep-both", onSent: () => setUpload("reading") })
      onUploaded()
      pick(result.name)
      if (result.notice) setNotice(result.notice)
      if (result.injection)
        toast("warning", `${result.name} is ready, with a warning — possible prompt injection: ${[
          ...new Set(result.injection.findings.map((f) => f.detail)),
        ].join("; ")}. The assistant will treat this content as untrusted.`)
    } catch (cause) {
      // Already in the matter: pick what it matched, when that can be compared.
      // Only this file's versions are offered, so an earlier version of another
      // file is named but not picked: picking that file would compare its current one.
      if (cause instanceof DuplicateUploadError) {
        const match = cause.existing !== name
          ? !cause.version || cause.version.current ? cause.existing : undefined
          : cause.version && cause.version.id !== thisHeld ? VERSION + cause.version.id : undefined
        if (match) pick(match)
        setNotice(`${duplicateMessage(cause.existing, cause.version)}${match ? " It is selected." : ""}`)
      } else setError(humanizeError(cause, `Could not add ${file.name}.`))
    } finally {
      setUpload(undefined)
      if (fileRef.current) fileRef.current.value = ""
    }
  }

  const current = versions.find((version) => version.current)
  // Each kept version can be picked; one whose copy was removed is listed, not chosen.
  const versionOptions = (list: typeof versions) => list.map((version) => {
    const gone = versionUnavailable(version)
    const option = versionOption(version, events)
    return (
      <option key={version.id} value={VERSION + version.id} disabled={Boolean(gone)} title={option.title}>
        {option.text}{gone ? " · not available" : ""}
      </option>
    )
  })
  const fileOptions = candidates.map((candidate) => (
    <option key={candidate} value={candidate}>
      {candidate}
    </option>
  ))
  const thisFile = (
    <>
      <div className="compare-file" title={name}>
        <span className="compare-file-name">{name}</span>
        <span className="compare-file-tag">This file</span>
      </div>
      {versions.length > 1 && (
        <select aria-label="Version of this file" value={thisVersion} onChange={(event) => pickThis(event.target.value)} disabled={upload !== undefined}>
          <option value="">{kept?.working_version && current ? `Current file (version ${current.number})` : "Current file"}</option>
          {versions.filter((version) => version.id !== kept?.working_version && !versionUnavailable(version)).map((version) => {
            const option = versionOption(version, events)
            return (
              <option key={version.id} value={version.id} title={option.title}>
                {option.text}
              </option>
            )
          })}
        </select>
      )}
    </>
  )
  const otherFile = (
    <div className="compare-pick">
      <select aria-label="Other version" value={other} onChange={(event) => pick(event.target.value)} disabled={upload !== undefined} autoFocus>
        <option disabled value="">
          {candidates.length || otherVersions.length ? "Choose a version or file…" : "No other Word files yet"}
        </option>
        {otherVersions.length ? (
          <>
            <optgroup label="Versions of this file">{versionOptions(otherVersions)}</optgroup>
            {fileOptions.length > 0 && <optgroup label="Other Word files">{fileOptions}</optgroup>}
          </>
        ) : (
          fileOptions
        )}
      </select>
      <button type="button" onClick={() => fileRef.current?.click()} disabled={upload !== undefined}>
        {upload === "reading" ? "Reading…" : upload === "uploading" ? "Uploading…" : "Upload a file…"}
      </button>
      <input ref={fileRef} type="file" accept=".docx" hidden onChange={(event) => add(event.target.files?.[0])} />
    </div>
  )
  const mine = { file: thisFile, label: thisLabel, setLabel: setThisLabel, placeholder: fileStem(name) }
  const theirs = { file: otherFile, label: otherLabel, setLabel: setOtherLabel, placeholder: otherName ? fileStem(otherName) : "", disabled: !other }
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
            onCompare(choiceParams({ name, other: otherName, thisVersion: thisVersion || undefined, otherVersion, thisSide, thisLabel, otherLabel, options }))
          }}
        >
          <CompareSide id="compare-left" role="Earlier version" {...left} />
          <button
            type="button"
            className="compare-swap"
            onClick={() => setThisSide((side) => (side === "left" ? "right" : "left"))}
            title="Swap the earlier and later versions"
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
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
          {notice && (
            <p className="compare-hint compare-notice" role="status">
              {notice}
            </p>
          )}
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
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="m16 3 4 4-4 4" />
      <path d="M20 7H4" />
      <path d="m8 21-4-4 4-4" />
      <path d="M4 17h16" />
    </svg>
  )
}
