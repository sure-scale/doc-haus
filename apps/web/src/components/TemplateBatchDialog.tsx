import { useCallback, useContext, useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"
import { useNavigate } from "react-router-dom"
import type { Template } from "../api/ingest/library"
import { listMatters, type Matter } from "../api/ingest/matters"
import { humanizeError } from "../api/ingest/errors"
import {
  createTemplateBatch,
  previewTemplateBatch,
  readTemplateBatch,
  rowMissing,
  templateBatchSampleUrl,
  type BatchField,
  type BatchMapping,
  type BatchResult,
  type BatchSheet,
} from "../api/ingest/template-batch"
import Dropzone from "./Dropzone"
import ModalHeader from "./ModalHeader"
import TemplateViewer from "./TemplateViewer"
import { Tooltip } from "./Tooltip"
import { Closing } from "./useExit"
import backdropClose from "./backdropClose"
import "./TemplateBatchDialog.css"

// "Fill from spreadsheet": one document per spreadsheet row from a library
// template, filled by ingest without any model. Choose a matter and drop a
// spreadsheet; match only the fields that did not match a column by name;
// preview each row as the filled document; create them all into one folder
// in the matter. Rows are identified by their own spreadsheet row number.
type Step = "choose" | "match" | "preview" | "done"

const plural = (count: number, noun: string) => `${count} ${noun}${count === 1 ? "" : "s"}`
const stem = (name: string) => name.replace(/\.docx$/i, "")
const IDENTIFYING = /name|party|client|company|counterparty|recipient|addressee|tenant|employee/i

// The column whose value tells rows apart in the stepper and lists: the
// server's choice while still matched, else the same rule on the current matching.
function identifyingColumn(fields: BatchField[], mapping: BatchMapping, preferred: string | null) {
  if (preferred && mapping[preferred]) return mapping[preferred]
  const mapped = fields.filter((field) => field.kind === "fill" && mapping[field.key])
  const field = mapped.find((item) => IDENTIFYING.test(item.label)) ?? mapped[0]
  return field ? mapping[field.key] ?? null : null
}

export default function TemplateBatchDialog({ template, onClose }: { template: Template; onClose: () => void }) {
  const navigate = useNavigate()
  const [step, setStep] = useState<Step>("choose")
  const [matters, setMatters] = useState<Matter[]>()
  const [matterId, setMatterId] = useState("")
  const [reading, setReading] = useState(false)
  const [error, setError] = useState<string>()
  const [sheet, setSheet] = useState<BatchSheet>()
  const [mapping, setMapping] = useState<BatchMapping>({})
  // Fill fields no column matched when the spreadsheet was read: the only ones
  // the Match step asks about, then and on "Change matching".
  const [manualKeys, setManualKeys] = useState<string[]>([])
  const [index, setIndex] = useState(0)
  const [skipped, setSkipped] = useState<Set<number>>(() => new Set())
  const [shownMissing, setShownMissing] = useState<{ row: number; missing: string[] }>()
  const [creating, setCreating] = useState(false)
  const [createError, setCreateError] = useState<string>()
  const [result, setResult] = useState<BatchResult>()
  const alive = useRef(true)
  // Set on every mount: React's development double mount runs this cleanup
  // once before the dialog is really in use.
  useEffect(() => {
    alive.current = true
    return () => { alive.current = false }
  }, [])

  useEffect(() => {
    listMatters()
      .then((list) => {
        if (!alive.current) return
        const sorted = [...list].sort((a, b) => b.created_at - a.created_at)
        setMatters(sorted)
        setMatterId((current) => current || sorted[0]?.id || "")
      })
      .catch((cause) => alive.current && setError(humanizeError(cause, "Could not load matters.")))
  }, [])

  const close = useCallback(() => { if (!creating) onClose() }, [creating, onClose])

  async function onFiles(files: File[]) {
    const file = files[0]
    if (!file) return
    if (!/\.(csv|xlsx|xls)$/i.test(file.name)) { setError("Choose a .csv, .xlsx or .xls file."); return }
    setReading(true)
    setError(undefined)
    try {
      const next = await readTemplateBatch(template.name, file)
      if (!alive.current) return
      if (!next.fields.length) { setError("This template has no named fields to fill."); return }
      if (!next.rows.length) { setError(`${next.source || file.name} has no rows to fill.`); return }
      const full = Object.fromEntries(next.fields.map((field) => [field.key, next.mapping[field.key] ?? null]))
      const unmatched = next.fields.filter((field) => field.kind === "fill" && !full[field.key]).map((field) => field.key)
      setSheet(next)
      setMapping(full)
      setManualKeys(unmatched)
      setIndex(0)
      setSkipped(new Set())
      setShownMissing(undefined)
      setStep(unmatched.length ? "match" : "preview")
    } catch (cause) {
      if (alive.current) setError(humanizeError(cause, `Could not read ${file.name}.`))
    } finally {
      if (alive.current) setReading(false)
    }
  }

  const rows = sheet?.rows ?? []
  const missing = useMemo(
    () => sheet ? sheet.rows.map((row) => rowMissing(sheet.fields, sheet.columns, row.values, mapping)) : [],
    [sheet, mapping],
  )
  const identity = sheet ? identifyingColumn(sheet.fields, mapping, sheet.identifyingKey) : null
  const identityAt = identity && sheet ? sheet.columns.indexOf(identity) : -1
  const rowName = (at: number) => (identityAt >= 0 ? rows[at]?.values[identityAt]?.trim() : "") || `Spreadsheet row ${rows[at]?.row}`
  const included = rows.filter((row) => !skipped.has(row.row)).length
  const needInput = rows.map((row, at) => ({ row, at })).filter(({ row, at }) => !skipped.has(row.row) && missing[at]?.length)
  const skippedRows = rows.map((row, at) => ({ row, at })).filter(({ row }) => skipped.has(row.row))

  const current = rows[index]
  const values = current?.values
  const loadBytes = useCallback(async () => {
    if (!sheet || !current || !values) return new Uint8Array()
    const preview = await previewTemplateBatch({ name: template.name, columns: sheet.columns, values, mapping })
    if (alive.current) setShownMissing({ row: current.row, missing: preview.missing })
    return preview.bytes
  }, [template.name, sheet, current, values, mapping])

  async function create() {
    if (!sheet || !matterId || creating) return
    setCreating(true)
    setCreateError(undefined)
    try {
      const done = await createTemplateBatch(matterId, {
        name: template.name,
        sha256: sheet.template.sha256,
        source: sheet.source,
        columns: sheet.columns,
        rows: sheet.rows,
        skipped: sheet.rows.filter((row) => skipped.has(row.row)).map((row) => row.row),
        mapping,
      })
      if (!alive.current) return
      setResult(done)
      setStep("done")
    } catch (cause) {
      if (alive.current) setCreateError(humanizeError(cause, "Could not create the documents."))
    } finally {
      if (alive.current) setCreating(false)
    }
  }

  function toggleSkip(row: number) {
    setSkipped((previous) => {
      const next = new Set(previous)
      if (!next.delete(row)) next.add(row)
      return next
    })
  }

  if (step === "preview" && sheet && current) {
    const rowMissingNow = shownMissing?.row === current.row ? shownMissing.missing : missing[index] ?? []
    const isSkipped = skipped.has(current.row)
    return <TemplateViewer
      name={template.name}
      loadBytes={loadBytes}
      onClose={close}
      sidebar={<PreviewRow
        position={index}
        total={rows.length}
        sheetRow={current.row}
        title={rowName(index)}
        missing={rowMissingNow}
        skipped={isSkipped}
        onStep={(delta) => setIndex((at) => Math.min(rows.length - 1, Math.max(0, at + delta)))}
        onSkip={() => toggleSkip(current.row)}
        // Always reachable: an automatic match can be wrong too. When every
        // field matched automatically, the lawyer asked to review them all.
        onChangeMatching={() => {
          if (!manualKeys.length) setManualKeys(sheet.fields.filter((field) => field.kind === "fill").map((field) => field.key))
          setStep("match")
        }}
      />}
      footer={<div className="template-batch-footer">
        <div className="template-batch-summary" aria-live="polite">
          <span>{included - needInput.length} ready</span>
          {needInput.length > 0 && <>
            <span aria-hidden="true">·</span>
            <RowMenu
              label={`${needInput.length} need input`}
              items={needInput.map(({ row, at }) => ({
                key: row.row,
                label: rowName(at),
                detail: `Row ${row.row} · ${missing[at]!.join(", ")}`,
                onSelect: () => setIndex(at),
              }))}
            />
          </>}
          {skippedRows.length > 0 && <>
            <span aria-hidden="true">·</span>
            <RowMenu
              label={`${skippedRows.length} skipped`}
              items={skippedRows.map(({ row, at }) => ({
                key: row.row,
                label: `Restore ${rowName(at)}`,
                detail: `Row ${row.row}`,
                onSelect: () => { toggleSkip(row.row); setIndex(at) },
              }))}
            />
          </>}
        </div>
        {createError && <p className="template-batch-error" role="alert">{createError}</p>}
        <button type="button" className="primary" disabled={!included || creating || !matterId} aria-busy={creating || undefined} onClick={() => void create()}>
          {creating ? `Creating ${plural(included, "document")}…` : `Create ${plural(included, "document")}`}
        </button>
      </div>}
    />
  }

  return <BatchPanel title={`Fill ${stem(template.name)} from a spreadsheet`} onClose={close} focusKey={step}
    actions={step === "match" ? <button type="button" className="primary" onClick={() => setStep("preview")}>Preview</button>
      : step === "done" && result ? <button type="button" className="primary" onClick={() => { onClose(); navigate(`/matter/${encodeURIComponent(matterId)}?view=documents`) }}>Open matter</button>
      : undefined}
  >
    {step === "choose" && <ChooseStep
      template={template}
      matters={matters}
      matterId={matterId}
      onMatter={setMatterId}
      reading={reading}
      error={error}
      onFiles={(files) => void onFiles(files)}
    />}
    {step === "match" && sheet && <MatchStep
      sheet={sheet}
      mapping={mapping}
      manualKeys={manualKeys}
      onMap={(key, column) => setMapping((previous) => ({ ...previous, [key]: column }))}
      onRestart={() => { setSheet(undefined); setError(undefined); setStep("choose") }}
    />}
    {step === "done" && result && <DoneStep result={result} />}
  </BatchPanel>
}

// The picker-sized shell for the Choose, Match and Done steps; the preview
// uses the template reader. A step's heading takes focus as the step changes,
// so focus never falls back to the page when the dropzone goes.
function BatchPanel({ title, actions, focusKey, onClose, children }: { title: string; actions?: ReactNode; focusKey: string; onClose: () => void; children: ReactNode }) {
  const closing = useContext(Closing)
  const body = useRef<HTMLDivElement>(null)
  const first = useRef(true)
  useLayoutEffect(() => {
    if (first.current) { first.current = false; return }
    body.current?.querySelector<HTMLElement>("h2[tabindex]")?.focus({ preventScroll: true })
  }, [focusKey])
  return createPortal(
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
      <div className="picker-panel template-batch" onClick={(event) => event.stopPropagation()}>
        <ModalHeader title={title} onClose={onClose} />
        <div className="picker-body template-batch-body" ref={body}>{children}</div>
        {actions && <div className="template-batch-actions">{actions}</div>}
      </div>
    </div>,
    document.body,
  )
}

function ChooseStep({ template, matters, matterId, onMatter, reading, error, onFiles }: {
  template: Template
  matters?: Matter[]
  matterId: string
  onMatter: (id: string) => void
  reading: boolean
  error?: string
  onFiles: (files: File[]) => void
}) {
  const id = useId()
  if (matters && matters.length === 0) return <p className="muted">No matters yet. Create one from the sidebar first.</p>
  return <>
    <div className="template-batch-field">
      <label htmlFor={id}>Matter</label>
      <select id={id} value={matterId} disabled={!matters || reading} onChange={(event) => onMatter(event.target.value)}>
        {!matters && <option value="">Loading matters…</option>}
        {matters?.map((matter) => <option key={matter.id} value={matter.id}>{matter.reference ? `${matter.reference} — ` : ""}{matter.title}</option>)}
      </select>
    </div>
    <div className="template-batch-drop">
      <Dropzone
        accept=".csv,.xlsx,.xls"
        busy={reading}
        onFiles={onFiles}
        title={reading ? "Reading spreadsheet…" : "Drop a spreadsheet"}
        hint="Excel or CSV. Each row becomes one document."
      />
      <a className="template-batch-link" href={templateBatchSampleUrl(template.name)} download>Download a spreadsheet for this template</a>
    </div>
    {error && <p className="template-batch-error" role="alert">{error}</p>}
  </>
}

function MatchStep({ sheet, mapping, manualKeys, onMap, onRestart }: {
  sheet: BatchSheet
  mapping: BatchMapping
  manualKeys: string[]
  onMap: (key: string, column: string | null) => void
  onRestart: () => void
}) {
  const id = useId()
  const asked = sheet.fields.filter((field) => manualKeys.includes(field.key))
  const matched = sheet.fields.filter((field) => !manualKeys.includes(field.key) && mapping[field.key])
  const dropped = sheet.dropped.blank + sheet.dropped.header
  return <>
    <div className="template-batch-intro">
      <h2 tabIndex={-1}>Match {plural(asked.length, "field")}</h2>
      <p>Choose the spreadsheet column for each field. A field left for you stays highlighted in every document.</p>
    </div>
    <ul className="template-batch-match">
      {asked.map((field, at) => <li key={field.key}>
        <label htmlFor={`${id}-${at}`}>{field.label}</label>
        <select id={`${id}-${at}`} value={mapping[field.key] ?? ""} onChange={(event) => onMap(field.key, event.target.value || null)}>
          <option value="">Leave for me to fill</option>
          {sheet.columns.map((column) => <option key={column} value={column}>{column}</option>)}
        </select>
      </li>)}
    </ul>
    <p className="template-batch-meta">
      {matched.length > 0 && <>
        <Tooltip label={<span className="template-batch-pairs">{matched.map((field) => <span key={field.key}>{field.label} ← {mapping[field.key]}</span>)}</span>}>
          <span className="template-batch-quiet" tabIndex={0}>{plural(matched.length, "field")} matched</span>
        </Tooltip>
        <span aria-hidden="true">·</span>
      </>}
      {dropped > 0 ? <Tooltip label={`${plural(dropped, "row")} ignored: blank rows and repeated headings.`}>
        <span className="template-batch-quiet" tabIndex={0}>{plural(sheet.rows.length, "row")} in {sheet.source}</span>
      </Tooltip> : <span>{plural(sheet.rows.length, "row")} in {sheet.source}</span>}
      <span aria-hidden="true">·</span>
      <button type="button" className="linklike" onClick={onRestart}>Use another spreadsheet</button>
    </p>
  </>
}

// The preview's side panel: which row is shown, whether it will be created,
// and what it leaves to fill. Missing fields stay highlighted in the document.
function PreviewRow({ position, total, sheetRow, title, missing, skipped, onStep, onSkip, onChangeMatching }: {
  position: number
  total: number
  sheetRow: number
  title: string
  missing: string[]
  skipped: boolean
  onStep: (delta: number) => void
  onSkip: () => void
  onChangeMatching?: () => void
}) {
  const id = useId()
  return <div className="template-batch-row">
    <div className="template-batch-stepper" role="group" aria-label="Row shown">
      <button type="button" className="template-batch-step" aria-label="Previous row" disabled={position === 0} onClick={() => onStep(-1)}>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m15 18-6-6 6-6" /></svg>
      </button>
      <span className="template-batch-position" aria-live="polite">{position + 1} of {total}</span>
      <button type="button" className="template-batch-step" aria-label="Next row" disabled={position >= total - 1} onClick={() => onStep(1)}>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m9 18 6-6-6-6" /></svg>
      </button>
    </div>
    {/* The row number sits on the name, clear of the stepper's buttons. */}
    <Tooltip label={`Spreadsheet row ${sheetRow}`}>
      <span className="template-batch-identity" tabIndex={0}>{title}</span>
    </Tooltip>
    <div className="template-batch-skip">
      <button id={id} type="button" role="switch" aria-checked={skipped} className={`switch${skipped ? " on" : ""}`} onClick={onSkip} />
      <label htmlFor={id}>Skip this row</label>
    </div>
    {skipped ? <p className="template-batch-status">This row will not be created.</p>
      : missing.length ? <p className="template-batch-status is-warn">Needs input: {missing.join(", ")}</p>
      : <p className="template-batch-status">Every field is filled.</p>}
    {onChangeMatching && <button type="button" className="linklike template-batch-change" onClick={onChangeMatching}>Change matching</button>}
  </div>
}

// A quiet count that opens a short list of rows, for the preview's summary.
function RowMenu({ label, items }: { label: string; items: { key: number; label: string; detail: string; onSelect: () => void }[] }) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLSpanElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const menuId = useId()
  useEffect(() => {
    if (!open) return
    const away = (event: PointerEvent) => { if (!root.current?.contains(event.target as Node)) setOpen(false) }
    document.addEventListener("pointerdown", away)
    return () => document.removeEventListener("pointerdown", away)
  }, [open])
  useEffect(() => { if (!items.length) setOpen(false) }, [items.length])
  return <span className="template-batch-menu" ref={root} onKeyDown={(event) => {
    if (event.key !== "Escape" || !open) return
    event.preventDefault()
    event.stopPropagation()
    setOpen(false)
    trigger.current?.focus()
  }}>
    <button ref={trigger} type="button" className="linklike" aria-expanded={open} aria-controls={menuId} onClick={() => setOpen((value) => !value)}>{label}</button>
    <div id={menuId} className="viewer-more-menu template-batch-menu-list" hidden={!open}>
      {items.map((item) => <button key={item.key} type="button" onClick={() => { setOpen(false); trigger.current?.focus(); item.onSelect() }}>
        <span>{item.label}<small>{item.detail}</small></span>
      </button>)}
    </div>
  </span>
}

function DoneStep({ result }: { result: BatchResult }) {
  const created = result.documents.length
  const toFill = result.documents.filter((document) => document.missing.length).length
  return <div className="template-batch-done">
    <span className="template-batch-done-icon" aria-hidden="true">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 6 9 17l-5-5" /></svg>
    </span>
    <h2 tabIndex={-1}>{created ? <>{plural(created, "document")} created in <strong>{result.folder}</strong></> : "No documents were created"}</h2>
    {toFill > 0 && <p>{toFill === 1 ? "1 has fields" : `${toFill} have fields`} left for you to fill, marked in the matter's documents.</p>}
    {result.failed.length > 0 && <Tooltip label={<span className="template-batch-pairs">{result.failed.map((item) => <span key={item.row}>Row {item.row}: {item.error}</span>)}</span>}>
      <span className="template-batch-quiet is-warn" tabIndex={0}>{plural(result.failed.length, "row")} could not be created</span>
    </Tooltip>}
  </div>
}
