import { SCANNED_CHARS_PER_PAGE } from "./pdf"

// Extraction quality gate (#239). After a source's text is extracted and
// normalised, a deterministic check says whether it is usable: ok, degraded
// (indexed, but the text may be incomplete or garbled, with the signals that
// tripped) or no_readable_text (the file is kept and nothing is indexed). The
// result is stored on the document (documents.extraction_quality) and shown
// wherever the document is used. It is never a prompt-injection finding.

export type QualitySignalCode =
  | "empty"
  | "no_letters_or_digits"
  | "replacement_characters"
  | "few_letters_or_digits"
  | "little_text_per_page"
  | "no_email_body"
  | "formulas_without_values"

export type QualitySignal = { code: QualitySignalCode; detail: string }

export type ExtractionQuality =
  | { status: "ok" }
  | { status: "degraded" | "no_readable_text"; signals: QualitySignal[] }

// What the extractor knows beyond the text itself.
// pdf: its page count. email: the message body (normalised like the text) and
// how many attachments it lists. workbook: how many cells hold anything, and how
// many of those are formulas saved without a calculated value.
export type ExtractionFacts =
  | { format: "docx" }
  | { format: "pdf"; pages: number }
  | { format: "email"; body: string; attachments: number }
  | { format: "workbook"; cells: number; formulasWithoutValue: number }

// The thresholds, all in one place. Each follows from what the signal means,
// the same for every document, never tuned to a document type, jurisdiction or
// test file.
export const QUALITY_THRESHOLDS = {
  // Correctly decoded text has no replacement (U+FFFD) or control residue; a
  // stray one can come from a single odd glyph. At one visible character in a
  // hundred the decoding has failed somewhere a reader would notice (degraded);
  // at half, residue outnumbers text and nothing usable is left.
  degradedResidueShare: 0.01,
  unreadableResidueShare: 0.5,
  // Of the visible characters other than ASCII punctuation and symbols (so table
  // rules, CSV commas and markup do not count), real text is mostly letters and
  // digits. Below half, the rest is symbol or glyph noise.
  degradedLetterShare: 0.5,
  // A PDF page with less text than this is the line at which ingest already
  // treats it as a scan (pdf.ts), so text this thin after OCR (or without it)
  // has missed what the pages show.
  minPdfCharactersPerPage: SCANNED_CHARS_PER_PAGE,
}

const READABLE = /[\p{L}\p{N}]/gu
// Replacement characters, control characters left after normalisation,
// unassigned code points and lone surrogates.
const RESIDUE = /[\uFFFD\p{Cc}\p{Cn}\p{Cs}]/gu
const ASCII_PUNCTUATION = /[!-/:-@[-`{-~]/gu

const count = (text: string, pattern: RegExp) => text.match(pattern)?.length ?? 0

// The status and signals for the text alone.
function textQuality(text: string): ExtractionQuality {
  const compact = text.replace(/\s+/gu, "")
  const visible = Array.from(compact).length
  if (!visible) return { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] }
  const residue = count(compact, RESIDUE)
  const residueSignal: QualitySignal = {
    code: "replacement_characters",
    detail: `${residue} of ${visible} characters are replacement or control characters`,
  }
  if (residue / visible >= QUALITY_THRESHOLDS.unreadableResidueShare) return { status: "no_readable_text", signals: [residueSignal] }
  const readable = count(text, READABLE)
  if (!readable)
    return { status: "no_readable_text", signals: [{ code: "no_letters_or_digits", detail: "The extracted text has no letters or digits" }] }
  const signals: QualitySignal[] = []
  if (residue / visible >= QUALITY_THRESHOLDS.degradedResidueShare) signals.push(residueSignal)
  const considered = visible - count(text, ASCII_PUNCTUATION)
  if (considered > 0 && readable / considered < QUALITY_THRESHOLDS.degradedLetterShare)
    signals.push({ code: "few_letters_or_digits", detail: `Only ${readable} of ${considered} text characters are letters or digits` })
  return signals.length ? { status: "degraded", signals } : { status: "ok" }
}

function result(signals: QualitySignal[], unreadable: boolean): ExtractionQuality {
  if (unreadable) return { status: "no_readable_text", signals }
  return signals.length ? { status: "degraded", signals } : { status: "ok" }
}

function signalsOf(quality: ExtractionQuality) {
  return quality.status === "ok" ? [] : quality.signals
}

// The quality of a source's normalised text.
export function assessExtraction(text: string, facts: ExtractionFacts): ExtractionQuality {
  // An email's headers are readable metadata; its content is the body. A body
  // that is missing or unreadable leaves the email indexed by its headers and
  // marked degraded, unless the email only carries attachments.
  if (facts.format === "email") {
    const body = textQuality(facts.body)
    if (body.status !== "no_readable_text") return body
    const empty = body.signals[0]?.code === "empty"
    if (empty && facts.attachments > 0) return { status: "ok" }
    const headers = textQuality(text.endsWith(facts.body) ? text.slice(0, text.length - facts.body.length) : text)
    if (headers.status === "no_readable_text") return headers
    return {
      status: "degraded",
      signals: [{
        code: "no_email_body",
        detail: empty ? "The email has no message body; only its headers were indexed" : `The message body has no readable text (${body.signals[0]!.detail.toLowerCase()}); only its headers were indexed`,
      }],
    }
  }
  const base = textQuality(text)
  if (base.status === "no_readable_text") return base
  const signals = signalsOf(base)
  if (facts.format === "pdf" && facts.pages > 0) {
    const visible = Array.from(text.replace(/\s+/gu, "")).length
    if (visible / facts.pages < QUALITY_THRESHOLDS.minPdfCharactersPerPage)
      signals.push({
        code: "little_text_per_page",
        detail: `${visible} characters over ${facts.pages} page${facts.pages === 1 ? "" : "s"}; pages may be images whose text was not read`,
      })
  }
  if (facts.format === "workbook" && facts.formulasWithoutValue > 0) {
    const all = facts.formulasWithoutValue === facts.cells
    signals.push({
      code: "formulas_without_values",
      detail: all
        ? "Every filled cell is a formula saved without a calculated value"
        : `${facts.formulasWithoutValue} of ${facts.cells} filled cells are formulas saved without a calculated value`,
    })
    // With no value anywhere, the formulas' own text is all that is left.
    return result(signals, all)
  }
  return result(signals, false)
}

// The stored form (documents.extraction_quality) read back; undefined for a
// document ingested before the gate, which shows nothing.
export function readExtractionQuality(value: string | null | undefined): ExtractionQuality | undefined {
  if (!value) return
  return JSON.parse(value) as ExtractionQuality
}
