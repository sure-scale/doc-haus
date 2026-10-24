// Pure marker patterns shared by the server (render-readiness, open-item
// highlights) and the web document viewer, which colour-codes the same spans.
// No imports on purpose: the browser bundle pulls this file directly.

export const PLACEHOLDER_PATTERN =
  /\[\s*(?:(?:insert|tbd|tbc|to be determined|to be confirmed|placeholder)\b[^\]\n]*|_{2,}[^\]\n]*)\]/gi

// A model often invents its own bracket convention for counsel asides
// ("[BRACKETED COMMENT: ...]", "[VERIFY BEFORE FINAL — ...]",
// "[DRAFTED-FOR-REVIEW: ...]") where the directive verb is not at the start of
// the bracket, so PLACEHOLDER_PATTERN misses it and the aside would otherwise
// ship in the rendered document undetected. This is detection only: it makes the
// render-readiness gate SEE these brackets so it blocks and routes the body to
// the LLM finalizer, which has the judgment to unwrap, rewrite, or keep each one.
// We deliberately do NOT auto-rewrite them — a regex cannot tell a stray tag from
// a fully drafted definition or a negotiated bracketed alternative. Keyed on
// directive markers, never on a specific phrase, so it stays transfer-safe and
// does not flag legitimate bracket usage that carries no instruction ("[●]",
// "[Signature page follows]", "[Reserved]").
export const INSTRUCTION_COMMENT_PATTERN =
  /\[[^[\]\n]*\b(?:insert|inserted|tbd|tbc|to be (?:determined|confirmed|inserted|supplied|provided)|verify before final|drafted[ -]for[ -]review|bracketed comment|attorney (?:comment|note)|decision required|bracketed alternatives|counsel to (?:confirm|verify|review|determine)|see\s+r\d+|research required|analysis required|route\s+(?:it|this|them)\s+to|do\s+not\s+treat|placeholder)\b[^[\]\n]*\]/gi

/** An open-item note or a blank line, as the generated DOCX highlights them. */
export const OPEN_ITEM_PATTERN = /\[open item:[^\]\r\n]*\]|\(open item:[^)\r\n]*\)|_{3,}/gi

/** A keep-or-omit clause marker in a library template (dochaus/lib/markdown.ts). */
export const OPTIONAL_CLAUSE_PATTERN = /\[optional:[^\]\r\n]*\]/gi

/** The label dochaus/tool/redact.ts writes by default. A custom label is not detectable. */
export const REDACTION_MARKER_PATTERN = /\[REDACTED\]/g
