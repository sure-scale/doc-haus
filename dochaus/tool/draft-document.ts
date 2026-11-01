import { tool } from "./legacy-definition"
import { existsSync } from "node:fs"
import path from "node:path"
import { renderMarkdownToDocx, uploadDeliverable } from "../lib/render"
import { fillTemplate, TemplateContentError } from "../lib/template-fill"
import { normalizePreferences } from "../lib/drafting-preferences"
import { pageFurniture } from "../lib/page-furniture"
import { readFirmPreferences } from "../lib/preferences"
import { safeFileName } from "../lib/safe-file-name"

// doc.haus draft-document tool. Creates a NEW Word (.docx) document in the
// current matter — either by filling a template from dochaus/templates, or from
// scratch by composing markdown onto the styled blank seed (_base.docx). Docxodus
// can edit a .docx but not create one from nothing, so every draft starts from
// seed bytes that carry the house style set (Title, Heading 1-3), which is what
// makes markdown headings land as real Word styles.
//
// The finished bytes are handed to the ingest service's deliverable route, which
// writes and registers the .docx for the web viewer without treating generated
// text as source evidence for retrieval.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    'Create a new Word (.docx) document in this matter. Two modes: pass "template" (a name from list-templates) plus "fills" — and "omit" for optional components to leave out — to draft from a template, or pass "content" (the full document body as markdown — headings, paragraphs, lists, one blank line between blocks) to draft from scratch. Template content that conflicts with the requested document, controlling sources, applicable authority, or firm instructions should be rewritten in this one call via "replaces". Call this tool at most ONCE per requested document: later changes go through the editing tools on that document, never a second draft-document call producing another version. Returns any placeholders still unfilled so they can be completed with the editing tools.',
  args: {
    name: tool.schema
      .string()
      .describe('File name for the new document, e.g. "requested-document.docx" (".docx" is appended if missing)'),
    template: tool.schema.string().optional().describe('Template file name from list-templates, e.g. "document-template.docx"'),
    fills: tool.schema
      .array(
        tool.schema.object({
          placeholder: tool.schema
            .string()
            .describe('Exact placeholder text from list-templates, e.g. "[insert required value]"'),
          value: tool.schema
            .string()
            .describe(
              "Replacement text (replaces the whole bracketed placeholder). Must be final document text: never put a bracketed placeholder inside a value — a part you cannot fill stays as its own unfilled placeholder.",
            ),
        }),
      )
      .optional()
      .describe("For template mode: the placeholder values to fill in"),
    omit: tool.schema
      .array(tool.schema.string())
      .optional()
      .describe(
        'For template mode: optional component names to leave OUT of the draft, exactly as listed by list-templates. Each named component is removed whole; every other optional component is kept with its marker stripped.',
      ),
    replaces: tool.schema
      .array(
        tool.schema.object({
          clause: tool.schema
            .string()
            .describe("Legacy schema key holding distinctive text from the template component to rewrite. Used to locate the paragraph."),
          replacement: tool.schema
            .string()
            .describe("The new component text. Replaces the whole located paragraph; markdown is supported."),
        }),
      )
      .optional()
      .describe(
        "For template mode: components to rewrite in the new draft as clean text. Each entry locates the paragraph containing the legacy `clause` anchor and replaces its text with `replacement`. Use when template content conflicts with the requested document, controlling sources, applicable authority, or firm instructions.",
      ),
    content: tool.schema.string().optional().describe("For from-scratch mode: the complete document body as markdown"),
  },
  async execute(args, ctx) {
    const base = path.basename(args.name.endsWith(".docx") ? args.name : `${args.name}.docx`)
    const name = `${safeFileName(base.slice(0, -".docx".length)) || "Draft"}.docx`
    const target = path.join(ctx.directory, name)
    if (existsSync(target)) return `A document named ${name} already exists in this matter. Pick a different name.`
    if (!args.template && !args.content) return 'Pass either "template" (with "fills") or "content".'

    // Template mode pulls the base bytes from the global library over HTTP (ingest
    // is the single writer over WORKSPACE_ROOT). Resolve the template before the ask
    // so an unknown-template error costs nothing.
    const templateRes = args.template
      ? await fetch(`${ingestUrl}/templates/content?name=${encodeURIComponent(args.template)}`)
      : undefined
    if (templateRes && !templateRes.ok)
      return `Unknown template: ${args.template}. Call list-templates for the available ones.`

    // Creating a document in the matter is gated on the matter owner's approval,
    // like every other tool that changes matter files. Asked before any Docxodus
    // work so a rejection costs nothing.
    await ctx.ask({
      permission: "draft-document",
      patterns: [target],
      always: [target],
      metadata: { document: name, template: args.template, mode: args.template ? "template" : "scratch" },
    })

    // From-scratch mode is the shared markdown -> seed -> docx -> ingest render
    // core; only template mode needs its own session for the
    // omit/fills/replaces machinery.
    if (!args.template && args.content) {
      // The tool knows no document form, so a from-scratch draft carries only
      // the firm's page numbers, and its clause references stay typed: it may
      // be advice citing the clauses of another document.
      const rendered = await renderMarkdownToDocx(args.content, undefined, pageFurniture("other", normalizePreferences(readFirmPreferences())))
        .catch((err: Error) => err)
      if (rendered instanceof Error) return `Draft failed: ${rendered.message}`
      const uploadError = await uploadDeliverable(
        ctx.directory,
        name,
        rendered.bytes,
        { kind: "create" },
        { actor: `agent:${ctx.agent}`, run: ctx.sessionID },
      )
      if (uploadError) return `Draft composed but ${uploadError}`
      return {
        title: `Drafted ${name}`,
        output: [
          `Created ${name} in this matter from scratch.`,
          rendered.lint.length
            ? `Consistency check: ${rendered.lint.join("; ")}. Fix these with the redline/editing tools during your read-back.`
            : "",
        ]
          .filter(Boolean)
          .join(" "),
        metadata: { document: target, lint: rendered.lint },
      }
    }

    // Template mode is the shared deterministic fill (lib/template-fill.ts),
    // the same path the spreadsheet batch uses.
    const filled = await fillTemplate(new Uint8Array(await templateRes!.arrayBuffer()), {
      fills: args.fills?.length ? Object.fromEntries(args.fills.map((f) => [f.placeholder, f.value])) : undefined,
      omit: args.omit,
      replaces: args.replaces,
      content: args.content,
    }).catch((err: unknown) => {
      if (err instanceof TemplateContentError) return err
      throw err
    })
    if (filled instanceof TemplateContentError) return `Draft failed: ${filled.message}`
    const { bytes, remaining, missingOmits, refusedOmits, missedReplaces, lint, removedReferences } = filled

    // Upload through ingest rather than writing the file directly: its
    // deliverable route writes and registers the canonical .docx without
    // admitting generated output into source retrieval.
    const uploadError = await uploadDeliverable(
      ctx.directory,
      name,
      bytes,
      { kind: "create" },
      { actor: `agent:${ctx.agent}`, run: ctx.sessionID },
    )
    if (uploadError) return `Draft composed but ${uploadError}`

    return {
      title: `Drafted ${name}`,
      output: [
        `Created ${name} in this matter${args.template ? ` from template ${args.template}` : " from scratch"}.`,
        remaining.length
          ? `Placeholders still unfilled: ${remaining.join(", ")}. Fill them with the editing tools or leave them for the lawyer.`
          : "No placeholders remain.",
        missingOmits.length
          ? `Optional components not found (left as-is): ${missingOmits.join(", ")}. Check the names against list-templates.`
          : "",
        refusedOmits.length
          ? `Optional components the document engine could not remove (left in the draft): ${refusedOmits.join(", ")}. Remove them with the editing tools.`
          : "",
        missedReplaces.length
          ? `Component rewrites whose anchor text was not found (template language kept as-is): ${missedReplaces.map((c) => JSON.stringify(c)).join(", ")}. Locate the exact text with read-document and retry, or fix the draft with the editing tools.`
          : "",
        lint.length
          ? `Consistency check: ${lint.join("; ")}. Fix these with the redline/editing tools during your read-back.`
          : "",
        removedReferences.length
          ? `${removedReferences.length} cross-reference${removedReferences.length === 1 ? "" : "s"} to parts no longer in the draft ${removedReferences.length === 1 ? "is" : "are"} marked as open items: ${[...new Set(removedReferences)].join(", ")}. Point each at the right clause or remove it with the editing tools.`
          : "",
      ]
        .filter(Boolean)
        .join(" "),
      metadata: {
        document: target,
        template: args.template,
        removedReferences,
        remainingPlaceholders: remaining,
        omitted: args.omit,
        refusedOmits,
        replaced: (args.replaces ?? []).filter((r) => !missedReplaces.includes(r.clause)).map((r) => r.clause),
        missedReplaces,
        lint,
      },
    }
  },
})
