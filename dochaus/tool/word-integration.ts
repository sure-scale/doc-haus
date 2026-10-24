import { tool } from "./legacy-definition"
import { createHash } from "node:crypto"
import { existsSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { findPastListLabel } from "../lib/list-label"
import { assertDirectWorkingDocxWriteAllowed } from "../lib/working-docx"
import { withArtifactMutationLocksAsync } from "../capabilities/artifacts/mutation-lock"

// doc.haus word-integration tool. Reads and edits the matter's canonical Word
// (.docx) document via Docxodus, so edits round-trip through Microsoft Word.
// `services/ingest` writes that canonical .docx into the matter directory; an
// edit is saved back through ingest's document revision route, which keeps the
// document's role, re-indexes a source, keeps a deliverable's paragraph pointers
// and rebuilds a saved combined file. Edits here are plain (non-tracked);
// tracked changes are a separate tool.

export default tool({
  description:
    'Read or edit the matter\'s canonical Word (.docx) document in place. Use action "read" to get the document text, or "replace" to change wording. Edits are plain (not tracked changes) and re-open correctly in Microsoft Word. A list number Word prints before a paragraph ("2.1") is not text; leave it out of find and replace.',
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    action: tool.schema.enum(["read", "replace"]).describe('"read" the document text, or "replace" text in place'),
    find: tool.schema.string().optional().describe("For replace: the exact text to find"),
    replace: tool.schema.string().optional().describe("For replace: the replacement text"),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    if (!existsSync(file)) return `Document not found in this matter: ${args.document}`
    const name = path.basename(file)

    if (args.action === "read") {
      const session = (await docxodus()).openDocxSession(await Bun.file(file).bytes(),{})
      try {return {title:`Read ${name}`,output:session.project().markdown}} finally {session.close()}
    }

    // Ingest registers documents by name in the matter folder itself, so only a
    // document there can be saved; a same-named file elsewhere would replace it.
    if (path.dirname(path.resolve(file)) !== path.resolve(ctx.directory))
      return `Only documents saved directly in this matter can be edited: ${args.document}`

    // Replace rewrites the canonical .docx with no review step, so it is gated
    // on the matter owner's approval (permission "word-integration" in
    // opencode.json). ctx.ask blocks until they reply and throws on reject —
    // asked before the Docxodus session opens so a rejection leaks nothing.
    // "Always" approves future edits to this document only.
    if (!args.find || args.replace === undefined) return 'Replace requires both "find" and "replace".'
    await ctx.ask({
      permission: "word-integration",
      patterns: [file],
      always: [file],
      metadata: { document: name, find: args.find, replace: args.replace },
    })

    // The edit is made under the document's lock, which is released before the
    // save: ingest is another process and takes the same lock for the revision.
    const edit = await withArtifactMutationLocksAsync(ctx.directory,[file],async () => {
      assertDirectWorkingDocxWriteAllowed(ctx.directory,file)
      const original = await Bun.file(file).bytes()
      const session = (await docxodus()).openDocxSession(original,{})
      try {
        const {find,replace,targets} = findPastListLabel(args.find!,args.replace!,(text) => session.findAllByText(text),(id) => session.getAnchorInfo(id)?.visibleText)
        if (!targets.length) return `Text not found in ${name}: ${JSON.stringify(args.find)}`
        const results = targets.flatMap((target) => session.replaceTextRange(target.id,find,replace))
        const failed = results.find((result) => !result.success)
        if (failed) return `Edit failed: ${failed.error?.message ?? JSON.stringify(failed.error)}`
        return {bytes:session.save(),priorSha256:createHash("sha256").update(original).digest("hex"),edits:results.length,find,replace}
      } finally {session.close()}
    })
    if (typeof edit === "string") return edit

    // Saved only through ingest; the file is never written here. The prior hash
    // lets ingest refuse a document that changed after it was read.
    const form = new FormData()
    form.append("file", new File([new Uint8Array(edit.bytes)], name))
    const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"
    const refusal = await fetch(`${ingestUrl}/matters/${encodeURIComponent(path.basename(ctx.directory))}/documents/revisions`, {
      method: "POST",
      headers: { "x-document-prior-sha256": edit.priorSha256 },
      body: form,
      signal: ctx.abort,
    }).then(
      async (response) => {
        if (response.ok) return
        const body = await response.json().catch(() => undefined) as { error?: unknown } | undefined
        const reason = typeof body?.error === "string" ? body.error.replace(/\.$/, "") : response.statusText
        return `Edit not saved: ingest refused the revision (${response.status}${reason ? `: ${reason}` : ""}). ${name} was not changed by this edit.`
      },
      // A request that fails after it was sent may still have been saved.
      (error: unknown) => `Edit not confirmed: the ingest service could not be reached (${error instanceof Error ? error.message : String(error)}). Read ${name} again before retrying.`,
    )
    if (refusal) return refusal
    return {title:`Edited ${name}`,
      output:`Replaced ${edit.edits} occurrence(s) of ${JSON.stringify(edit.find)} with ${JSON.stringify(edit.replace)} in ${name}.`,
      metadata:{document:file,find:edit.find,replace:edit.replace,edits:edit.edits}}
  },
})
