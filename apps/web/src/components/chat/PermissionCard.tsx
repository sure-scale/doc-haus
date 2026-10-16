import type { PermissionReply, PermissionRequest } from "../../api/opencode/permissions"
import { humanizeTool, basename } from "./step-labels"

// What each gated tool is asking to do, phrased for the approval card's title.
const PERMISSION_VERBS: Record<string, string> = {
  "word-integration": "edit",
  "tracked-changes": "propose a tracked change in",
  redline: "propose a redline in",
  "draft-document": "create",
  "create-template": "create the template",
  redact: "permanently remove content from",
}

// One parked edit-tool call awaiting the user's decision. The metadata the tool
// attached (find/replace or clause/replacement) previews the change in the same
// red/green the redline cards use. "Always allow" saves the action/resources
// supplied by the tool in the engine's durable project permissions.
export default function PermissionCard({
  request,
  onReply,
}: {
  request: PermissionRequest
  onReply: (id: string, reply: PermissionReply) => void
}) {
  const meta = request.metadata
  const document = typeof meta.document === "string" ? meta.document : request.patterns[0] && basename(request.patterns[0])
  const oldText = (meta.find ?? meta.clause ?? meta.text) as string | undefined
  const newText = (meta.replace ?? meta.replacement ?? meta.label) as string | undefined
  const verb = PERMISSION_VERBS[request.permission] ?? `use ${humanizeTool(request.permission)} on`
  const documentAllowance =
    request.always.length === 1 && request.patterns.length === 1 &&
    request.always[0] === request.patterns[0] && typeof meta.document === "string" &&
    basename(request.always[0]) === basename(meta.document) && !/[?*\[\]{}]/.test(request.always[0])
  // Redaction is irreversible, so its card spells out the scope (occurrence
  // count) and the reason that will go to the redaction log.
  const detail =
    request.permission === "redact"
      ? `Removes ${meta.occurrences ?? "all"} occurrence(s) and any metadata mentions — this cannot be undone. Reason: ${meta.reason}`
      : undefined
  return (
    <div className="permission-card">
      <div className="permission-title">
        The assistant wants to {verb} {document ?? "a document"}
      </div>
      {(oldText || newText) && (
        <div className="permission-preview">
          {oldText && <div className="redline-old">{oldText}</div>}
          {newText && <div className="redline-new">{newText}</div>}
        </div>
      )}
      {detail && <div className="permission-title">{detail}</div>}
      {request.always.length > 0 && !documentAllowance && (
        <div className="permission-title">A saved allowance for this action covers: {request.always.join(", ")}</div>
      )}
      <div className="permission-actions">
        <button className="primary" onClick={() => onReply(request.id, "once")}>
          Allow once
        </button>
        {request.always.length > 0 && (
          <button onClick={() => onReply(request.id, "always")}>
            {documentAllowance
              ? "Always allow this action for this document"
              : `Always allow this action for ${request.always.length === 1 ? "this item" : "these items"}`}
          </button>
        )}
        <button className="permission-reject" onClick={() => onReply(request.id, "reject")}>
          Reject
        </button>
      </div>
    </div>
  )
}
