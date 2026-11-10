import type { SourceDependencyStatus } from "../../../../dochaus/lib/source-version-dependencies-contract"
import "./SourceBasisNotice.css"

/** A saved input basis is descriptive. Advancing a source does not rewrite an
 * existing memo, checklist or draft and does not approve its conclusions. */
export default function SourceBasisNotice({ basis }: { basis?: SourceDependencyStatus }) {
  if (!basis || basis.status === "current") return null
  return <div className="source-basis-notice" role="status">
    <strong>{basis.status === "stale" ? "Source material changed" : "Source basis needs checking"}</strong>
    <p>{basis.status === "stale" ? "Source material changed after this work was prepared. Review it against the latest available sources before relying on it." : "The complete source-version basis was not recorded or could not be verified."}</p>
    {basis.sources.some(source => source.status !== "current") && <ul>{basis.sources.filter(source => source.status !== "current").map(source => <li key={source.versionId}>
      {source.name}{source.versionNumber !== null ? ` · Built from version ${source.versionNumber}` : ""}
      {source.status === "newer-version" ? ` · Current version ${source.currentVersionNumber ?? "unknown"}` : source.status === "newer-received-branch" ? ` · New linked return${source.linkedReturn ? `: ${source.linkedReturn.sourceName} · Version ${source.linkedReturn.versionNumber}` : " available"}` : source.status === "removed" ? " · Source removed or redacted" : source.status === "unrecorded-change" ? " · Working file changed outside its saved version" : " · Source unavailable"}
    </li>)}</ul>}
  </div>
}
