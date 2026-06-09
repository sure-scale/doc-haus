import { useEffect, useState } from "react"
import { useParams, useSearchParams } from "react-router-dom"
import { getMatter, type MatterDetail as Detail } from "../api/ingest"
import { listAgents, matterClient } from "../api/opencode"
import DocumentUpload from "../components/DocumentUpload"
import DocumentViewer from "../components/DocumentViewer"
import ChatPanel from "../components/ChatPanel"
import ReviewGrid from "../components/ReviewGrid"
import WorkflowLauncher from "../components/WorkflowLauncher"
import WorkflowArtifact from "../components/WorkflowArtifact"

type Agent = { name: string; description?: string; mode?: string }

// One matter, four surfaces, switched by the `view` query param from the sidebar:
// chat | review | documents | workflows. The content is a single canvas whose
// width flexes to the surface — the grid and the documents manager run full-width,
// while chat keeps a documents rail for reference. Default surface is chat.
export default function MatterDetail() {
  const { id } = useParams<{ id: string }>()
  const [params] = useSearchParams()
  const view = params.get("view") ?? "chat"
  const session = params.get("session") ?? undefined
  const [matter, setMatter] = useState<Detail>()
  const [agents, setAgents] = useState<Agent[]>([])
  const [agent, setAgent] = useState("qa")
  const [viewing, setViewing] = useState<string>()
  const [workflow, setWorkflow] = useState<string>()
  const [docsOpen, setDocsOpen] = useState(() => localStorage.getItem("dh.docs") !== "0")

  useEffect(() => {
    localStorage.setItem("dh.docs", docsOpen ? "1" : "0")
  }, [docsOpen])

  function refresh() {
    if (id) getMatter(id).then(setMatter)
  }

  useEffect(refresh, [id])

  useEffect(() => {
    if (!matter) return
    listAgents(matterClient(matter.dir)).then((list) => {
      setAgents(list as Agent[])
      if (list.some((a) => (a as Agent).name === "qa")) setAgent("qa")
    })
  }, [matter])

  if (!matter) return <p className="muted">Loading matter...</p>

  const available = new Set(agents.map((a) => a.name))

  return (
    <>
      <div style={{ marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>
          {matter.reference && <span className="matter-ref">{matter.reference}</span>}
          {matter.title}
        </h2>
      </div>

      {view === "chat" && (
        <div className={`matter-body${docsOpen ? "" : " docs-collapsed"}`}>
          <div className="workspace">
            <ChatPanel
              key={session ?? "new"}
              directory={matter.dir}
              sessionID={session}
              agent={agent}
              available={available}
              onAgentChange={setAgent}
            />
          </div>
          <DocumentUpload
            matterId={matter.id}
            documents={matter.documents}
            onUploaded={refresh}
            onView={setViewing}
            collapsed={!docsOpen}
            onToggle={() => setDocsOpen((o) => !o)}
          />
        </div>
      )}

      {view === "review" && <ReviewGrid matterId={matter.id} directory={matter.dir} documents={matter.documents} />}

      {view === "documents" && (
        <DocumentUpload matterId={matter.id} documents={matter.documents} onUploaded={refresh} onView={setViewing} />
      )}

      {view === "workflows" && (
        <div className={`workspace${workflow ? " with-artifact" : ""}`}>
          <div className="card">
            <h2 style={{ marginTop: 0 }}>Workflows</h2>
            <p className="muted" style={{ marginBottom: 12 }}>
              Multi-step routines that coordinate several reviewers across this matter's documents and return one report.
            </p>
            <WorkflowLauncher available={available} onLaunch={setWorkflow} />
          </div>
          {workflow && (
            <WorkflowArtifact key={workflow} directory={matter.dir} workflow={workflow} onClose={() => setWorkflow(undefined)} />
          )}
        </div>
      )}

      {viewing && <DocumentViewer matterId={matter.id} name={viewing} onClose={() => setViewing(undefined)} />}
    </>
  )
}
