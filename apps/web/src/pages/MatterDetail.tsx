import { useEffect, useState } from "react"
import { Link, useParams, useSearchParams } from "react-router-dom"
import { getMatter, type MatterDetail as Detail } from "../api/ingest"
import { listAgents, matterClient } from "../api/opencode"
import DocumentUpload from "../components/DocumentUpload"
import DocumentViewer from "../components/DocumentViewer"
import ChatPanel from "../components/ChatPanel"
import WorkflowArtifact from "../components/WorkflowArtifact"

type Agent = { name: string; description?: string; mode?: string }

export default function MatterDetail() {
  const { id } = useParams<{ id: string }>()
  const [params] = useSearchParams()
  const session = params.get("session") ?? undefined
  const [matter, setMatter] = useState<Detail>()
  const [agents, setAgents] = useState<Agent[]>([])
  const [agent, setAgent] = useState("qa")
  const [viewing, setViewing] = useState<string>()
  const [workflow, setWorkflow] = useState<string>()

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

  return (
    <>
      <div style={{ marginBottom: 16 }}>
        <Link to="/" className="muted">
          &larr; All matters
        </Link>
        <h2 style={{ margin: "4px 0 0" }}>
          {matter.reference && <span className="matter-ref">{matter.reference}</span>}
          {matter.title}
        </h2>
      </div>

      <DocumentUpload matterId={matter.id} documents={matter.documents} onUploaded={refresh} onView={setViewing} />

      <div className={`workspace${workflow ? " with-artifact" : ""}`}>
        <ChatPanel
          key={session ?? "new"}
          directory={matter.dir}
          sessionID={session}
          agent={agent}
          available={new Set(agents.map((a) => a.name))}
          onAgentChange={setAgent}
          onLaunchWorkflow={setWorkflow}
        />
        {workflow && (
          <WorkflowArtifact
            key={workflow}
            directory={matter.dir}
            workflow={workflow}
            onClose={() => setWorkflow(undefined)}
          />
        )}
      </div>

      {viewing && <DocumentViewer matterId={matter.id} name={viewing} onClose={() => setViewing(undefined)} />}
    </>
  )
}
