import { useEffect, useState } from "react"
import { Link, useParams } from "react-router-dom"
import { getMatter, type MatterDetail as Detail } from "../api/ingest"
import { listAgents, matterClient } from "../api/opencode"
import DocumentUpload from "../components/DocumentUpload"
import DocumentViewer from "../components/DocumentViewer"
import ChatPanel from "../components/ChatPanel"
import AgentPanel from "../components/AgentPanel"
import ModelSelector from "../components/ModelSelector"

type Agent = { name: string; description?: string; mode?: string }

export default function MatterDetail() {
  const { id } = useParams<{ id: string }>()
  const [matter, setMatter] = useState<Detail>()
  const [agents, setAgents] = useState<Agent[]>([])
  const [agent, setAgent] = useState("qa")
  const [viewing, setViewing] = useState<string>()

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
      <div className="row" style={{ justifyContent: "space-between", marginBottom: 16 }}>
        <div>
          <Link to="/" className="muted">
            &larr; All matters
          </Link>
          <h2 style={{ margin: "4px 0 0" }}>{matter.title}</h2>
        </div>
        <ModelSelector agents={agents} value={agent} onChange={setAgent} />
      </div>

      <DocumentUpload matterId={matter.id} documents={matter.documents} onUploaded={refresh} onView={setViewing} />

      <div className="grid">
        <ChatPanel directory={matter.dir} agent={agent} />
        <AgentPanel directory={matter.dir} />
      </div>

      {viewing && <DocumentViewer matterId={matter.id} name={viewing} onClose={() => setViewing(undefined)} />}
    </>
  )
}
