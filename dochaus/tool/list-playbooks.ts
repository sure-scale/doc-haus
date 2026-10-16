import { tool } from "./legacy-definition"

// doc.haus list-playbooks tool. Returns the firm's playbook library — the
// repo-shipped playbook skills (read-only) plus the imported ones in
// WORKSPACE_ROOT/.playbooks — with full bodies, version/date metadata, and the
// matters each playbook is currently bound to. Playbooks are skills named
// playbook-*, but they have their own library and lifecycle; list-skills never
// shows them.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "List the firm's playbooks (name, description, version and last-reviewed date when declared, full markdown body, which matters each is bound to, and whether each is repo-shipped and read-only). Always call this first before creating, updating, or deleting a playbook — names must not collide, repo-shipped playbooks cannot be modified, and a playbook bound to a matter cannot be deleted.",
  args: {},
  async execute() {
    const playbooks = (await (await fetch(`${ingestUrl}/playbooks`)).json()) as {
      name: string
      description: string
      content: string
      version: string
      last_reviewed: string
      builtin: boolean
      updated_at: number
      matters: string[]
    }[]
    const lines =
      playbooks.length === 0
        ? "No playbooks yet."
        : playbooks
            .map((p) =>
              [
                `[${p.name}]${p.builtin ? " (repo-shipped, read-only)" : ""} ${p.description}`,
                [
                  p.version && `version ${p.version}`,
                  p.last_reviewed && `last reviewed ${p.last_reviewed}`,
                  `updated ${new Date(p.updated_at).toISOString().slice(0, 10)}`,
                  p.matters.length ? `bound to: ${p.matters.join(", ")}` : "not bound to any matter",
                ]
                  .filter(Boolean)
                  .join(" | "),
                `---\n${p.content}\n---`,
              ].join("\n"),
            )
            .join("\n\n")
    return {
      title: `${playbooks.length} playbook(s)`,
      output: lines,
      metadata: { playbooks: playbooks.map((p) => p.name) },
    }
  },
})
