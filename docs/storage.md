# Storage and backup

DOCHAUS_DATA_ROOT selects a common root: workspace/ contains matters and
installation preferences; runtime/ contains native credentials, sessions,
configuration, state and cache. WORKSPACE_ROOT and DOCHAUS_RUNTIME_DIR override
the corresponding path explicitly. Without the data-root option, legacy local
workspace and .dochaus-runtime paths remain in use. Selecting a new root does
not migrate existing data.

Mutable model/provider settings and agent/skill toggles are stored at
<runtime>/config/opencode/opencode.json. Setup progress is stored beside that
configuration. Ingest validates incremental changes and publishes serialized,
restricted-permission atomic replacements. The shipped dochaus/opencode.json
is product configuration, not an installation preference file.

Research keys and drafting preferences live under workspace/.preferences/.
Native provider credentials and durable sessions use runtime storage. Matters
also contain documents, legal.db databases, workflow state and working files.
Some UI preferences remain in browser storage; server backup does not capture
browser-only choices.

Reusable templates live in workspace/.templates. A source-template conversation
keeps its captured Word source, proposed copy and review metadata in its own
workspace/.template-drafts directory; its native conversation uses runtime
storage. Unsent answers to questions in library conversations are browser-only.
Include both workspace and runtime in a backup to retain the source copy and
the conversation that prepared it.

Custom agent/workflow definitions currently use ignored files under the source
checkout: dochaus/agents.json, dochaus/workflows.json and generated Markdown under
dochaus/agent/. Builder conversations remain in workspace/.agents and .workflows.
Back up the custom registry files and custom Markdown as well; when restoring,
review custom definitions separately so shipped legal agents are not overwritten.
These generated definitions are excluded from Git.

Back up workspace and runtime together while services are stopped, or use a
consistent live backup procedure that includes SQLite WAL companions. Include
externalized libraries, configuration and explicitly configured capture paths.
A SQLite-only backup cannot recover documents and file-backed workflow state.
Treat backups as confidential. Keep a recovery copy before changing layouts.

Docker preserves the existing dochaus-workspace volume at /data/workspace and
adds dochaus-runtime at /data/runtime. Container recreation retains both.
The current Compose volumes do not persist the custom builder definition files
under dochaus/. Export/back up those definitions before recreating a container;
custom builder definitions require a separate persistence arrangement. Model
settings, toggles, native credentials, sessions and matter files use the volumes
described above.
Removing volumes deletes their data; no automatic rename or deletion is provided.
