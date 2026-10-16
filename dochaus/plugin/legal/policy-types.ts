// Transitional transcript types for the legal policy layer. These are display
// projections, never a v1 transport or execution path.
export type Message = import("@opencode-ai/sdk").Message & { agent?: string }
export type Part = import("@opencode-ai/sdk").Part
type Hooks = import("@opencode-ai/plugin").Hooks
export type LegalHooks = Omit<Hooks, "event" | "tool.execute.before" | "experimental.chat.messages.transform" | "experimental.chat.system.transform"> & {
  "tool.execute.before"?: (
    input: { tool: string; sessionID: string; callID: string; messageID?: string },
    output: { args: Record<string, unknown> },
  ) => Promise<void>
  "experimental.chat.messages.transform"?: (
    input: { sessionID?: string; messageID?: string; agent?: string; purpose?: "generation" | "compaction" },
    output: { messages: Array<{ info: Message; parts: Part[] }> },
  ) => Promise<void>
  "experimental.chat.system.transform"?: (
    input: { sessionID: string; model: import("@opencode-ai/sdk").Model },
    output: { system: string[] },
  ) => Promise<void>
  "tool.execute.error"?: (
    input: { tool: string; sessionID: string; callID: string; args: Record<string, unknown> },
    output: { error: string; metadata?: Record<string, unknown> },
  ) => Promise<void>
}
export type LegalPolicy = (input: Pick<import("@opencode-ai/plugin").PluginInput, "directory">) => Promise<LegalHooks>
