import type { listProviders } from "../../api/opencode/providers"

export type Provider = Awaited<ReturnType<typeof listProviders>>["all"][number]
export type Methods = Record<string, { type: "oauth" | "api"; label: string }[]>

export type Tab = "models" | "providers" | "drafting" | "research" | "data" | "matters" | "approvals" | "appearance"
