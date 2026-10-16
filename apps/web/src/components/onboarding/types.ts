import type { listProviders } from "../../api/opencode/providers"

export type Provider = Awaited<ReturnType<typeof listProviders>>["all"][number]
