import { z } from "zod"

export interface LegalToolContext {
  directory: string
  worktree: string
  agent: string
  sessionID: string
  messageID: string
  callID: string
  abort: AbortSignal
  metadata(input: { title?: string; metadata?: Record<string, unknown> }): void
  ask(input: {
    permission: string
    patterns: string[]
    always?: string[]
    metadata?: Record<string, unknown>
  }): Promise<void>
}

export type LegalToolResult = string | {
  title?: string
  output: string
  metadata?: Record<string, unknown>
}

// The existing legal tools share one Zod runtime with their imported schemas.
// V2 registers their definitions through its native tool.transform API.
export function tool<Args extends z.ZodRawShape, Result extends LegalToolResult | undefined>(definition: {
  description: string
  args: Args
  execute(args: z.infer<z.ZodObject<Args>>, context: LegalToolContext): Promise<Result>
}) {
  return definition
}

export namespace tool {
  export const schema = z
}
