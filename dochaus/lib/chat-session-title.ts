export type ChatTitleTransport = {
  get(sessionID: string): Promise<{ title?: string }>
  rename(sessionID: string, title: string, expectedTitle?: string): Promise<void>
}

/** Delegate naming to the configured native title agent. A scoped host can
 * name its clean user-input child, then display that title on its conversation
 * container without fabricating a user message in the container. */
export async function generateConversationTitle(input: {
  transport: ChatTitleTransport
  sessionID: string
  targetSessionID?: string
}): Promise<string | undefined> {
  const target = input.targetSessionID ?? input.sessionID
  const original = target !== input.sessionID ? await input.transport.get(target) : undefined
  const childOriginal = original ? await input.transport.get(input.sessionID) : undefined
  await input.transport.rename(input.sessionID, "")
  const generated = (await input.transport.get(input.sessionID)).title
  if (!generated?.trim() || /[\r\n\x00-\x1f]/.test(generated)) return
  if (childOriginal && generated === childOriginal.title) return
  if (original && original.title !== undefined)
    await input.transport.rename(target, generated, original.title)
  return generated
}
