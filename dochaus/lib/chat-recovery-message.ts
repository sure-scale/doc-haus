const IDENTITY = /^[A-Za-z0-9_-]{1,512}$/
const RECOVERY_PREFIX = "msg_recovery_"

// Recovery is an application-owned message identity, never an interpretation of
// lawyer prose. This is the exact envelope emitted by the continuation writer;
// older persisted user rows need no synthetic flag, but must retain that identity.
export function isChatRecoveryMessage(message: { id: unknown; sessionID: unknown; text: string }) {
  if (typeof message.id !== "string" || !message.id.startsWith(RECOVERY_PREFIX) ||
      typeof message.sessionID !== "string" || !IDENTITY.test(message.sessionID)) return false
  const failed = message.id.slice(RECOVERY_PREFIX.length)
  if (!IDENTITY.test(failed)) return false
  const marker = `<dochaus-recovery kind="malformed-tool" run-id="${message.sessionID}" progress-token="malformed:${failed}" />`
  return message.text === marker || message.text.startsWith(`${marker}\n`) || message.text.startsWith(`${marker}\r\n`)
}
