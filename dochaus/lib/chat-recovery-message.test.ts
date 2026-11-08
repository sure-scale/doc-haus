import { expect, test } from "bun:test"
import { isChatRecoveryMessage } from "./chat-recovery-message"

const message = {
  id: "msg_recovery_msg_failed",
  sessionID: "ses_chat",
  text: '<dochaus-recovery kind="malformed-tool" run-id="ses_chat" progress-token="malformed:msg_failed" />',
}

test("persisted recovery messages retain the writer's identity without a synthetic flag", () => {
  for (const suffix of ["", "\nContinue.", "\r\nContinue."])
    expect(isChatRecoveryMessage({ ...message, text: message.text + suffix })).toBe(true)
})

test("literal markers and inconsistent identities never become control messages", () => {
  for (const input of [
    { ...message, id: "msg_real_user" },
    { ...message, id: "msg_recovery_some_other_failure" },
    { ...message, sessionID: "ses_other" },
    { ...message, sessionID: undefined },
    { ...message, id: undefined },
    { ...message, text: `Explain this marker: ${message.text}` },
    { ...message, text: `\`\`\`xml\n${message.text}\n\`\`\`` },
    { ...message, text: message.text + " remains literal." },
    { ...message, text: message.text.replace("malformed:msg_failed", "malformed:wrong") },
    { ...message, text: message.text.replace(" />", " >") },
    { ...message, text: message.text.replace("malformed-tool", "unknown") },
  ]) expect(isChatRecoveryMessage(input)).toBe(false)
})
