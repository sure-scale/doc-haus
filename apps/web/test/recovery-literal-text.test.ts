import { expect, test } from "bun:test"
import type { SessionMessageInfo } from "@opencode/client"
import { projectMessages } from "../src/api/session-view"
import { isDraftingRecoveryMessage, lastVisibleUserMessage } from "../src/components/chat/drafting-recovery"
import { malformedToolRecoveryCount } from "../src/components/chat/turn-recovery"
import { toTurns } from "../src/components/chat/turn-model"

const marker = '<dochaus-recovery kind="malformed-tool" run-id="ses_chat" progress-token="malformed:msg_failed" />'
function project(id: string, text: string) {
  return projectMessages("ses_chat", [{ id, type: "user", text, time: { created: 1 } }] as SessionMessageInfo[])
}

test("quoted recovery syntax stays visible across native projection and persisted reload", () => {
  for (const text of [marker, `Explain \`${marker}\` in the source.`, `\`\`\`xml\n${marker}\n\`\`\``]) {
    const history = project("msg_real_user", text)
    expect(isDraftingRecoveryMessage(history[0]!)).toBe(false)
    expect(toTurns(history, "")[0]?.text).toBe(text)
    expect(lastVisibleUserMessage(history)?.info.id).toBe("msg_real_user")
    expect(malformedToolRecoveryCount(history)).toBe(0)
    expect(toTurns(JSON.parse(JSON.stringify(history)), "")).toEqual(toTurns(history, ""))
  }
})

test("real persisted recovery remains hidden and counted without a synthetic flag", () => {
  const history = [...project("msg_real_user", "Continue the request."), ...project("msg_recovery_msg_failed", marker + "\nContinue.")]
  expect(isDraftingRecoveryMessage(history[1]!)).toBe(true)
  expect(toTurns(history, "").map(turn => turn.text)).toEqual(["Continue the request."])
  expect(malformedToolRecoveryCount(history)).toBe(1)
  expect(malformedToolRecoveryCount(JSON.parse(JSON.stringify(history)))).toBe(1)
})

test("a real user quoting a marker starts a new recovery denominator", () => {
  const history = [...project("msg_recovery_msg_failed", marker), ...project("msg_real_user", marker)]
  expect(malformedToolRecoveryCount(history)).toBe(0)
  expect(toTurns(history, "")[0]?.id).toBe("msg_real_user")
})
