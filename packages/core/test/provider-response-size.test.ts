import { describe, expect, test } from "bun:test"
import {
  DEFAULT_MAX_DOWNLOAD_SIZE,
  DownloadError,
  createJsonErrorResponseHandler,
  createJsonResponseHandler,
  createStatusCodeErrorResponseHandler,
} from "@ai-sdk/provider-utils"
import { z } from "zod/v4"
import { createOpenAICompatible } from "@ai-sdk/openai-compatible"

const schema = z.object({ error: z.string() })
const handlers = [
  ["JSON success", createJsonResponseHandler(schema)],
  ["JSON error", createJsonErrorResponseHandler({ errorSchema: schema, errorToMessage: (error) => error.error })],
  ["status error", createStatusCodeErrorResponseHandler()],
] as const

describe("provider response size limits", () => {
  for (const [name, handler] of handlers) {
    test(`${name} rejects and cancels an oversized response before reading its body`, async () => {
      let reads = 0
      let cancelled = false
      const body = new ReadableStream<Uint8Array>(
        {
          pull(controller) {
            reads++
            controller.enqueue(new TextEncoder().encode('{"error":"provider response"}'))
            controller.close()
          },
          cancel() {
            cancelled = true
          },
        },
        { highWaterMark: 0 },
      )
      const response = new Response(body, {
        headers: { "content-length": String(DEFAULT_MAX_DOWNLOAD_SIZE + 1) },
      })

      await expect(
        Promise.resolve(handler({ response, url: "https://provider.example/test", requestBodyValues: {} })),
      ).rejects.toBeInstanceOf(DownloadError)
      expect(reads).toBe(0)
      expect(cancelled).toBe(true)
    })
  }
})

test("the patched compatible adapter preserves structured streaming errors", async () => {
  const error = { message: "Provider busy", code: "capacity_exceeded" }
  const provider = createOpenAICompatible({
    name: "scripted",
    baseURL: "https://provider.example",
    fetch: Object.assign(
      async () =>
        new Response(`data: ${JSON.stringify({ error })}\n\ndata: [DONE]\n\n`, {
          headers: { "content-type": "text/event-stream" },
        }),
      { preconnect() {} },
    ),
  })
  const { stream } = await provider.chatModel("scripted").doStream({
    prompt: [{ role: "user", content: [{ type: "text", text: "Hello" }] }],
  })
  const parts = await Array.fromAsync(stream)

  expect(parts.find((part) => part.type === "error")).toEqual({ type: "error", error })
})
