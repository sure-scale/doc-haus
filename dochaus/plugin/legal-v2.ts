import { observeStartup, startupObservation } from "../lib/startup-observation"
import { recordSessionAgent } from "./legal/session"
import { nativeReadEvidence } from "../lib/native-read-evidence"
import path from "node:path"
import { applyDeclaredModelSampling, nativeAgentSampling, declaredModelReasoning, compatibleThinkingRequest } from "./legal/v2-model-policy"
import { LegalPlugin } from "./legal/policy"
import { Wildcard } from "@opencode/core/util/wildcard"
import { applyLegalContextChanges } from "./legal/v2-context"
import type { Message, Part } from "./legal/policy-types"
import { Plugin } from "@opencode/plugin"
import { Tool } from "@opencode/schema/tool"
import { nativeSessionReader } from "../lib/native-session-reader"
import { isProjectedUserAdmission } from "../lib/engine-v2-projection"
import { writeLegalCompactionPrompt } from "./legal/compaction"
import { normalizeToolMetadata, registerLegalTools } from "../tool/register-v2"
import { createProviderCapture } from "./legal/provider-capture"
import { readEffectiveSettings } from "../../services/engine/src/effective-settings"
import { describeError } from "../lib/trace-sink"

// V2 owns execution. Legal hooks retain matter isolation and tool policy at the
// native context and tool boundaries.
export default Plugin.define({
  id: "dochaus.legal",
  setup: async (context) => {
    const directory = context.location.directory
    startupObservation("legal.setup", "enter", directory)
    // The engine already refused to start on an unusable trace root; a failure
    // specific to this matter is reported loudly without blocking its sessions.
    const providerCapture = await createProviderCapture({ directory, configDirectory: process.env.OPENCODE_CONFIG_DIR ?? path.resolve(import.meta.dir, "..") })
      .catch((error: unknown) => {
        console.error(`[drafting-trace] TRACE CAPTURE DISABLED for matter ${directory}: ${describeError(error)}`)
        return undefined
      })
    // Observation-only correlation: the native runner prepares one primary
    // provider turn per Session at a time, so its HTTP request belongs to the
    // assistant message most recently prepared for that Session.
    const preparedAssistantMessages = new Map<string, string>()
    await observeStartup("legal.tools", directory, () => registerLegalTools(context))

    const readNative = (await observeStartup("legal.reader", directory, () => nativeSessionReader(directory, context.session))).read
    const policy = await observeStartup("legal.policy", directory, () => LegalPlugin({ directory }))

    await context.session.hook("context", async event => {
      if (!event.messageID) throw new Error("Engine request identity is unavailable.")
      preparedAssistantMessages.set(event.sessionID, event.messageID)
      const current = await readNative(event.sessionID)
      const history = current.messages as unknown as Array<{ info: Message; parts: Part[] }>
      const latestUser = history.findLast(isProjectedUserAdmission)
      if (latestUser && !event.messages.some(message => message.id === latestUser.info.id) &&
        !history.slice(history.indexOf(latestUser) + 1).some(message => message.info.role === "assistant" && message.info.summary))
        throw new Error("The current durable input is absent without a durable compaction checkpoint")
      const before = structuredClone(history)
      const models = await context.model.list()
      const selectedModel = models.data.find(model => model.providerID === event.model.providerID && model.id === event.model.id)
      const model = selectedModel && { ...selectedModel, capabilities: { ...selectedModel.capabilities, reasoning: declaredModelReasoning(readEffectiveSettings(), selectedModel) } }
      if (!model) throw new Error("The selected generation model is unavailable.")
      const agents = (await context.agent.list()).data
      Object.assign(event.options, nativeAgentSampling(agents.find(agent => agent.id === event.agent)?.request?.body))
      const params = {
        options: { ...event.options },
        maxOutputTokens: event.options.maxTokens,
        temperature: event.options.temperature,
        topP: event.options.topP,
        topK: event.options.topK,
      }
      await policy["chat.params"]?.({ sessionID: event.sessionID, agent: event.agent, model } as never, params as never)
      Object.assign(event.options, params.options, {
        maxTokens: params.maxOutputTokens,
        temperature: params.temperature,
        topP: params.topP,
        topK: params.topK,
      })
      applyDeclaredModelSampling(readEffectiveSettings(), model, event.options)
      await policy["experimental.chat.messages.transform"]?.({ sessionID: event.sessionID, messageID: event.messageID, agent: event.agent, purpose: "generation" } as never, { messages: history })
      applyLegalContextChanges(event, before, history)
      const user = history.findLast(isProjectedUserAdmission)
      if (user?.info.role === "user")
        for (const [name, enabled] of Object.entries(user.info.tools ?? {})) if (!enabled) delete event.tools[name]
      if (event.tools.task) {
        const selected = agents.find(agent => agent.id === event.agent)
        const available = agents.filter(agent => agent.mode !== "primary" && !agent.hidden && selected &&
          selected.permissions.findLast(rule => Wildcard.match("subagent", rule.action) && Wildcard.match(agent.id, rule.resource))?.effect !== "deny")
        event.tools.task.description += "\n\nAvailable subagents:\n" + available.map(agent => `- ${agent.id}: ${agent.description ?? agent.name}`).join("\n")
      }
      const system = { system: [] as string[] }
      await policy["experimental.chat.system.transform"]?.({ sessionID: event.sessionID, model } as never, system)
      event.system.push(...system.system.map(text => ({ type: "text" as const, text })))
    })

    await context.session.hook("http.request", async event => {
      const model = (await context.model.list()).data.find(model => model.providerID === event.model.providerID && model.id === event.model.id)
      if (!model) throw new Error("The selected generation model is unavailable.")
      event.request = await compatibleThinkingRequest(event.request, {
        directory, sessionID: event.sessionID, agent: event.agent, model, kind: event.kind,
        reasoning: declaredModelReasoning(readEffectiveSettings(), model),
      })
      providerCapture?.request(event, event.kind === "primary" ? preparedAssistantMessages.get(event.sessionID) : undefined)
    })
    if (providerCapture) await context.session.hook("http.response", event => {
      providerCapture.response(event, event.kind === "primary" ? preparedAssistantMessages.get(event.sessionID) : undefined)
    })

    await context.session.hook("compaction", (event) => {
      const prompt = { context: [] as string[], prompt: "" }
      writeLegalCompactionPrompt(directory, prompt)
      event.system = [{ type: "text", text: prompt.prompt }]
      event.prompt = prompt.prompt
    })

    await context.tool.hook("execute.before", async event => {
      recordSessionAgent(event.sessionID, event.agent)
      const args = typeof event.input === "object" && event.input !== null ? event.input as Record<string, unknown> : {}
      if (Object.keys(args).some(key => key.startsWith("__")))
        throw new Error("Tool input cannot supply runtime-private preparation fields")
      if (event.tool === "read" && typeof args.path === "string") args.filePath = args.path
      const output = { args }
      await policy["tool.execute.before"]?.({ tool: event.tool, sessionID: event.sessionID, callID: event.id, messageID: event.messageID }, output)
      event.input = output.args
    })
    // Hooks run in registration order: this observer sees the tool's own result
    // before legal after-hooks rewrite it; the one below sees the final result.
    if (providerCapture) await context.tool.hook("execute.after", event => { providerCapture.tool("raw", event) })
    await context.tool.hook("execute.after", async event => {
      recordSessionAgent(event.sessionID, event.agent)
      const input = { tool: event.tool, sessionID: event.sessionID, callID: event.id, args: event.input }
      if (event.status === "error") {
        const output = { error: event.error.message, metadata: { ...event.error.metadata } }
        await policy["tool.execute.error"]?.({ ...input, args: event.input as Record<string, unknown> }, output)
        event.error = new Tool.Error({ ...event.error, message: output.error, metadata: output.metadata })
        return
      }
      const output = { title: String(event.result.metadata?.title ?? event.tool), metadata: { ...event.result.metadata }, output: typeof event.result.content === "string" ? event.result.content : event.result.content?.filter(part => part.type === "text").map(part => part.text).join("\n") ?? "" }
      const originalText = output.output
      if (event.tool === "read") {
        const read = nativeReadEvidence(event.input as Record<string, unknown>, output.output, output.metadata)
        input.args = read.input
        output.output = read.output
      }
      await policy["tool.execute.after"]?.(input as never, output)
      event.result = {
        ...event.result,
        content: originalText === output.output ? event.result.content :
          typeof event.result.content === "string" ? output.output :
            [{ type: "text", text: output.output }, ...(event.result.content ?? []).filter(part => part.type !== "text")],
        metadata: normalizeToolMetadata(output.metadata),
      }
    })
    if (providerCapture) await context.tool.hook("execute.after", event => { providerCapture.tool("final", event) })
    startupObservation("legal.setup", "exit", directory)
  },
})
