import { RequestExecutor } from "@opencode/ai/route"
import { requestExecutor } from "@opencode/core/effect/app-node-platform"
import { Context, Effect, Layer } from "effect"
import { FetchHttpClient } from "effect/unstable/http"

// Quiet model turns can exceed Bun's socket idle timeout. Decorate only the
// native model executor; the shared HTTP client and caller cancellation survive.
export const modelHttp = requestExecutor.replace(requestExecutor.mapLayer(layer => layer.pipe(
  Layer.flatMap(context => {
    const executor = Context.get(context, RequestExecutor.Service)
    return Layer.succeed(RequestExecutor.Service, {
      execute: (request, middleware) => Effect.contextWith(context => {
        const options = { ...Context.getOrUndefined(context, FetchHttpClient.RequestInit), timeout: false }
        return executor.execute(request, middleware).pipe(Effect.provideService(FetchHttpClient.RequestInit, options))
      }),
    })
  }),
)))
