import { modelConcurrency, readEffectiveSettings } from "../../../services/engine/src/effective-settings"

/** Size each workflow pool from the installation setting. The engine's shared
 * request gate enforces the actual limit across all pools and sessions. */
export function workflowConcurrency() {
  return modelConcurrency(readEffectiveSettings())
}
