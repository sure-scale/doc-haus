import { existsSync } from "node:fs"
import os from "node:os"
import path from "node:path"

type ProviderMetadata = { id: string; name: string; activation: "auto" | "enabled" | "disabled"; integrationID?: string; settings?: Readonly<Record<string, unknown>> }

// Catalog discovery describes possible setup, never authentication success.
// Only identifiers, labels and explicitly allowlisted nonsecret cloud options
// leave the engine; provider settings can otherwise contain credentials.
export function providerCatalog(providers: readonly ProviderMetadata[], environment: Record<string, string | undefined> = process.env,
  home = os.homedir(), exists: (file: string) => boolean = existsSync) {
  return providers.map(provider => {
    const google = provider.id === "google-vertex" || provider.id === "google-vertex-anthropic"
    const aws = provider.id === "amazon-bedrock"
    const options: Record<string, string> = {}
    for (const key of google ? ["project", "location"] : aws ? ["region", "profile"] : []) {
      const value = provider.settings?.[key]
      if (typeof value === "string") options[key] = value
    }
    if (aws && !options.profile && environment.AWS_PROFILE) options.profile = environment.AWS_PROFILE
    const sources = google ? [
      ...(environment.GOOGLE_APPLICATION_CREDENTIALS ? ["GOOGLE_APPLICATION_CREDENTIALS"] : []),
      ...(exists(path.join(home, ".config", "gcloud", "application_default_credentials.json")) ? ["Google ADC file"] : []),
    ] : aws ? [
      ...["AWS_PROFILE", "AWS_ACCESS_KEY_ID", "AWS_WEB_IDENTITY_TOKEN_FILE", "AWS_CONTAINER_CREDENTIALS_RELATIVE_URI", "AWS_CONTAINER_CREDENTIALS_FULL_URI", "AWS_BEARER_TOKEN_BEDROCK"].filter(key => environment[key]),
      ...(exists(environment.AWS_SHARED_CREDENTIALS_FILE ?? path.join(home, ".aws", "credentials")) || exists(environment.AWS_CONFIG_FILE ?? path.join(home, ".aws", "config")) ? ["AWS shared profiles"] : []),
    ] : []
    return { id: provider.id, name: provider.name, activation: provider.activation,
      ...(provider.integrationID === undefined ? {} : { integrationID: provider.integrationID }),
      ...(google || aws ? { host: { type: google ? "google-adc" as const : "aws-chain" as const, options, sources } } : {}),
    }
  })
}
