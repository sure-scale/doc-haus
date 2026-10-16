

// The host-credential setups that still need a project/region to call. One card
// can cover several provider ids that share the same credentials and settings —
// Gemini (google-vertex) and Claude (google-vertex-anthropic) are separate
// provider ids but the same GCP project and gcloud sign-in, so they get one
// card. ids[0] is the primary: the card's visibility and displayed error track
// it alone. Each field maps to a provider.options key the engine reads (see
// provider.ts): vertex -> { project, location }, bedrock -> { region, profile }.
// A field is either a select over a fixed list (locations/regions, defaulting
// via fallback), or an input whose suggestions come from what the host's
// sign-in can actually see (GCP projects, AWS profiles) — still typeable, since
// the host may have no sign-in to enumerate yet.
export type CloudField = {
  key: string
  label: string
  placeholder: string
  required: boolean
  fallback?: string
  options?: string[]
  suggest?: "gcp-projects" | "aws-profiles"
}
export type CloudSpec = { ids: string[]; name: string; signin: string; fields: CloudField[] }

// The Vertex endpoint for Anthropic publisher models, mirroring Google's
// scheme: regional locations prefix the host, "global" does not.
export function vertexAnthropicBaseURL(project: string, location: string) {
  const host = location === "global" ? "aiplatform.googleapis.com" : `${location}-aiplatform.googleapis.com`
  return `https://${host}/v1/projects/${project}/locations/${location}/publishers/anthropic/models`
}

export const VERTEX_LOCATIONS = [
  "global",
  "us-central1",
  "us-east1",
  "us-east4",
  "us-east5",
  "us-west1",
  "us-west4",
  "northamerica-northeast1",
  "southamerica-east1",
  "europe-west1",
  "europe-west2",
  "europe-west3",
  "europe-west4",
  "europe-west9",
  "europe-north1",
  "asia-east1",
  "asia-northeast1",
  "asia-northeast3",
  "asia-south1",
  "asia-southeast1",
  "australia-southeast1",
  "me-central1",
]

export const BEDROCK_REGIONS = [
  "us-east-1",
  "us-east-2",
  "us-west-2",
  "ca-central-1",
  "sa-east-1",
  "eu-central-1",
  "eu-central-2",
  "eu-west-1",
  "eu-west-2",
  "eu-west-3",
  "eu-north-1",
  "ap-northeast-1",
  "ap-northeast-2",
  "ap-south-1",
  "ap-southeast-1",
  "ap-southeast-2",
]

export const CLOUD_SETUPS: CloudSpec[] = [
  {
    ids: ["google-vertex", "google-vertex-anthropic"],
    name: "Google Vertex AI",
    signin: "gcloud auth application-default login",
    fields: [
      { key: "project", label: "GCP project id", placeholder: "my-gcp-project", required: true, suggest: "gcp-projects" },
      { key: "location", label: "Location", placeholder: "global", required: false, fallback: "global", options: VERTEX_LOCATIONS },
    ],
  },
  {
    ids: ["amazon-bedrock"],
    name: "Amazon Bedrock",
    signin: "aws configure (profile or role)",
    fields: [
      { key: "region", label: "AWS region", placeholder: "us-east-1", required: true, fallback: "us-east-1", options: BEDROCK_REGIONS },
      { key: "profile", label: "AWS profile", placeholder: "default", required: false, suggest: "aws-profiles" },
    ],
  },
]
