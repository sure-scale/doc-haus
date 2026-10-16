// Pure validators for the doc.haus skill format. Single source of truth for
// every shape rule a skill must satisfy. Used by both the commit-time lint
// pass (script/lint-legal-runtime.ts) and the author-time tool seam
// (tool/create-skill.ts, tool/update-skill.ts).
//
// Rules are value-agnostic: structural shape and general prose clarity only.
// No legal content, no named jurisdictions, no fixture values. The prose
// rule targets higher-frequency paraphrases of instructional wording only
// (federal plain-language guidance), never legal terms of art. The built-in
// skill list is passed in by the caller so this module never hardcodes it.

export type Severity = "error" | "warn"

export type Finding = {
  severity: Severity
  rule: string
  message: string
  line?: number
}

export type SkillInput = {
  name: string
  path: string
  frontmatter: string
  description: string
  body: string
  references?: string[]
  builtins: Set<string>
}

const TRIGGER_REGEX = /Use\s+(when|whenever|before|after|while)\b/
const NAME_REGEX = /^[a-z0-9]+(-[a-z0-9]+)*$/
const COMPLETION_REGEX = /\b(verdict|complete|done|stop|only when|never|final|safe to|done when)\b/

export function parseSkillFile(raw: string): {
  frontmatter: string
  description: string
  body: string
} {
  const match = raw.match(/^---\n([\s\S]*?)\n---\n?([\s\S]*)$/)
  if (!match) return { frontmatter: "", description: "", body: raw }
  const frontmatter = match[1]
  const body = match[2]
  const descriptionLine = frontmatter.match(/^description:\s+(.+?)\s*$/m)?.[1] ?? ""
  const description = stripMatchingQuotes(descriptionLine)
  return { frontmatter, description, body }
}

function stripMatchingQuotes(value: string) {
  if (value.length < 2) return value
  const first = value[0]
  if ((first === '"' || first === "'") && value[value.length - 1] === first) {
    return value.slice(1, -1)
  }
  return value
}

export function lintSkill(input: SkillInput): Finding[] {
  const findings: Finding[] = []

  if (!input.description) {
    findings.push({
      severity: "error",
      rule: "frontmatter-description-missing",
      message: "Skill is missing a description in frontmatter.",
    })
  } else if (!TRIGGER_REGEX.test(input.description)) {
    findings.push({
      severity: "error",
      rule: "description-trigger-missing",
      message:
        "Description must contain a 'Use when/whenever/before/after/while ...' trigger phrase naming when the skill applies.",
    })
  }

  if (!NAME_REGEX.test(input.name)) {
    findings.push({
      severity: "error",
      rule: "name-pattern",
      message: `Skill name "${input.name}" must match ^[a-z0-9]+(-[a-z0-9]+)*$.`,
    })
  }

  if (input.builtins.has(input.name)) {
    findings.push({
      severity: "error",
      rule: "name-collision-builtin",
      message: `Skill name "${input.name}" collides with a built-in skill.`,
    })
  }

  if (input.name.startsWith("playbook-")) {
    findings.push({
      severity: "error",
      rule: "name-playbook-reserved",
      message: `Skill name "${input.name}" starts with "playbook-" — playbooks have their own importer.`,
    })
  }

  for (const ref of input.references ?? []) {
    const slashes = (ref.match(/\//g) ?? []).length
    if (slashes > 2) {
      findings.push({
        severity: "error",
        rule: "references-nesting-too-deep",
        message: `Reference file "${ref}" is nested more than one level under references/.`,
      })
    }
  }

  const lines = input.body.split("\n")
  const lineCount = lines.length
  if (lineCount > 1000) {
    findings.push({
      severity: "error",
      rule: "body-over-budget-hard",
      message: `Body is ${lineCount} lines; hard limit is 1000.`,
    })
  } else if (lineCount > 500) {
    findings.push({
      severity: "warn",
      rule: "body-over-budget-soft",
      message: `Body is ${lineCount} lines; budget is 500.`,
    })
  }

  const seen = new Set<string>()
  const duplicates = new Set<string>()
  for (const line of lines) {
    const md = line.match(/^#{2,3}\s+(.+?)\s*$/)
    if (md) {
      const heading = md[1]
      if (seen.has(heading)) duplicates.add(heading)
      seen.add(heading)
      continue
    }
    const angle = line.match(/^<([a-zA-Z][\w-]*)>/)
    if (angle) {
      const tag = angle[1]
      if (seen.has(tag)) duplicates.add(tag)
      seen.add(tag)
    }
  }
  for (const dup of duplicates) {
    findings.push({
      severity: "warn",
      rule: "duplicate-section-heading",
      message: `Section heading "${dup}" appears more than once in the body.`,
    })
  }

  if (/<procedure>|<workflow>/.test(input.body)) {
    const tailStart = Math.max(0, lineCount - Math.max(1, Math.floor(lineCount * 0.3)))
    const tail = lines.slice(tailStart).join("\n").toLowerCase()
    if (!COMPLETION_REGEX.test(tail)) {
      findings.push({
        severity: "warn",
        rule: "procedure-missing-completion",
        message:
          "Procedure/workflow skill has no terminal-state language (verdict, complete, stop, never, final, safe to, ...) in the last 30% of the body.",
      })
    }
  }

  findings.push(...lintProseFrequency(input.body))

  return findings
}

// Higher-frequency paraphrases of common instructional wording. The pairs are
// general plain-language substitutions (federal Plain Language guidelines);
// legal terms of art are deliberately omitted. Independence from readability
// is the point: these words are rarer than their everyday synonyms, which
// costs model processing even when the sentence is otherwise simple.
const PROSE_FREQUENCY_SUBSTITUTIONS: Array<[RegExp, string]> = [
  [/\butiliz(?:e[ds]?|ing|ation)\b/gi, "use"],
  [/\bendeavo[u]?r(?:ed|ing|s)?\b/gi, "try"],
  [/\bascertain(?:ed|ing|s)?\b/gi, "find or check"],
  [/\bfacilitat(?:e[ds]?|ing)\b/gi, "help"],
  [/\belucidat(?:e[ds]?|ing)\b/gi, "explain"],
  [/\bdelineat(?:e[ds]?|ing)\b/gi, "outline"],
  [/\bsubsequent to\b/gi, "after"],
  [/\bin order to\b/gi, "to"],
  [/\bdue to the fact that\b/gi, "because"],
  [/\bin the event that\b/gi, "if"],
  [/\bfor the purpose of\b/gi, "to"],
  [/\bit should be noted that\b/gi, "(omit)"],
  [/\ba number of\b/gi, "several or many"],
]

const CODE_FENCE_RE = /```[\s\S]*?```/g
const INLINE_CODE_RE = /`[^`\n]*`/g

// Lints instructional prose for lower-frequency wording that has a common
// everyday synonym. Warn-only and subjective; legal terminology is out of
// scope by construction (the substitution list is prose-only). Code spans
// and fenced blocks are stripped before scanning so identifiers, examples,
// and command syntax are not flagged.
export function lintProseFrequency(body: string): Finding[] {
  const stripped = body.replace(CODE_FENCE_RE, " ").replace(INLINE_CODE_RE, " ")
  const findings: Finding[] = []
  for (const [pattern, suggestion] of PROSE_FREQUENCY_SUBSTITUTIONS) {
    const matches = stripped.match(pattern)
    if (!matches || !matches.length) continue
    const phrase = matches[0].toLowerCase()
    findings.push({
      severity: "warn",
      rule: "prose-frequency-plainness",
      message:
        matches.length > 1
          ? `Instructional prose uses lower-frequency phrase "${phrase}" (${matches.length}x); consider "${suggestion}".`
          : `Instructional prose uses lower-frequency phrase "${phrase}"; consider "${suggestion}".`,
    })
  }
  return findings
}
