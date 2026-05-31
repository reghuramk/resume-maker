import OpenAI from 'openai'

const apiKey = import.meta.env.VITE_OPENAI_API_KEY
const client = apiKey
  ? new OpenAI({ apiKey, dangerouslyAllowBrowser: true })
  : null

const MODEL = (import.meta.env.VITE_OPENAI_MODEL as string) || 'gpt-4o'

/** The technical persona inferred from a JD (+ optional role title).
 *  Used by the "aggressive persona rewrite" mode to transmute every bullet's
 *  tech stack to a coherent target.
 *
 *  Confidence drives whether persona mode runs or falls back to keyword-swap. */
export interface Persona {
  primaryLanguage: string
  roleType: 'backend' | 'frontend' | 'fullstack'
  backendFramework?: string
  frontendFramework?: string
  database?: string
  cloud?: string
  secondaryTech: string[]
  confidence: 'high' | 'medium' | 'low'
  /** One-line rationale the model gives — surfaced to the user. */
  reasoning: string
}

const SYSTEM = `You analyze a job description (plus an optional role title) to extract the single technical persona a candidate should present. Be precise and concrete.

CONFIDENCE LEVELS:
- high: the JD names a specific primary language AND a specific framework (e.g. "Go engineer using Gin and PostgreSQL"). Use this freely; don't over-qualify.
- medium: the JD names a specific primary language but no framework — this is COMMON and still very usable. A JD saying "Mandatory Skills: Golang" or "Golang Preferred" or "experienced Python engineer" qualifies as MEDIUM, not low. The app will auto-infer the framework.
- low: ONLY when the JD truly names no specific language at all (e.g. "experienced software engineer" with no language listed anywhere). Do NOT return low just because the framework is unspecified or the role type is broad.

In short: if you can identify a primary language anywhere in the JD, you owe at least MEDIUM confidence. "Preferred" / "ideally" / "at least one of" wording does NOT reduce confidence — those are just JD politeness.

ROLE TYPE:
- backend: server-side focus (APIs, databases, queues, microservices).
- frontend: client-side focus (UI, components, user flows).
- fullstack: explicit end-to-end work or "fullstack engineer" in the title.

FIELDS:
- primaryLanguage: the dominant language (Go, Python, TypeScript, Java, C#, Ruby, Rust, etc.). Use the exact name the JD uses.
- backendFramework: only if the JD names one (Gin, FastAPI, Express, NestJS, Spring Boot, etc.).
- frontendFramework: only if the JD names one (React, Next.js, Vue, Angular, Svelte).
- database: only if named (PostgreSQL, MySQL, DynamoDB, etc.).
- cloud: only if named (AWS, GCP, Azure).
- secondaryTech: up to 10 other concrete tech terms, standards, protocols, or domain-specific concepts the JD emphasises ANYWHERE — including in "Nice to have" / "Good to have" / "Bonus" / "Preferred" sections. These often hold the niche keywords (C2PA, ODRL, W3C Verifiable Credentials, KYC, zero-knowledge proofs, etc.) that recruiters Boolean-search for. Pull them all in. Real names, not categories.
- reasoning: one short sentence explaining your choice.

Return JSON: { "persona": { ... } }
No commentary.`

export async function extractPersona(args: {
  jobDescription: string
  roleTitle?: string
}): Promise<Persona> {
  if (!client) {
    throw new Error('VITE_OPENAI_API_KEY is not set.')
  }

  const userBlock = [
    args.roleTitle ? `=== ROLE TITLE ===\n${args.roleTitle}\n` : '',
    `=== JOB DESCRIPTION ===\n${args.jobDescription}`,
  ]
    .filter(Boolean)
    .join('\n')

  const response = await client.chat.completions.create({
    model: MODEL,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: SYSTEM },
      { role: 'user', content: userBlock },
    ],
  })

  const raw = response.choices[0]?.message?.content ?? '{}'
  const parsed = JSON.parse(raw) as { persona?: Partial<Persona> }
  const p = parsed.persona ?? {}

  const primaryLanguage =
    typeof p.primaryLanguage === 'string' ? p.primaryLanguage : ''
  const roleType =
    p.roleType === 'backend' ||
    p.roleType === 'frontend' ||
    p.roleType === 'fullstack'
      ? p.roleType
      : 'fullstack'
  let backendFramework =
    typeof p.backendFramework === 'string' ? p.backendFramework : undefined
  let frontendFramework =
    typeof p.frontendFramework === 'string' ? p.frontendFramework : undefined

  // When the JD names a language confidently but no framework, infer the
  // most-common framework for that language. Without this, the persona
  // transmute step has nothing to swap "Spring Boot" against, drops the
  // bullet, and the original survives. The model can override at extract
  // time when the JD actually names a different framework.
  const defaults = DEFAULT_FRAMEWORKS[normalizeLang(primaryLanguage)]
  if (defaults) {
    if (
      !backendFramework &&
      (roleType === 'backend' || roleType === 'fullstack')
    ) {
      backendFramework = defaults.backend
    }
    if (
      !frontendFramework &&
      (roleType === 'frontend' || roleType === 'fullstack')
    ) {
      frontendFramework = defaults.frontend
    }
  }

  return {
    primaryLanguage,
    roleType,
    backendFramework,
    frontendFramework,
    database: typeof p.database === 'string' ? p.database : undefined,
    cloud: typeof p.cloud === 'string' ? p.cloud : undefined,
    secondaryTech: Array.isArray(p.secondaryTech)
      ? p.secondaryTech.filter((t): t is string => typeof t === 'string').slice(0, 10)
      : [],
    confidence:
      p.confidence === 'high' || p.confidence === 'medium' || p.confidence === 'low'
        ? p.confidence
        : 'low',
    reasoning: typeof p.reasoning === 'string' ? p.reasoning : '',
  }
}

/** Hard-coded defaults — the most-common framework choice per language for
 *  resume-tailoring contexts. Used only when the JD itself doesn't name one
 *  explicitly. */
const DEFAULT_FRAMEWORKS: Record<
  string,
  { backend?: string; frontend?: string }
> = {
  go: { backend: 'Gin' },
  golang: { backend: 'Gin' },
  python: { backend: 'FastAPI' },
  typescript: { backend: 'Express', frontend: 'React' },
  javascript: { backend: 'Express', frontend: 'React' },
  java: { backend: 'Spring Boot' },
  kotlin: { backend: 'Ktor' },
  csharp: { backend: 'ASP.NET' },
  'c#': { backend: 'ASP.NET' },
  ruby: { backend: 'Rails' },
  rust: { backend: 'Actix' },
  php: { backend: 'Laravel' },
  scala: { backend: 'Play' },
  elixir: { backend: 'Phoenix' },
}

function normalizeLang(s: string): string {
  return s.toLowerCase().trim()
}

/** Format the persona as a short stack summary (for prompts and UI). */
export function describePersona(p: Persona): string {
  const parts: string[] = []
  if (p.primaryLanguage) parts.push(p.primaryLanguage)
  if (p.backendFramework) parts.push(p.backendFramework)
  if (p.frontendFramework) parts.push(p.frontendFramework)
  if (p.database) parts.push(p.database)
  if (p.cloud) parts.push(p.cloud)
  parts.push(...p.secondaryTech)
  return `${p.roleType} (${parts.join(' + ')})`
}
