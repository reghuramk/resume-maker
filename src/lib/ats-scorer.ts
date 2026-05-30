import OpenAI from 'openai'

const apiKey = import.meta.env.VITE_OPENAI_API_KEY
const client = apiKey
  ? new OpenAI({ apiKey, dangerouslyAllowBrowser: true })
  : null

/**
 * A must-have requirement extracted from the job description.
 *
 * In practice, recruiters search ATSes (Workday, Greenhouse, Lever, …) with
 * Boolean AND-queries against parsed resume text. To pass a recruiter
 * filter, the resume must contain at least one form of each must-have term.
 *
 *  - term: canonical phrase as a recruiter would phrase it ("Python", "Node.js").
 *  - aliases: surface forms a candidate might use ("python", "Py3", "Node", "NodeJS").
 *             Workday's Skills Cloud does this internally; we replicate it.
 *  - category: language / database / cloud / etc. — for UI grouping only.
 */
export interface MustHave {
  term: string
  aliases: string[]
  category: string
}

export interface BooleanMatch {
  /** Must-haves that have at least one alias present in the text. */
  matched: MustHave[]
  /** Must-haves with no alias present anywhere. */
  missing: MustHave[]
  /** Convenience flag: every must-have matched. */
  passed: boolean
  /** 0–100 fraction of must-haves matched. */
  scorePct: number
}

/**
 * Ask the model to extract the recruiter-search must-haves from the JD.
 * Run once per JD; reused for every per-bullet gate + final resume check.
 *
 * Why few must-haves (target 6–12) rather than dozens: real recruiter
 * Boolean queries are focused. A 20-term AND-query filters out everyone.
 * Recruiters typically AND 3–6 hard requirements and treat the rest as
 * "nice to surface" / ranking signals.
 */
export async function extractMustHaves(jd: string): Promise<MustHave[]> {
  if (!client) {
    throw new Error('VITE_OPENAI_API_KEY is not set.')
  }

  const SYSTEM = `You are a technical recruiter. From a job description, identify the MUST-HAVE terms you would put in a Boolean AND-query to filter ATS results.

What qualifies as a must-have:
- Explicit hard requirements (in "Requirements" / "Technical Skills" / "Key Responsibilities").
- Specific, named technologies / frameworks / databases / cloud services / methodologies.
- Recruiters can realistically AND together 6–12 of these. Pick the ones THIS JD treats as non-negotiable.

What does NOT qualify:
- Soft skills (communication, teamwork, leadership) — recruiters don't Boolean-search these.
- Generic verbs (build, develop, collaborate, design).
- Aspirational nice-to-haves buried in the prose.
- Years-of-experience phrases.

For each must-have, also list realistic ALIASES — the surface forms a candidate might write on their resume. E.g.:
- "Node.js" → aliases: ["Node.js", "NodeJS", "Node JS", "Node"]
- "Python" → aliases: ["Python", "python", "Py3", "Python 3"]
- "PostgreSQL" → aliases: ["PostgreSQL", "Postgres", "PG"]
- "REST APIs" → aliases: ["REST APIs", "RESTful APIs", "REST API"]
- "CI/CD" → aliases: ["CI/CD", "CI / CD", "CICD", "continuous integration"]

Include the canonical term itself in aliases. Don't include misspellings.

Output: JSON { "mustHaves": [ { "term": "...", "aliases": [...], "category": "language|backend|frontend|database|cache|queue|cloud|orchestration|iac|observability|cicd|api|ai|methodology|concept" } ] }
6–12 entries. No commentary.`

  const response = await client.chat.completions.create({
    model: 'gpt-4o',
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: SYSTEM },
      { role: 'user', content: jd },
    ],
  })

  const raw = response.choices[0]?.message?.content ?? '{}'
  const parsed = JSON.parse(raw) as { mustHaves?: MustHave[] }
  return (parsed.mustHaves ?? []).filter(
    (m) =>
      typeof m?.term === 'string' &&
      m.term.trim().length > 0 &&
      Array.isArray(m.aliases) &&
      m.aliases.length > 0,
  )
}

/**
 * Run the recruiter's Boolean AND-query against a text blob.
 * Each must-have is satisfied if ANY of its aliases appears as a token in
 * the text (case-insensitive, tech-punctuation-aware boundaries).
 */
export function checkBooleanQuery(
  text: string,
  mustHaves: MustHave[],
): BooleanMatch {
  const haystack = text.toLowerCase()
  const matched: MustHave[] = []
  const missing: MustHave[] = []
  for (const m of mustHaves) {
    const hit = m.aliases.some((alias) => containsToken(haystack, alias.toLowerCase()))
    ;(hit ? matched : missing).push(m)
  }
  const total = mustHaves.length || 1
  return {
    matched,
    missing,
    passed: missing.length === 0,
    scorePct: Math.round((matched.length / total) * 100),
  }
}

/** 1 if `text` contains any alias of the given must-have, else 0. Used to
 *  count "how many bullets contain this term" across the resume so we can
 *  cap repetition. */
export function bulletContainsMustHave(text: string, m: MustHave): 0 | 1 {
  const hay = text.toLowerCase()
  return m.aliases.some((a) => containsToken(hay, a.toLowerCase())) ? 1 : 0
}

/**
 * Detects rewrites that mix a frontend framework with strong backend
 * activity markers — the failure mode where the LLM stuffs "Next.js" into
 * a bullet about REST APIs and connection pooling. The model is told not
 * to do this in the prompt; this is the belt-and-suspenders code check.
 *
 * Returns the offending pair if detected (so the UI can surface it),
 * otherwise null.
 */
export function detectStackMismatch(
  bullet: string,
): { frontendTerm: string; backendTerm: string } | null {
  const hay = bullet.toLowerCase()
  // Frameworks that are unambiguously frontend in resume context. React is
  // ambiguous on its own (could be React Native, full-stack components,
  // etc.) so we skip it here to avoid false positives. Next.js, Vue,
  // Angular, Svelte, and Redux are far more clearly frontend.
  const FRONTEND_ONLY = [
    'next.js',
    'nextjs',
    'next js',
    'vue.js',
    'vuejs',
    'angular',
    'svelte',
    'redux',
  ]
  // Activities a real frontend bullet would not describe. If one of these
  // appears alongside a frontend-only term, the bullet is incoherent.
  const STRONG_BACKEND = [
    'rest api',
    'restful',
    'rest apis',
    'grpc',
    'connection pool',
    'inter-service',
    'microservice',
    'message queue',
    'event bridge',
    'sqs',
    'kafka',
    'kubernetes pod',
    'distributed system',
    'async worker',
    'data pipeline',
  ]
  let frontendHit: string | null = null
  for (const f of FRONTEND_ONLY) {
    if (containsToken(hay, f)) {
      frontendHit = f
      break
    }
  }
  if (!frontendHit) return null
  for (const b of STRONG_BACKEND) {
    if (containsToken(hay, b)) {
      return { frontendTerm: frontendHit, backendTerm: b }
    }
  }
  return null
}

/** True iff `bullet` contains an alias for some must-have that `original` did not. */
export function bulletAddsMustHave(
  bullet: string,
  original: string,
  mustHaves: MustHave[],
): MustHave[] {
  const bHay = bullet.toLowerCase()
  const oHay = original.toLowerCase()
  const added: MustHave[] = []
  for (const m of mustHaves) {
    const inOriginal = m.aliases.some((a) => containsToken(oHay, a.toLowerCase()))
    if (inOriginal) continue
    const inBullet = m.aliases.some((a) => containsToken(bHay, a.toLowerCase()))
    if (inBullet) added.push(m)
  }
  return added
}

/**
 * Token-boundary check that survives tech punctuation. "Node.js" matches
 * inside "built a Node.js gateway" but not inside "Nodejsfun" — boundary
 * is defined by surrounding char being a non-word char (not by \b, which
 * treats "." as a boundary itself).
 */
function containsToken(haystack: string, needle: string): boolean {
  if (!needle) return false
  let from = 0
  while (true) {
    const idx = haystack.indexOf(needle, from)
    if (idx === -1) return false
    const before = idx === 0 ? ' ' : haystack[idx - 1]
    const after =
      idx + needle.length >= haystack.length
        ? ' '
        : haystack[idx + needle.length]
    if (!isWordChar(before) && !isWordChar(after)) return true
    from = idx + 1
  }
}

function isWordChar(c: string): boolean {
  return /[a-z0-9]/i.test(c)
}
