import OpenAI from 'openai'
import type { DocxParagraph } from './extract-docx'
import {
  extractMustHaves,
  checkBooleanQuery,
  bulletAddsMustHave,
  bulletContainsMustHave,
  detectStackMismatch,
  type MustHave,
  type BooleanMatch,
} from './ats-scorer'
import { extractPersona, describePersona, type Persona } from './persona'

const apiKey = import.meta.env.VITE_OPENAI_API_KEY

const client = apiKey
  ? new OpenAI({ apiKey, dangerouslyAllowBrowser: true })
  : null

/** OpenAI rate limits are per-model. If you hit the gpt-4o quota, set
 *  VITE_OPENAI_MODEL=gpt-4o-mini (or gpt-4-turbo, gpt-4.1, etc.) in .env.local
 *  and restart vite — switches to a different rate-limit pool. */
const MODEL = (import.meta.env.VITE_OPENAI_MODEL as string) || 'gpt-4o'
const MAX_ATTEMPTS = 5
/** ±6 keeps line wrapping nearly identical for most bullets (a single word
 *  rarely crosses Word's wrap point) while making rewrites land more
 *  reliably. Was ±4; widened because too many keyword-positive rewrites
 *  were being dropped, which hurt Boolean coverage. */
const WORD_COUNT_TOLERANCE = 6
/** Hard cap: a single must-have keyword may appear in at most this many
 *  bullets across the final resume. Enforced in code, not just in prompt
 *  (prompt-only caps are routinely ignored by the model). */
const MAX_BULLETS_PER_KEYWORD = 2

const SYSTEM_PROMPT = `You are an ATS-optimization expert who rewrites resumes by iterating through JD KEYWORDS (not through bullets) and placing each keyword where it fits naturally. Your output must read coherently to an engineer — every bullet you touch must describe a plausible, real-world system.

==============================
ALGORITHM (FOLLOW EXACTLY)
==============================

STEP 1 — EXTRACT JD KEYWORDS
List every concrete keyword from the JD, grouped by category:
- Languages (Python, TypeScript, Java, etc.)
- Backend frameworks (FastAPI, NestJS, Express, etc.)
- Frontend frameworks (React, Vue, Angular)
- Databases / ORM (MySQL, PostgreSQL, DynamoDB; ORM tools)
- Cloud / DevOps (AWS, Docker, Kubernetes, Terraform)
- Observability (Datadog, CloudWatch, OpenTelemetry)
- Queues / events (SQS, EventBridge, Kafka)
- AI / agentic (agentic workflows, AI-powered tools, LLM, RAG, LangChain)
- Concepts (full-stack, mentoring, high-velocity delivery, testing rigor)

STEP 2 — DETERMINE THE COHERENT STACK FOR THIS RESUME
Choose ONE primary backend language and ONE primary frontend (from the JD). Every backend bullet you rewrite uses that ONE backend language; every frontend bullet you rewrite uses that ONE frontend. Do not mix Python + Node.js + TypeScript randomly across the same role. This is the #1 cause of broken-sounding rewrites — fix it before writing anything.

STEP 3 — PLAN KEYWORD PLACEMENT (iterate keyword by keyword)
For each JD keyword in order:
  a. Walk through the work experience bullets. Find the bullet whose ACTIVITY genuinely aligns with the keyword's CATEGORY:
     - "MySQL" → a bullet whose activity involves a database
     - "agentic workflows" / "AI-powered tools" → a bullet whose activity involves AI or automation
     - "Docker" → a bullet about containerization, deployment, or environments
     - "ORM" → a bullet about data access / database queries
     - "mentoring" → a bullet about working with junior engineers, code review
     - "React" → a bullet about UI, frontend, dashboards
  b. If you find a bullet where this keyword can be woven in NATURALLY (the sentence still describes a coherent, real system after insertion): place it there.
  c. If no bullet's activity aligns: SKIP this keyword. Do not force it.
  d. A keyword may appear in at most ${MAX_BULLETS_PER_KEYWORD} bullets (HARD CAP — the system silently reverts any rewrite that violates this) total. Stop placing once you hit the cap.

STEP 4 — VERIFY THE PLAN
Before generating rewrites, run this checklist:
  □ Is the backend language consistent across each role's bullets?
  □ Did I avoid any nonsensical combination (Python + Spring Batch, AWS Lambda powering an OpenAI assistant, Spring Web Flux on Lambda)?
  □ Does every rewrite still describe a system that could exist in the real world?
  □ Is every keyword I added actually present in the JD verbatim?

If any check fails, revise the plan or drop that keyword.

STEP 5 — GENERATE REWRITES
Now write the final rewrites. For each bullet you're modifying:
- Weave the planned keyword(s) in naturally — restructure the sentence if needed.
- Keep verbs (architected, built, designed), metrics (10K+, 90%, 50ms), employer, role, dates, business context verbatim.
- Stay within ±${WORD_COUNT_TOLERANCE} words of the original.

Bullets where no keyword could be placed: OMIT from your output. Don't generate cosmetic edits.

==============================
ACTIVITY-LOCKED SWAPS (BACKEND vs FRONTEND vs DEVOPS vs AI)
==============================

Before adding a keyword to a bullet, classify the bullet's primary activity:

- BACKEND activity: REST APIs, GraphQL servers, gRPC, microservices, connection pooling, database queries, async workers, message queues, server-side data processing.
- FRONTEND activity: UI components, dashboards, pages, user-facing flows, client-side state.
- DEVOPS activity: containerization, orchestration, CI/CD, infrastructure-as-code, observability/monitoring setup.
- AI activity: LLM integrations, embeddings, RAG, prompt engineering, ML pipelines.

A keyword may ONLY land in a bullet whose activity matches its own category:

- Frontend frameworks (React, Next.js, Vue, Angular, Svelte, Redux) → ONLY frontend bullets. Never a bullet about REST APIs, gRPC, microservices, or queues.
- Backend frameworks (Express, NestJS, FastAPI, Spring Boot) → ONLY backend bullets.
- Databases (PostgreSQL, MySQL, DynamoDB) → backend or devops bullets — NEVER frontend bullets.
- gRPC / GraphQL server / REST API → backend bullets.
- Kubernetes / Docker / Terraform → devops bullets.

If a must-have's only natural-fit bullet doesn't exist in this resume, LEAVE IT UNPLACED. Missing coverage is better than a sentence that doesn't describe a real system.

==============================
NEGATIVE EXAMPLES (DO NOT DO THESE)
==============================

✗ "Architected a production-grade quoting platform in TypeScript, building REST APIs for 10K+ daily users, processing 1K+ quotes/day via connection pooling and Next.js-based dynamic interfaces."
   → Next.js is a FRONTEND framework. This bullet's activity is BACKEND (REST APIs, connection pooling). Frontend frameworks NEVER land in backend bullets. Skip Next.js for this bullet entirely.

✗ "Migrated inter-service communication to gRPC, reducing latency by 40%; a Next.js gateway consumed gRPC to serve structured frontend responses."
   → "Next.js gateway" is not a real architecture — Next.js is a React framework, not a server-side gateway service. Real gateways are Express / Nginx / Kong. Skip Next.js for this bullet.

✗ "Python FastAPI data processing pipeline using Spring Batch"
   → Spring Batch is Java-only. You cannot use it from Python. Cross-stack nonsense.

✗ "AWS Lambda-powered assistants and Serverless workflows for quote generation, collaborating with ML teams"
   → Lambda is compute, not an LLM. Replace OpenAI with LangChain / Anthropic / another LLM tool — never with infrastructure.

✗ "Node.js gateway consumed GraphQL to serve structured frontend responses" in a role that previously said Python
   → Tech stack within a single role flips mid-sentence. Pick ONE language per role.

✗ "Spring Web Flux reactive pipelines" → "Lambda reactive pipelines"
   → Spring Web Flux is a reactive framework. Lambda is a FaaS runtime. Not parallel.

✗ Same keyword (e.g., "AWS Lambda") appearing in 5 different bullets
   → Distribution cap is ${MAX_BULLETS_PER_KEYWORD} bullets (HARD CAP — the system silently reverts any rewrite that violates this) per keyword max.

==============================
POSITIVE EXAMPLE
==============================

JD says: Python, React, MySQL, ORM, Docker, AWS, agentic workflows, AI-powered tools, mentoring.

Original bullets (sample):
[1] "Architected a quoting platform in Spring Boot, building REST APIs for 10K+ daily users via connection pooling."
[2] "Built and productionized OpenAI-powered assistants for quote generation, collaborating with ML teams."
[3] "Containerised distributed microservices on AWS/GCP Kubernetes with Redis caching and monitoring."
[4] "Mentored junior developers, contributing to code reviews and technical documentation."

Plan:
- Python → bullet [1] (replaces Spring Boot)
- agentic workflows → bullet [2] (already an AI-tooling bullet)
- AI-powered tools → bullet [2] (same bullet, complements)
- MySQL + ORM → bullet [1] (REST + DB activity)
- Docker → bullet [3] (containerisation)
- AWS → bullet [3] (already AWS)
- mentoring → bullet [4] (already mentoring)
- React → no clearly-aligned bullet in this role → SKIP

Rewrites:
[1] "Architected a quoting platform in Python with MySQL and an ORM layer, building REST APIs for 10K+ daily users via connection pooling."
[2] "Built and productionized agentic AI-powered assistants for quote generation, collaborating with ML teams."
[3] "Containerised distributed microservices on AWS using Docker and Kubernetes with Redis caching and monitoring."
[4] (unchanged — original is already JD-aligned, omit from response)

==============================
OUTPUT
==============================

Return JSON: { "replacements": { "<index>": "<rewritten text>", ... } }

Include only the bullets you actually modified. Bullets that don't have a keyword placement → omit.

FINAL SELF-CHECK (do this for every rewrite before responding):
Read each rewrite as if you were a senior engineer reviewing the resume. Ask: "Could this system actually exist? Does the tech stack inside this single sentence make architectural sense?"
- A frontend framework cannot be a gateway.
- A Python service cannot use a Java-only library (e.g. Spring Batch).
- Lambda is compute, not an LLM.
- Microservices and connection pooling are backend concerns — frontend frameworks don't belong there.
If a rewrite describes an impossible or implausible system, REMOVE it from your output (revert to the original) rather than ship it.

Re-run the STEP 4 checklist one more time. If anything fails, fix or drop.

No commentary outside the JSON.`

export interface RewriteResult {
  replacements: Record<number, string>
  wordCountMismatches: Array<{
    index: number
    originalWords: number
    newWords: number
  }>
  attemptsUsed: number
  /** Per-bullet: which must-have aliases this rewrite introduced (vs. original). */
  addedByIndex: Record<number, string[]>
  /** The recruiter's Boolean must-have list extracted from the JD. */
  mustHaves: MustHave[]
  /** Result of running the recruiter's AND-query against the tailored resume. */
  booleanMatch: BooleanMatch
  /** New bullets generated to fill must-have gaps; each is inserted after
   *  `templateAfterIndex` (the last bullet of the chosen role) so its
   *  formatting matches the surrounding bullets. */
  generatedBullets: Array<{
    text: string
    templateAfterIndex: number
    mustHaveTerm: string
    roleHeader: string
  }>
  /** Bullet indices whose rewrite was reverted because it would have pushed
   *  some must-have keyword past the repetition cap. Useful for UI hints. */
  bumpedForRepetition: number[]
  /** Bullet indices where the rewrite was accepted as a last resort despite
   *  exceeding word-count tolerance, because dropping it would have lost a
   *  must-have that's otherwise missing globally. Layout may shift a line. */
  alignmentRelaxed: number[]
  /** Bullet indices initially bumped by the repetition cap but then
   *  resurrected because their rewrite was the only way to cover a
   *  currently-missing must-have. Trade-off: one keyword may now appear in
   *  3 bullets instead of 2, but Boolean coverage is higher. */
  resurrectedForCoverage: number[]
  /** Tier-3: low-value bullets that were ENTIRELY REPLACED with fabricated
   *  content to cover a still-missing must-have. The replacement uses a
   *  scale figure already present in the resume; the original text is
   *  discarded. Strong UI warning surfaces these for review. */
  fabricatedReplacements: Array<{
    index: number
    originalText: string
    newText: string
    mustHaveTerm: string
  }>
  /** Which rewrite mode actually ran:
   *   - "persona-transmute": user asked for persona mode AND extraction was
   *     confident; every bullet's tech stack was transmuted to the persona.
   *   - "keyword-swap": the default mode; either the user didn't ask for
   *     persona mode, or extraction came back low-confidence and we fell
   *     back automatically. */
  mode: 'persona-transmute' | 'keyword-swap'
  /** Populated when mode === 'persona-transmute'. */
  persona?: Persona
  /** Populated when persona mode was REQUESTED but we fell back. Tells the
   *  user why (e.g. "JD too generic to extract a confident persona"). */
  personaFallbackReason?: string
  /** Must-have terms appended to the SKILLS section because they were still
   *  missing globally after every other pass. Keyed by paragraph index. */
  skillsPaddingApplied: Record<number, string[]>
}

function wordCount(s: string): number {
  return s.trim().split(/\s+/).filter(Boolean).length
}

/** Walk accepted rewrites in document order and revert any that would cause
 *  a must-have keyword to appear in more than `cap` bullets. A rewrite is
 *  blocked only if it (a) ADDS a new occurrence of some keyword AND (b) that
 *  would push the keyword's total past the cap. Rewrites that don't change
 *  a keyword's count are always allowed; rewrites that REMOVE a keyword
 *  somewhere can free budget for other rewrites later.
 *
 *  `exemptLowerTerms`: keywords that bypass the cap entirely. In persona
 *  transmute mode, the persona's core stack (primary language, backend
 *  framework, etc.) goes in here — a Go resume *should* say "Golang" in
 *  every backend bullet without that counting as keyword stuffing. */
function enforceRepetitionCap(
  bullets: Array<{ index: number; originalText: string }>,
  acceptedRewrites: Record<number, string>,
  mustHaves: MustHave[],
  cap: number,
  exemptLowerTerms: Set<string> = new Set(),
): { final: Record<number, string>; bumped: number[] } {
  const isExempt = (m: MustHave): boolean =>
    exemptLowerTerms.has(m.term.toLowerCase()) ||
    m.aliases.some((a) => exemptLowerTerms.has(a.toLowerCase()))
  // Seed: count of bullets containing each must-have ASSUMING all originals
  // are still in place. As we accept rewrites, we adjust this count.
  const counts = new Map<string, number>()
  for (const m of mustHaves) {
    let c = 0
    for (const b of bullets) c += bulletContainsMustHave(b.originalText, m)
    counts.set(m.term, c)
  }

  const final: Record<number, string> = {}
  const bumped: number[] = []

  for (const b of bullets) {
    const rewrite = acceptedRewrites[b.index]
    if (!rewrite) continue

    // Decide whether this rewrite causes any must-have to exceed cap.
    let blocked = false
    for (const m of mustHaves) {
      if (isExempt(m)) continue
      const inOrig = bulletContainsMustHave(b.originalText, m)
      const inNew = bulletContainsMustHave(rewrite, m)
      const delta = inNew - inOrig
      const newTotal = (counts.get(m.term) ?? 0) + delta
      // Only block when the rewrite is ACTIVELY adding a new occurrence AND
      // that would put us over the cap. Don't penalize rewrites that left
      // an already-over count alone.
      if (delta > 0 && newTotal > cap) {
        blocked = true
        break
      }
    }

    if (blocked) {
      bumped.push(b.index)
      continue
    }
    // Apply the delta for every must-have.
    for (const m of mustHaves) {
      const inOrig = bulletContainsMustHave(b.originalText, m)
      const inNew = bulletContainsMustHave(rewrite, m)
      counts.set(
        m.term,
        (counts.get(m.term) ?? 0) - inOrig + inNew,
      )
    }
    final[b.index] = rewrite
  }
  return { final, bumped }
}

/** Same idea but for gap-fill bullets: they're additive (no original to
 *  subtract). Drop any gap-fill that would push a must-have past the cap.
 *  Same `exemptLowerTerms` semantics as above. */
function filterGapFillsByCap(
  bullets: Array<{ originalText: string }>,
  acceptedRewrites: Record<number, string>,
  gapFills: GapFillBullet[],
  mustHaves: MustHave[],
  cap: number,
  exemptLowerTerms: Set<string> = new Set(),
): GapFillBullet[] {
  const isExempt = (m: MustHave): boolean =>
    exemptLowerTerms.has(m.term.toLowerCase()) ||
    m.aliases.some((a) => exemptLowerTerms.has(a.toLowerCase()))
  const counts = new Map<string, number>()
  for (const m of mustHaves) {
    let c = 0
    for (const b of bullets) {
      const text = acceptedRewrites[(b as { index?: number }).index ?? -1] ?? b.originalText
      c += bulletContainsMustHave(text, m)
    }
    counts.set(m.term, c)
  }

  const accepted: GapFillBullet[] = []
  for (const g of gapFills) {
    let blocked = false
    for (const m of mustHaves) {
      if (isExempt(m)) continue
      const inGap = bulletContainsMustHave(g.text, m)
      if (inGap === 0) continue
      const newTotal = (counts.get(m.term) ?? 0) + inGap
      if (newTotal > cap) {
        blocked = true
        break
      }
    }
    if (blocked) continue
    for (const m of mustHaves) {
      counts.set(
        m.term,
        (counts.get(m.term) ?? 0) + bulletContainsMustHave(g.text, m),
      )
    }
    accepted.push(g)
  }
  return accepted
}

interface PendingBullet {
  index: number
  originalText: string
  targetWords: number
  lastAttempt?: {
    text: string
    words: number
    addedMustHaves: MustHave[]
    /** True if word-count is the reason this bullet is being retried. */
    wordCountFailed: boolean
    /** True if no new must-have keyword was introduced vs. original. */
    coverageFailed: boolean
  }
}

/** A single work-experience role inferred from the document order. The
 *  header is the company/title text immediately above the role's bullets;
 *  `lastBulletIndex` is where we'd insert a gap-fill bullet so it lands at
 *  the end of that role's bullet list. */
interface WorkRole {
  header: string
  bulletIndices: number[]
  lastBulletIndex: number
  avgWords: number
}

/** Group consecutive bullets in the 'work' section into roles. A role is
 *  the bullets that follow a block of non-bullet header lines. */
function identifyWorkRoles(paragraphs: DocxParagraph[]): WorkRole[] {
  const roles: WorkRole[] = []
  let currentHeader = ''
  let currentBullets: number[] = []
  let currentWordSum = 0

  const flush = () => {
    if (currentBullets.length === 0) return
    roles.push({
      header: currentHeader.trim(),
      bulletIndices: currentBullets.slice(),
      lastBulletIndex: currentBullets[currentBullets.length - 1],
      avgWords: Math.round(currentWordSum / currentBullets.length),
    })
    currentHeader = ''
    currentBullets = []
    currentWordSum = 0
  }

  for (const p of paragraphs) {
    if (p.section !== 'work' || !p.text.trim()) continue
    if (p.isBullet) {
      currentBullets.push(p.index)
      currentWordSum += wordCount(p.text)
    } else {
      // A non-bullet line after bullets ends the previous role.
      if (currentBullets.length > 0) flush()
      currentHeader += (currentHeader ? ' — ' : '') + p.text.trim()
    }
  }
  flush()
  return roles
}

interface GapFillBullet {
  text: string
  templateAfterIndex: number
  mustHaveTerm: string
  roleHeader: string
}

/** Generate new bullets to cover must-haves the rewrite pass couldn't place.
 *  Each new bullet is attributed to ONE existing role; we insert it after
 *  that role's last existing bullet so the formatting (numbering, indent,
 *  font) is cloned from a sibling. */
async function generateGapFillBullets(args: {
  missing: MustHave[]
  roles: WorkRole[]
  jobDescription: string
  fullResumeText: string
}): Promise<GapFillBullet[]> {
  if (!client) throw new Error('client not initialized')
  if (args.missing.length === 0 || args.roles.length === 0) return []

  const SYSTEM = `You write resume bullets to fill keyword gaps. For each missing must-have, generate ONE bullet that:
- Uses the must-have term verbatim (the exact spelling provided).
- Is attributed to ONE existing role from the candidate's resume (you'll pick by roleIndex).
- Matches the bullet style and word count of that role's existing bullets (±3 words of role avg).
- Reads like a plausible engineering accomplishment for someone in that role at that time.
- Does NOT contradict or duplicate any existing bullet's specific claims.
- Does NOT use fake metrics that would obviously stand out (avoid "200% improvement", "$5M saved" — keep it modest and plausible).

If two missing must-haves are in the same category and could share a bullet, generate ONE bullet that uses both rather than two separate ones.

If a missing must-have makes no plausible sense for any role (e.g. "React" for a candidate with zero frontend signals anywhere), SKIP it.

Output JSON: { "newBullets": [{ "missingTerm": "...", "roleIndex": 0, "text": "..." }, ...] }
No commentary.`

  const rolesBlock = args.roles
    .map(
      (r, i) =>
        `[role ${i}] ${r.header}\n  Existing bullets (avg ${r.avgWords} words):\n${r.bulletIndices
          .map(
            (idx, j) =>
              `    ${j + 1}. ${args.fullResumeText.split('\n').find((l) => l.trim() === `- ${l.trim().slice(2)}`) ? '' : ''}`,
          )
          .join('') || ''}`,
    )
    .join('\n')

  const missingBlock = args.missing
    .map((m) => `- "${m.term}" (${m.category}); aliases recognized: ${m.aliases.join(', ')}`)
    .join('\n')

  const response = await client.chat.completions.create({
    model: MODEL,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: SYSTEM },
      {
        role: 'user',
        content: [
          `=== JOB DESCRIPTION ===`,
          args.jobDescription,
          ``,
          `=== CANDIDATE'S FULL RESUME ===`,
          args.fullResumeText,
          ``,
          `=== ROLES (by index) ===`,
          args.roles
            .map((r, i) => `[role ${i}] ${r.header} (avg ${r.avgWords} words/bullet)`)
            .join('\n'),
          ``,
          `=== MISSING MUST-HAVES TO COVER ===`,
          missingBlock,
          ``,
          `For each missing must-have, decide if a plausible bullet exists; if yes, write it and assign roleIndex. Skip implausible ones.`,
        ].join('\n'),
      },
    ],
  })

  const raw = response.choices[0]?.message?.content ?? '{}'
  const parsed = JSON.parse(raw) as {
    newBullets?: Array<{ missingTerm: string; roleIndex: number; text: string }>
  }

  const out: GapFillBullet[] = []
  for (const b of parsed.newBullets ?? []) {
    const role = args.roles[b.roleIndex]
    if (!role) continue
    if (typeof b.text !== 'string' || b.text.trim().length === 0) continue
    // Validate the must-have alias actually appears in the generated text.
    const term = args.missing.find((m) => m.term === b.missingTerm)
    if (!term) continue
    const lower = b.text.toLowerCase()
    if (!term.aliases.some((a) => lower.includes(a.toLowerCase()))) continue
    // Validate word count is within ±5 of the role's avg.
    const w = wordCount(b.text)
    if (Math.abs(w - role.avgWords) > 5) continue

    out.push({
      text: b.text.trim(),
      templateAfterIndex: role.lastBulletIndex,
      mustHaveTerm: b.missingTerm,
      roleHeader: role.header,
    })
  }
  // Suppress runtime warning if unused var slips through.
  void rolesBlock
  return out
}

interface FabricatedReplacement {
  index: number
  text: string
  missingTerm: string
  originalText: string
}

/** Tier 3 eligibility: which existing bullets can be replaced with
 *  fabricated content? A bullet is eligible only if ALL four of the rules
 *  the user agreed to are satisfied. */
function findReplaceableBullets(
  paragraphs: DocxParagraph[],
  acceptedRewrites: Record<number, string>,
  mustHaves: MustHave[],
  roles: WorkRole[],
): number[] {
  const roleOfBullet = new Map<number, WorkRole>()
  for (const role of roles) {
    for (const idx of role.bulletIndices) roleOfBullet.set(idx, role)
  }
  // The most-recent bullet under the most-recent role is the first bullet of
  // the first role — top of page, top of resume's high-signal real estate.
  const mostRecentBullet = roles[0]?.bulletIndices[0]

  const eligible: number[] = []
  for (const p of paragraphs) {
    if (p.section !== 'work' || !p.isBullet) continue
    if (!p.text.trim()) continue
    // Rule 1: not already a tailored rewrite (rewrite found a keyword fit).
    if (p.index in acceptedRewrites) continue
    // Rule 2: original contains zero must-haves (already a generic line).
    if (mustHaves.some((m) => bulletContainsMustHave(p.text, m) === 1)) continue
    // Rule 3: not the only bullet in its role.
    const role = roleOfBullet.get(p.index)
    if (!role || role.bulletIndices.length <= 1) continue
    // Rule 4: not the most-recent bullet under the most-recent role.
    if (p.index === mostRecentBullet) continue
    eligible.push(p.index)
  }
  return eligible
}

/** Tier 3 generation: replace low-value bullets with fabricated content that
 *  covers still-missing must-haves. Uses scale figures from the candidate's
 *  resume (so metrics blend in) and refuses to ship anything with a stack
 *  mismatch or implausible word count. */
async function generateFabricatedReplacements(args: {
  missing: MustHave[]
  eligibleIndices: number[]
  paragraphs: DocxParagraph[]
  roles: WorkRole[]
  jobDescription: string
  fullResumeText: string
}): Promise<FabricatedReplacement[]> {
  if (!client) throw new Error('client not initialized')
  if (args.missing.length === 0 || args.eligibleIndices.length === 0) return []

  const roleOfBullet = new Map<number, WorkRole>()
  for (const role of args.roles) {
    for (const idx of role.bulletIndices) roleOfBullet.set(idx, role)
  }

  const eligibleBlock = args.eligibleIndices
    .map((idx) => {
      const p = args.paragraphs.find((p) => p.index === idx)
      const role = roleOfBullet.get(idx)
      if (!p || !role) return ''
      return `[${idx}] role: "${role.header}" (avg ${role.avgWords} words)\n  original (will be discarded): "${p.text}"`
    })
    .filter(Boolean)
    .join('\n\n')

  const missingBlock = args.missing
    .map((m) => `- "${m.term}" (${m.category}); aliases recognized: ${m.aliases.join(', ')}`)
    .join('\n')

  const SYSTEM = `You write FABRICATED resume bullets to cover JD must-haves that have no natural home in the candidate's actual experience. Each fabricated bullet REPLACES an existing low-value bullet entirely.

RULES:

1. Use the assigned must-have term verbatim (one of its aliases).
2. The bullet must describe a plausible system for the role's seniority and timeframe — coherent architecture, real-world tech stack pairings.
3. Match the role's average word count (±5 words).
4. Use scale figures and metrics that ALREADY APPEAR somewhere in the candidate's resume. If the resume mentions "10K+ users", "1K+ quotes/day", "99.8% success rate", "across 4 microservices" — reuse those exact scale shapes. Do NOT invent new ones.
5. Avoid round-number metrics that signal AI generation (50%, 100%, 2x, 10x). Prefer odd or asymmetric figures lifted from elsewhere in the resume.
6. Respect category constraints: frontend framework keywords belong in frontend activities; backend frameworks in backend activities; etc.

ASSIGNMENT:
For each missing must-have, pick ONE eligible bullet to replace (each eligible bullet can be used at most once). Try to match the role's domain — a frontend keyword should replace a bullet in a role that has other frontend signals if possible. If a missing must-have has no plausible home even among the eligible bullets, SKIP it.

Output JSON: { "replacements": [ { "originalIndex": <number>, "missingTerm": "<term>", "text": "<fabricated bullet>" } ] }
No commentary.`

  const response = await client.chat.completions.create({
    model: MODEL,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: SYSTEM },
      {
        role: 'user',
        content: [
          `=== JOB DESCRIPTION ===`,
          args.jobDescription,
          ``,
          `=== CANDIDATE'S FULL RESUME (for scale-figure reuse) ===`,
          args.fullResumeText,
          ``,
          `=== ELIGIBLE LOW-VALUE BULLETS (any can be replaced) ===`,
          eligibleBlock,
          ``,
          `=== STILL-MISSING MUST-HAVES ===`,
          missingBlock,
        ].join('\n'),
      },
    ],
  })

  const raw = response.choices[0]?.message?.content ?? '{}'
  const parsed = JSON.parse(raw) as {
    replacements?: Array<{
      originalIndex: number
      missingTerm: string
      text: string
    }>
  }

  const out: FabricatedReplacement[] = []
  const usedIndices = new Set<number>()
  for (const r of parsed.replacements ?? []) {
    if (typeof r.originalIndex !== 'number') continue
    if (!args.eligibleIndices.includes(r.originalIndex)) continue
    if (usedIndices.has(r.originalIndex)) continue
    if (typeof r.text !== 'string' || r.text.trim().length === 0) continue

    // Must contain the assigned missing term's alias.
    const term = args.missing.find((m) => m.term === r.missingTerm)
    if (!term) continue
    const lower = r.text.toLowerCase()
    if (!term.aliases.some((a) => lower.includes(a.toLowerCase()))) continue

    // Word count must be within ±5 of the role avg.
    const role = roleOfBullet.get(r.originalIndex)
    if (!role) continue
    const w = wordCount(r.text)
    if (Math.abs(w - role.avgWords) > 5) continue

    // Stack-mismatch guard (same rule as for regular rewrites).
    if (detectStackMismatch(r.text)) continue

    const originalText =
      args.paragraphs.find((p) => p.index === r.originalIndex)?.text ?? ''
    out.push({
      index: r.originalIndex,
      text: r.text.trim(),
      missingTerm: r.missingTerm,
      originalText,
    })
    usedIndices.add(r.originalIndex)
  }
  return out
}

async function callLlm(
  pending: PendingBullet[],
  jobDescription: string,
  fullResumeText: string,
  mustHaves: MustHave[],
  /** Must-haves the resume-as-a-whole still lacks. Hints the model toward
   *  filling these gaps when natural. */
  resumeGaps: MustHave[],
  /** Optional system prompt override; when set (persona mode), replaces the
   *  default keyword-swap system prompt with a stack-transmute one. */
  systemPromptOverride?: string,
  /** Block describing the active persona; appended to the user message in
   *  persona mode so the model has the full transmute target inline. */
  personaBlock?: string,
): Promise<Record<number, string>> {
  if (!client) throw new Error('client not initialized')

  const lines = pending.map((b) => {
    const lo = b.targetWords - WORD_COUNT_TOLERANCE
    const hi = b.targetWords + WORD_COUNT_TOLERANCE
    const base = `[${b.index}] TARGET: ${lo}–${hi} words (original was ${b.targetWords}).\nORIGINAL: ${b.originalText}`
    if (!b.lastAttempt) return base

    const last = b.lastAttempt
    const sections: string[] = [
      base,
      ``,
      `YOUR PREVIOUS ATTEMPT (${last.words} words):`,
      `"${last.text}"`,
      ``,
      `WHAT WENT WRONG:`,
    ]

    if (last.wordCountFailed) {
      const overBy = last.words - b.targetWords
      sections.push(
        overBy > 0
          ? `  • Word count: ${overBy} words OVER. Compress filler/redundant phrasing without dropping must-have keywords.`
          : `  • Word count: ${-overBy} words UNDER. Add a JD-relevant detail or restore context.`,
      )
    }
    if (last.coverageFailed) {
      const added = last.addedMustHaves.map((m) => m.term).join(', ') || 'nothing new'
      sections.push(
        `  • Recruiter Boolean check: this rewrite did NOT introduce a new must-have term. You added: [${added}]. The original already covered everything you added (or you added nothing).`,
      )
    }

    sections.push(
      ``,
      `RULES FOR THIS RETRY:`,
      `  - Keep every must-have term from your previous attempt that genuinely fits.`,
      `  - You MUST introduce at least one NEW must-have term that wasn't in the original bullet.`,
      `  - Compression tactics: "in order to" → "to"; "leveraging the use of" → "using"; trim redundant adjectives.`,
      `  - Metrics, employer, role, dates, verbs stay verbatim.`,
      `  - Stay within ${lo}–${hi} words.`,
    )

    return sections.join('\n')
  })

  const mustHavesBlock = mustHaves
    .map(
      (m) =>
        `  - "${m.term}" (${m.category}) — aliases recognized by the Boolean check: ${m.aliases.join(', ')}`,
    )
    .join('\n')

  const gapsBlock =
    resumeGaps.length > 0
      ? [
          ``,
          `=== RESUME-WIDE GAPS — these must-haves are not yet anywhere in the tailored resume; prefer placing them ===`,
          resumeGaps.map((m) => `  - ${m.term}`).join('\n'),
        ]
      : []

  const systemPrompt = systemPromptOverride ?? SYSTEM_PROMPT
  const userContent = personaBlock
    ? [
        `=== JOB DESCRIPTION ===`,
        jobDescription,
        ``,
        personaBlock,
        ``,
        `=== RECRUITER MUST-HAVES (Boolean AND-query terms — layered on after persona transmute) ===`,
        mustHavesBlock,
        ...gapsBlock,
        ``,
        `=== CANDIDATE'S FULL RESUME (context only) ===`,
        fullResumeText,
        ``,
        `=== BULLETS TO REWRITE ===`,
        lines.join('\n\n'),
      ].join('\n')
    : [
        `=== JOB DESCRIPTION ===`,
        jobDescription,
        ``,
        `=== RECRUITER MUST-HAVES (Boolean AND-query terms) ===`,
        `Your tailored resume will be checked against these. Each one is matched case-insensitively against its aliases. The resume passes only when ALL must-haves are present somewhere.`,
        mustHavesBlock,
        ...gapsBlock,
        ``,
        `=== CANDIDATE'S FULL RESUME (context only — do NOT rewrite paragraphs that aren't in the list below) ===`,
        fullResumeText,
        ``,
        `=== BULLETS TO REWRITE (work experience only) ===`,
        lines.join('\n\n'),
        ``,
        `Return the JSON object. Before responding: (a) each rewrite uses must-have aliases verbatim where possible; (b) word counts are in range; (c) prefer placing keywords that fill resume-wide gaps.`,
      ].join('\n')

  const response = await client.chat.completions.create({
    model: MODEL,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userContent },
    ],
  })

  const raw = response.choices[0]?.message?.content ?? '{}'
  const parsed = JSON.parse(raw) as { replacements?: Record<string, string> }
  const out: Record<number, string> = {}
  for (const [k, v] of Object.entries(parsed.replacements ?? {})) {
    out[Number(k)] = v
  }
  return out
}

const PERSONA_SYSTEM_PROMPT_BASE = `You are completely transmuting a resume's tech stack to fit a specific technical persona. The candidate is presenting as the persona described below. Rewrite every eligible bullet so the language, frameworks, databases, and infrastructure all align with that persona.

ACTIVITY MATCHING (REVISED — read carefully):
For each bullet, decide whether its underlying ACTIVITY is pure UI work, mixed backend/full-stack work, or pure server-side work.

- PURE UI bullets (React components, dashboards, layouts, design systems, user-facing flows): under a BACKEND persona → SKIP. Under a FRONTEND or FULLSTACK persona → transmute.
- PURE BACKEND bullets (REST APIs, queues, microservices, database work, async workers): under a FRONTEND persona → SKIP. Under BACKEND or FULLSTACK → transmute.
- MIXED bullets (Node.js + React, Express + Vue, "full-stack features", Socket.io + frontend): under any persona → TRANSMUTE the backend/runtime tech to the persona's language and framework. Keep frontend mentions where they describe real candidate work (a backend candidate can have shipped APIs that React consumed; that's fine to say).

CRITICAL — common runtime-vs-framework confusion:
- Node.js is a BACKEND runtime. It always transmutes under a backend persona, even if the same bullet mentions React.
- Express, NestJS, Fastify, Socket.io are BACKEND frameworks/libraries. Always transmute.
- React, Vue, Angular, Svelte are FRONTEND frameworks. Keep these in the original sentence if the candidate genuinely did frontend; don't claim React experience for someone who didn't have it.

Do NOT conflate "the bullet mentions React" with "this is a frontend bullet."

For bullets you DO transmute:
- Preserve verbs (architected, built, designed, processed).
- Preserve metrics (10K+, 99.8%, across 4 microservices), scale figures, and business context (quoting platform, financial workflows).
- Replace the LANGUAGE in the bullet with the persona's primary language.
- Replace any FRAMEWORK with the persona's framework of the matching layer (backend bullets get persona's backend framework; frontend bullets get persona's frontend framework).
- Replace any DATABASE with the persona's database (only if persona has one defined).
- Pull in 1–2 secondary tech terms from the persona naturally — don't keyword-stuff.

FRAMEWORK INFERENCE WHEN PERSONA LACKS ONE (CRITICAL):
If the persona lists a primary language but no backend (or frontend) framework, INFER the most popular framework for that language and use it as if the persona named it. Do NOT skip the bullet just because the persona's framework field is empty.

Inferred defaults to use:
- Go / Golang → Gin
- Python → FastAPI
- TypeScript / JavaScript → Express (backend) / React (frontend)
- Java → Spring Boot
- C# → ASP.NET
- Ruby → Rails
- Rust → Actix
- Kotlin → Ktor
- PHP → Laravel

Never drop a bullet for "no framework in persona" — infer one.

LANGUAGE-BOUND TOOLING SWAP (CRITICAL — failure mode):
A bullet may contain TESTING TOOLS, DEPENDENCY-INJECTION CONTAINERS, IDIOMS, or LIBRARIES that are specific to a language DIFFERENT from the persona's. Even if these aren't in the JD, they MUST be swapped to the persona-language equivalents — otherwise the resume reads as "Golang BOM engine using Java interfaces and Spring IoC", which is incoherent.

Common swaps when the persona's primary language is GO and the bullet contains:
- JUnit / JUnit 5 → Go's "testing" package + Testify
- Mockito → GoMock or Testify mock
- Spring Boot Test → Go integration tests with httptest
- Spring IoC / Spring DI → Go DI via Wire or constructor injection
- Spring Events → Go channels (or NATS / Kafka if event-bus)
- Spring Web Flux → Go goroutines + channels (or drop "reactive" if no clean parallel)
- Java interfaces / abstract classes → Go interfaces (Go has interfaces natively; just call them "Go interfaces")
- Maven / Gradle → Go modules

When the persona is PYTHON and the bullet contains:
- JUnit, Mockito, Spring → pytest, unittest.mock, FastAPI/Django patterns
- Java enums / interfaces → Python protocols or abstract base classes

When the persona is TYPESCRIPT/JAVASCRIPT and the bullet contains:
- JUnit / pytest → Jest or Vitest
- Spring / Django → Express middleware patterns

Apply the swap aggressively. If you see any framework / tooling word that doesn't belong to the persona's language, replace it. If unsure of a direct equivalent, drop the language-bound clause rather than leave it.

END-OF-SENTENCE STRAGGLER SWEEP (CRITICAL):
A common failure mode is transmuting the START of a sentence correctly but leaving a wrong-language word in a TRAILING clause. Examples that have shipped broken:
- "Enhanced CI/CD pipelines using Jenkins ... ensuring consistent delivery of Spring Boot microservices." → trailing "Spring Boot" must be "Golang" / persona's language.
- "Migrated inter-service communication to gRPC ... Node.js gateway consumed gRPC ..." → trailing "Node.js gateway" must be "Golang gateway" / persona's language.

Before responding, READ EACH REWRITE FROM END TO START. If any tech word names a language or framework that is NOT the persona's, replace it with the persona's equivalent or rewrite the clause to drop it. No language/framework name from a non-persona ecosystem should survive anywhere in any rewrite.

OUTPUT JSON: { "replacements": { "<index>": "<rewritten>" } }
Include only bullets you actually transmuted. Skip-decisions (activity doesn't match persona) are silent — omit from output.

WORD COUNT: stay within ±${WORD_COUNT_TOLERANCE} words of original.
COHERENCE: every transmuted bullet must describe a system that could actually exist. No frontend frameworks in REST API bullets, no Spring Batch in Python services, no AWS Lambda powering an LLM.
No commentary.`

/** Find the paragraph indices that contain comma-separated skill lists under
 *  the SKILLS section heading. These are the targets for must-have padding. */
function findSkillsParagraphs(paragraphs: DocxParagraph[]): number[] {
  const out: number[] = []
  let inSkills = false
  for (const p of paragraphs) {
    const trimmed = p.text.trim()
    if (!trimmed) continue
    if (/^(skills|technical skills|technologies)$/i.test(trimmed)) {
      inSkills = true
      continue
    }
    if (
      /^(work experience|experience|education|projects|certifications|awards|achievements|contact|summary|profile)$/i.test(
        trimmed,
      )
    ) {
      inSkills = false
      continue
    }
    if (inSkills && !p.isBullet) {
      const commas = (trimmed.match(/,/g) || []).length
      // Only paragraphs with ≥2 commas — that's a content line, not a subheading.
      if (commas >= 2) out.push(p.index)
    }
  }
  return out
}

/** Decide which existing skills paragraph a missing must-have should be
 *  appended to, based on what kind of skills the paragraph already contains. */
function classifySkillsParagraph(text: string): 'lang' | 'infra' {
  const lower = text.toLowerCase()
  const infraSignals = [
    'aws', 'gcp', 'azure', 'docker', 'kubernetes', 'jenkins', 'grafana',
    'kafka', 'redis', 'rabbitmq', 'opentelemetry', 'nginx', 'ci/cd', 'cloud',
    'terraform', 'github actions', 'datadog',
  ]
  const langSignals = [
    'java', 'python', 'javascript', 'typescript', 'react', 'spring', 'mysql',
    'postgresql', 'mongodb', 'graphql', 'rest', 'flask', 'django', 'fastapi',
    'go', 'golang',
  ]
  let infra = 0
  let lang = 0
  for (const s of infraSignals) if (lower.includes(s)) infra++
  for (const s of langSignals) if (lower.includes(s)) lang++
  return infra > lang ? 'infra' : 'lang'
}

const SKILLS_CATEGORY_TARGET: Record<string, 'lang' | 'infra'> = {
  language: 'lang',
  backend: 'lang',
  frontend: 'lang',
  database: 'lang',
  api: 'lang',
  methodology: 'lang',
  concept: 'lang',
  ai: 'lang',
  cache: 'infra',
  queue: 'infra',
  cloud: 'infra',
  orchestration: 'infra',
  iac: 'infra',
  observability: 'infra',
  cicd: 'infra',
}

/** Append a comma-separated list of new terms to an existing skills line,
 *  preserving any trailing punctuation. */
function appendToSkillsLine(originalText: string, newTerms: string[]): string {
  if (newTerms.length === 0) return originalText
  let text = originalText.trimEnd()
  let trailingPeriod = ''
  if (text.endsWith('.')) {
    trailingPeriod = '.'
    text = text.slice(0, -1)
  }
  text = text.replace(/,\s*$/, '')
  return `${text}, ${newTerms.join(', ')}${trailingPeriod}`
}

export async function rewriteResume(args: {
  paragraphs: DocxParagraph[]
  jobDescription: string
  /** When true, the model extracts a persona from the JD (+ optional title)
   *  and transmutes every bullet's tech stack to fit it. If extraction comes
   *  back low-confidence, we fall back to keyword-swap automatically. */
  personaMode?: boolean
  /** Optional role title to help persona extraction (e.g. pasted from the
   *  job posting). Only used if personaMode is true. */
  roleTitle?: string
}): Promise<RewriteResult> {
  if (!client) {
    throw new Error(
      'VITE_OPENAI_API_KEY is not set. Add it to .env.local and restart the dev server.',
    )
  }

  const bullets = args.paragraphs.filter(
    (p) => p.isBullet && p.section === 'work' && p.text.trim().length > 0,
  )
  const allowed = new Set(bullets.map((b) => b.index))

  const fullResumeText = args.paragraphs
    .filter((p) => p.text.trim().length > 0)
    .map((p) => (p.isBullet ? `  - ${p.text}` : p.text))
    .join('\n')

  // Extract recruiter must-haves once. ~$0.01, reused for every per-bullet
  // and final-resume Boolean check below.
  const mustHaves = await extractMustHaves(args.jobDescription)

  // If persona mode was requested, extract the persona BEFORE the rewrite
  // loop. If the extraction comes back low-confidence, fall back to
  // keyword-swap and surface the reason in the result.
  let persona: Persona | undefined
  let personaFallbackReason: string | undefined
  let mode: 'persona-transmute' | 'keyword-swap' = 'keyword-swap'
  if (args.personaMode) {
    const candidate = await extractPersona({
      jobDescription: args.jobDescription,
      roleTitle: args.roleTitle,
    })
    // Only fall back when there's no primary language at all. The confidence
    // field is informational only — the model is too conservative about
    // returning "high" (it routinely says "low" for clear JDs like
    // "Mandatory Skills: Golang"). A primary language is enough signal.
    if (!candidate.primaryLanguage) {
      personaFallbackReason =
        candidate.reasoning ||
        'The JD did not name a specific primary language. Falling back to keyword-swap mode.'
    } else {
      persona = candidate
      mode = 'persona-transmute'
    }
  }

  // Build the persona block injected into every rewrite call (only used
  // when mode === 'persona-transmute').
  let personaSystemPrompt: string | undefined
  let personaBlock: string | undefined
  if (mode === 'persona-transmute' && persona) {
    personaSystemPrompt = PERSONA_SYSTEM_PROMPT_BASE
    personaBlock = [
      `=== ACTIVE PERSONA (transmute all bullets to this stack) ===`,
      `Role type: ${persona.roleType}`,
      `Primary language: ${persona.primaryLanguage}`,
      persona.backendFramework ? `Backend framework: ${persona.backendFramework}` : '',
      persona.frontendFramework ? `Frontend framework: ${persona.frontendFramework}` : '',
      persona.database ? `Database: ${persona.database}` : '',
      persona.cloud ? `Cloud: ${persona.cloud}` : '',
      persona.secondaryTech.length > 0
        ? `Secondary tech (use 1–2 per bullet, max): ${persona.secondaryTech.join(', ')}`
        : '',
      ``,
      `Reasoning: ${persona.reasoning}`,
    ]
      .filter(Boolean)
      .join('\n')
  }

  let pending: PendingBullet[] = bullets.map((b) => ({
    index: b.index,
    originalText: b.text,
    targetWords: wordCount(b.text),
  }))

  const accepted: Record<number, string> = {}
  const addedByIndex: Record<number, string[]> = {}
  let attempts = 0

  const computeResumeGaps = (): MustHave[] => {
    const finalText = bullets
      .map((b) => accepted[b.index] ?? b.text)
      .join('\n')
    return checkBooleanQuery(finalText, mustHaves).missing
  }

  while (pending.length > 0 && attempts < MAX_ATTEMPTS) {
    attempts++
    const resumeGaps = computeResumeGaps()
    const responses = await callLlm(
      pending,
      args.jobDescription,
      fullResumeText,
      mustHaves,
      resumeGaps,
      personaSystemPrompt,
      personaBlock,
    )

    const nextPending: PendingBullet[] = []
    for (const b of pending) {
      if (!allowed.has(b.index)) continue
      const candidate = responses[b.index]
      if (!candidate) {
        // Model dropped this bullet — retry it next round.
        nextPending.push(b)
        continue
      }
      const got = wordCount(candidate)
      const addedMustHaves = bulletAddsMustHave(
        candidate,
        b.originalText,
        mustHaves,
      )

      // Hard sanity check: reject incoherent rewrites that mix a frontend
      // framework with strong backend activity (e.g. "Next.js gateway
      // consumed gRPC"). No amount of retry will fix this — the model
      // misclassified the bullet. Treat as a coverage failure so the retry
      // round can replace it; if it persists, the bullet stays original.
      const mismatch = detectStackMismatch(candidate)
      if (mismatch) {
        nextPending.push({
          ...b,
          lastAttempt: {
            text: candidate,
            words: got,
            addedMustHaves: [],
            wordCountFailed: false,
            coverageFailed: true,
          },
        })
        continue
      }

      const wordCountOk =
        Math.abs(got - b.targetWords) <= WORD_COUNT_TOLERANCE

      // In persona mode, accept a rewrite even if it didn't add a JD must-have
      // — as long as it transmuted the bullet's tech stack to the persona
      // (specifically: it now contains the persona's primary language and
      // the original didn't). Otherwise high-value swaps like "Spring Boot →
      // Gin" get dropped on coverage because Gin isn't a must-have.
      const personaLangAdded =
        mode === 'persona-transmute' &&
        persona !== undefined &&
        persona.primaryLanguage.length > 0 &&
        candidate.toLowerCase().includes(persona.primaryLanguage.toLowerCase()) &&
        !b.originalText
          .toLowerCase()
          .includes(persona.primaryLanguage.toLowerCase())

      const coverageOk = addedMustHaves.length >= 1 || personaLangAdded

      if (wordCountOk && coverageOk) {
        accepted[b.index] = candidate
        addedByIndex[b.index] = addedMustHaves.map((m) => m.term)
        continue
      }

      // If we're on the final attempt and word count is fine but no must-have
      // could be added, this bullet just has no natural Boolean keyword to
      // fit. Leave the original — no point burning more calls.
      const lastAttempt = attempts >= MAX_ATTEMPTS
      if (lastAttempt && wordCountOk && !coverageOk) continue

      nextPending.push({
        ...b,
        lastAttempt: {
          text: candidate,
          words: got,
          addedMustHaves,
          wordCountFailed: !wordCountOk,
          coverageFailed: !coverageOk,
        },
      })
    }
    pending = nextPending
  }

  // Last-resort acceptance: for bullets that hit MAX_ATTEMPTS without word-
  // count fit, accept the best attempt anyway IF (a) it added a must-have
  // that no other bullet has placed yet, AND (b) the bullet's surrounding
  // role still has missing must-haves. This trades a small line-wrap shift
  // for keyword coverage we'd otherwise lose entirely.
  const alignmentRelaxed: number[] = []
  const currentMissingTerms = new Set(
    checkBooleanQuery(
      bullets.map((b) => accepted[b.index] ?? b.text).join('\n'),
      mustHaves,
    ).missing.map((m) => m.term),
  )

  const stillPending: typeof pending = []
  for (const b of pending) {
    const last = b.lastAttempt
    if (!last) {
      stillPending.push(b)
      continue
    }
    // Would accepting this rewrite plug at least one currently-missing must-have?
    const fillsAGap = last.addedMustHaves.some((m) =>
      currentMissingTerms.has(m.term),
    )
    if (fillsAGap && last.coverageFailed === false) {
      accepted[b.index] = last.text
      addedByIndex[b.index] = last.addedMustHaves.map((m) => m.term)
      alignmentRelaxed.push(b.index)
      for (const m of last.addedMustHaves) currentMissingTerms.delete(m.term)
    } else {
      stillPending.push(b)
    }
  }

  const wordCountMismatches: RewriteResult['wordCountMismatches'] = stillPending
    .filter((b) => b.lastAttempt && b.lastAttempt.wordCountFailed)
    .map((b) => ({
      index: b.index,
      originalWords: b.targetWords,
      newWords: b.lastAttempt?.words ?? 0,
    }))

  // In persona transmute mode, the candidate's CORE STACK (language,
  // backend/frontend framework, database, cloud) should be allowed to appear
  // in every bullet — a Go engineer's resume says "Golang" everywhere by
  // design. Secondary tech still gets the cap (we don't want Docker in 9
  // bullets). Build the exempt set lowercased for case-insensitive matching.
  const exemptLowerTerms = new Set<string>()
  if (mode === 'persona-transmute' && persona) {
    if (persona.primaryLanguage)
      exemptLowerTerms.add(persona.primaryLanguage.toLowerCase())
    if (persona.backendFramework)
      exemptLowerTerms.add(persona.backendFramework.toLowerCase())
    if (persona.frontendFramework)
      exemptLowerTerms.add(persona.frontendFramework.toLowerCase())
    if (persona.database) exemptLowerTerms.add(persona.database.toLowerCase())
    if (persona.cloud) exemptLowerTerms.add(persona.cloud.toLowerCase())
  }

  // Enforce repetition cap: revert any rewrite that would push a must-have
  // past MAX_BULLETS_PER_KEYWORD occurrences in the final resume.
  const bulletShape = bullets.map((b) => ({
    index: b.index,
    originalText: b.text,
  }))
  const { final: cappedAccepted, bumped: initiallyBumped } =
    enforceRepetitionCap(
      bulletShape,
      accepted,
      mustHaves,
      MAX_BULLETS_PER_KEYWORD,
      exemptLowerTerms,
    )

  // Resurrection pass: a bumped rewrite carried away every keyword it added,
  // including ones that were genuinely missing globally. Walk the bumped
  // list and resurrect any rewrite that fills a currently-missing must-have.
  // We accept that this may push some keyword from 2 → 3 occurrences — the
  // alternative is leaving a recruiter must-have entirely uncovered.
  const postCapText = bullets
    .map((b) => cappedAccepted[b.index] ?? b.text)
    .join('\n')
  const missingNow = new Set(
    checkBooleanQuery(postCapText, mustHaves).missing.map((m) => m.term),
  )

  const resurrectedForCoverage: number[] = []
  const resurrectedRewrites: Record<number, string> = {}
  for (const idx of initiallyBumped) {
    if (missingNow.size === 0) break
    const rewriteText = accepted[idx] // pre-cap accepted map keeps the text
    if (!rewriteText) continue
    // Don't resurrect an incoherent rewrite even for coverage.
    if (detectStackMismatch(rewriteText)) continue

    // Which currently-missing must-haves does this rewrite plug?
    const filledTerms: string[] = []
    for (const m of mustHaves) {
      if (!missingNow.has(m.term)) continue
      if (bulletContainsMustHave(rewriteText, m)) filledTerms.push(m.term)
    }
    if (filledTerms.length === 0) continue

    resurrectedForCoverage.push(idx)
    resurrectedRewrites[idx] = rewriteText
    for (const term of filledTerms) missingNow.delete(term)
  }

  const finalAccepted = { ...cappedAccepted, ...resurrectedRewrites }
  const bumpedForRepetition = initiallyBumped.filter(
    (i) => !resurrectedForCoverage.includes(i),
  )

  // Mid-pipeline Boolean check before gap-fill: what's still missing now?
  const interimText = bullets
    .map((b) => finalAccepted[b.index] ?? b.text)
    .join('\n')
  const interimMatch = checkBooleanQuery(interimText, mustHaves)

  // Gap-fill pass: generate new bullets to cover the must-haves that no
  // existing bullet could absorb. Each new bullet clones a sibling's
  // structure (font, indent, numbering) via templateAfterIndex.
  const roles = identifyWorkRoles(args.paragraphs)
  const rawGapFills = await generateGapFillBullets({
    missing: interimMatch.missing,
    roles,
    jobDescription: args.jobDescription,
    fullResumeText,
  })
  // Apply the same cap to gap-fill bullets (no point introducing a 3rd
  // occurrence of jMeter via a new bullet).
  const generatedBullets = filterGapFillsByCap(
    bullets.map((b) => ({
      ...{ index: b.index } as { index?: number },
      originalText: b.text,
    })),
    finalAccepted,
    rawGapFills,
    mustHaves,
    MAX_BULLETS_PER_KEYWORD,
    exemptLowerTerms,
  )

  // Mid-point Boolean check after rewrites + gap-fill bullets.
  const postGapText =
    interimText + '\n' + generatedBullets.map((g) => g.text).join('\n')
  const postGapMissing = checkBooleanQuery(postGapText, mustHaves).missing

  // Tier 3: replace low-value bullets with fabricated content for any
  // must-haves that still couldn't find a home. Always runs (default-on).
  const fabricatedReplacements: RewriteResult['fabricatedReplacements'] = []
  if (postGapMissing.length > 0) {
    const eligibleForReplace = findReplaceableBullets(
      args.paragraphs,
      finalAccepted,
      mustHaves,
      roles,
    )
    const fabricated = await generateFabricatedReplacements({
      missing: postGapMissing,
      eligibleIndices: eligibleForReplace,
      paragraphs: args.paragraphs,
      roles,
      jobDescription: args.jobDescription,
      fullResumeText,
    })
    for (const fr of fabricated) {
      finalAccepted[fr.index] = fr.text
      fabricatedReplacements.push({
        index: fr.index,
        originalText: fr.originalText,
        newText: fr.text,
        mustHaveTerm: fr.missingTerm,
      })
    }
  }

  // Skills padding: append any must-haves still missing AFTER fabrication to
  // the appropriate skills paragraph. Uses categorization to split between
  // the "Languages/Frameworks/Databases" paragraph and the "Cloud/Infra"
  // paragraph. The result is silent on the actual flow (no LLM call), so
  // it's free and always runs.
  const beforeSkillsText = (() => {
    // Full-resume check: include every paragraph (work bullets + skills +
    // education + headers). Previously we only scanned work bullets — which
    // meant terms already in your Skills line counted as "missing".
    return args.paragraphs
      .map((p) => finalAccepted[p.index] ?? p.text)
      .join('\n') +
      '\n' +
      generatedBullets.map((g) => g.text).join('\n')
  })()
  const stillMissingAfterFabrication = checkBooleanQuery(
    beforeSkillsText,
    mustHaves,
  ).missing

  const skillsPaddingApplied: Record<number, string[]> = {}
  if (stillMissingAfterFabrication.length > 0) {
    const skillsIndices = findSkillsParagraphs(args.paragraphs)
    if (skillsIndices.length > 0) {
      // Categorize each skills paragraph by what it currently holds.
      const skillsCategory = new Map<number, 'lang' | 'infra'>()
      for (const idx of skillsIndices) {
        const p = args.paragraphs.find((x) => x.index === idx)
        if (p) skillsCategory.set(idx, classifySkillsParagraph(p.text))
      }
      // Find one paragraph for each category — fall back to either if only
      // one type of skills paragraph exists.
      let langIdx: number | undefined
      let infraIdx: number | undefined
      for (const [idx, cat] of skillsCategory.entries()) {
        if (cat === 'lang' && langIdx === undefined) langIdx = idx
        if (cat === 'infra' && infraIdx === undefined) infraIdx = idx
      }
      if (langIdx === undefined && infraIdx !== undefined) langIdx = infraIdx
      if (infraIdx === undefined && langIdx !== undefined) infraIdx = langIdx

      const langTerms: string[] = []
      const infraTerms: string[] = []
      for (const m of stillMissingAfterFabrication) {
        const target = SKILLS_CATEGORY_TARGET[m.category] ?? 'lang'
        if (target === 'lang') langTerms.push(m.term)
        else infraTerms.push(m.term)
      }

      if (langIdx !== undefined && langTerms.length > 0) {
        const p = args.paragraphs.find((x) => x.index === langIdx)
        if (p) {
          finalAccepted[langIdx] = appendToSkillsLine(p.text, langTerms)
          skillsPaddingApplied[langIdx] = langTerms
        }
      }
      if (infraIdx !== undefined && infraTerms.length > 0 && infraIdx !== langIdx) {
        const p = args.paragraphs.find((x) => x.index === infraIdx)
        if (p) {
          finalAccepted[infraIdx] = appendToSkillsLine(p.text, infraTerms)
          skillsPaddingApplied[infraIdx] = infraTerms
        }
      } else if (infraIdx === langIdx && infraTerms.length > 0 && langIdx !== undefined) {
        // Both categories share the same paragraph — combine into the
        // existing append we already did.
        const p = args.paragraphs.find((x) => x.index === langIdx)
        if (p) {
          const combined = [...langTerms, ...infraTerms]
          finalAccepted[langIdx] = appendToSkillsLine(p.text, combined)
          skillsPaddingApplied[langIdx] = combined
        }
      }
    }
  }

  // Final Boolean check covers the WHOLE resume, including skills paragraphs.
  // The previous bug — checking only work bullets — meant the Skills line was
  // invisible to coverage.
  const finalText =
    args.paragraphs
      .map((p) => finalAccepted[p.index] ?? p.text)
      .join('\n') +
    '\n' +
    generatedBullets.map((g) => g.text).join('\n')
  const booleanMatch = checkBooleanQuery(finalText, mustHaves)

  // Cleanup: scrub `addedByIndex` entries only for rewrites that stayed
  // bumped (not resurrected ones — they're back in the resume).
  for (const idx of bumpedForRepetition) delete addedByIndex[idx]

  return {
    replacements: finalAccepted,
    wordCountMismatches,
    attemptsUsed: attempts,
    addedByIndex,
    mustHaves,
    booleanMatch,
    generatedBullets,
    bumpedForRepetition,
    alignmentRelaxed,
    resurrectedForCoverage,
    fabricatedReplacements,
    mode,
    persona,
    personaFallbackReason,
    skillsPaddingApplied,
  }
}

// Re-export so the UI can render persona summaries without importing
// directly from persona.ts.
export { describePersona }
