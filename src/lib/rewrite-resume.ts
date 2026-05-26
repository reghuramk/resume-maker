import OpenAI from 'openai'
import type { DocxParagraph } from './extract-docx'
import {
  extractMustHaves,
  checkBooleanQuery,
  bulletAddsMustHave,
  bulletContainsMustHave,
  type MustHave,
  type BooleanMatch,
} from './ats-scorer'

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
NEGATIVE EXAMPLES (DO NOT DO THESE)
==============================

✗ "Python FastAPI data processing pipeline using Spring Batch"
   → Spring Batch is Java. You cannot use it from Python. Cross-stack nonsense.

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

Before responding, re-run the STEP 4 checklist one more time. If anything fails, fix or drop.

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
}

function wordCount(s: string): number {
  return s.trim().split(/\s+/).filter(Boolean).length
}

/** Walk accepted rewrites in document order and revert any that would cause
 *  a must-have keyword to appear in more than `cap` bullets. A rewrite is
 *  blocked only if it (a) ADDS a new occurrence of some keyword AND (b) that
 *  would push the keyword's total past the cap. Rewrites that don't change
 *  a keyword's count are always allowed; rewrites that REMOVE a keyword
 *  somewhere can free budget for other rewrites later. */
function enforceRepetitionCap(
  bullets: Array<{ index: number; originalText: string }>,
  acceptedRewrites: Record<number, string>,
  mustHaves: MustHave[],
  cap: number,
): { final: Record<number, string>; bumped: number[] } {
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
 *  subtract). Drop any gap-fill that would push a must-have past the cap. */
function filterGapFillsByCap(
  bullets: Array<{ originalText: string }>,
  acceptedRewrites: Record<number, string>,
  gapFills: GapFillBullet[],
  mustHaves: MustHave[],
  cap: number,
): GapFillBullet[] {
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

async function callLlm(
  pending: PendingBullet[],
  jobDescription: string,
  fullResumeText: string,
  mustHaves: MustHave[],
  /** Must-haves the resume-as-a-whole still lacks. Hints the model toward
   *  filling these gaps when natural. */
  resumeGaps: MustHave[],
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

  const response = await client.chat.completions.create({
    model: MODEL,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      {
        role: 'user',
        content: [
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
        ].join('\n'),
      },
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

export async function rewriteResume(args: {
  paragraphs: DocxParagraph[]
  jobDescription: string
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

  let pending: PendingBullet[] = bullets.map((b) => ({
    index: b.index,
    originalText: b.text,
    targetWords: wordCount(b.text),
  }))

  const accepted: Record<number, string> = {}
  const addedByIndex: Record<number, string[]> = {}
  let attempts = 0

  // Helper: given current accepted rewrites + untouched bullets, which
  // must-haves does the resume-as-a-whole still lack? Used to bias each
  // retry round toward filling resume-wide gaps instead of duplicating.
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

      const wordCountOk =
        Math.abs(got - b.targetWords) <= WORD_COUNT_TOLERANCE
      const coverageOk = addedMustHaves.length >= 1

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
  )

  // Final Boolean check including the surviving gap-fill bullets.
  const finalText =
    interimText + '\n' + generatedBullets.map((g) => g.text).join('\n')
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
  }
}
