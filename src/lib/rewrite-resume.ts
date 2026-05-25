import OpenAI from 'openai'
import type { DocxParagraph } from './extract-docx'

const apiKey = import.meta.env.VITE_OPENAI_API_KEY

const client = apiKey
  ? new OpenAI({ apiKey, dangerouslyAllowBrowser: true })
  : null

const MODEL = 'gpt-4o'
const MAX_ATTEMPTS = 3
const WORD_COUNT_TOLERANCE = 3

const SYSTEM_PROMPT = `You are an ATS-optimization expert. Rewrite WORK EXPERIENCE bullets by SWAPPING tech keywords within their semantic categories — never across categories. The result must read naturally to a human engineer.

CATEGORY RULE (THE MOST IMPORTANT RULE):
Every swap must stay within one category. A category is the role the technology plays in a system.

VALID swap categories (swap within, never across):
- Programming language: Java ↔ TypeScript ↔ Python ↔ C#
- Backend framework: Spring Boot ↔ Express ↔ NestJS ↔ FastAPI
- Frontend framework: Vue ↔ React ↔ Angular
- Database: PostgreSQL ↔ MySQL ↔ DynamoDB ↔ MongoDB
- Cache: Redis ↔ Memcached
- Message queue / events: Kafka ↔ SQS ↔ EventBridge ↔ RabbitMQ
- Cloud compute: EC2 ↔ ECS ↔ AWS Lambda ↔ Kubernetes pods
- Orchestration: Kubernetes ↔ ECS ↔ Nomad
- IaC: Terraform ↔ CloudFormation ↔ Pulumi
- Observability: Grafana ↔ Datadog ↔ CloudWatch ↔ OpenTelemetry
- CI/CD: GitHub Actions ↔ GitLab CI ↔ Jenkins
- API style: REST ↔ GraphQL ↔ gRPC
- LLM/AI tooling: OpenAI ↔ Anthropic ↔ LangChain ↔ RAG pipelines

INVALID swaps (these are forbidden — they produce nonsense):
- "OpenAI-powered assistants" ✗→ "AWS Lambda-powered assistants" (OpenAI is an LLM provider; Lambda is compute)
- "LangChain pipelines" ✗→ "DynamoDB pipelines" (AI framework ≠ database)
- "RAG pipelines" ✗→ "Kubernetes pipelines" (AI pattern ≠ orchestrator)
- "Spring Web Flux reactive pipelines" ✗→ "Lambda reactive pipelines" (reactive framework ≠ FaaS runtime)
- "Connection pooling" ✗→ "SQS" (DB technique ≠ message queue)

If a bullet's tech has NO category-parallel in the JD, LEAVE THE BULLET ALONE — omit it from your output.

DISTRIBUTION RULE (HARD LIMIT):
Before responding, count how many times each JD keyword appears across all your rewrites combined. Any single keyword may appear at most TWICE across the entire resume. If your draft has "AWS Lambda" in 4 bullets, you've failed — go back and use other categories' keywords instead, or leave the extra bullets unchanged.

STRATEGY:
1. From the JD, extract concrete keywords grouped by category (languages, frameworks, databases, queues, compute, observability, AI tooling, etc.).
2. From the resume, map each bullet's tech terms to their categories.
3. Match category-to-category between resume bullet and JD: only swap if both sides offer the same category.
4. Distribute the JD keywords so each shows up ≤2 times. Prefer placing each major JD keyword once in its most-natural-fitting bullet.
5. Skip bullets that have no category-parallel swap available. Don't force.

WORKED EXAMPLE:

JD: Node.js, TypeScript, DynamoDB, AWS Lambda, Serverless, SQS, EventBridge, REST APIs.

Original: "Architected a quoting platform in Spring Boot, building REST APIs for 10K+ daily users, processing 1K+ quotes/day via connection pooling and Spring Web Flux reactive pipelines."

Good rewrite (categories: framework, REST, message-pipeline):
"Architected a quoting platform on Node.js and TypeScript, building REST APIs for 10K+ daily users and processing 1K+ quotes/day via EventBridge and SQS reactive pipelines."
- Spring Boot (backend framework) → Node.js + TypeScript ✓ category match
- REST APIs → REST APIs ✓ already aligned
- Spring Web Flux (reactive messaging) → EventBridge + SQS ✓ category match (event-driven messaging)
- "connection pooling" removed (no parallel in JD, no forced swap)

INTEGRITY RULES (don't violate):
- Every metric, percentage, scale figure, dollar amount stays verbatim.
- Verbs (architected, built, designed, etc.) and business context (quoting platform, financial workflows) stay.
- Employer, role title, dates never change.
- New keyword must appear in the JD text verbatim. No synonyms, no invented phrases.
- Stay within ±${WORD_COUNT_TOLERANCE} words. "real-time" = 1 word.

OUTPUT:
Return JSON: { "replacements": { "<index>": "<rewritten text>", ... } }

Before responding, run this internal checklist:
1. Did I do any cross-category swap? If yes, undo it and omit that bullet.
2. Does any single keyword appear >2 times in my output? If yes, redistribute or drop.
3. Does every rewrite read naturally to an engineer? If not, fix it or omit.

No commentary outside the JSON.`

export interface RewriteResult {
  replacements: Record<number, string>
  wordCountMismatches: Array<{
    index: number
    originalWords: number
    newWords: number
  }>
  attemptsUsed: number
}

function wordCount(s: string): number {
  return s.trim().split(/\s+/).filter(Boolean).length
}

interface PendingBullet {
  index: number
  originalText: string
  targetWords: number
  lastAttempt?: { text: string; words: number }
}

async function callLlm(
  pending: PendingBullet[],
  jobDescription: string,
  fullResumeText: string,
): Promise<Record<number, string>> {
  if (!client) throw new Error('client not initialized')

  const lines = pending.map((b) => {
    const lo = b.targetWords - WORD_COUNT_TOLERANCE
    const hi = b.targetWords + WORD_COUNT_TOLERANCE
    const base = `[${b.index}] TARGET: ${lo}–${hi} words (original was ${b.targetWords}).\nORIGINAL: ${b.originalText}`
    if (b.lastAttempt) {
      return `${base}\nYOUR PREVIOUS ATTEMPT had ${b.lastAttempt.words} words (outside the ${lo}–${hi} range): ${b.lastAttempt.text}\nTry again. Pull JD vocabulary verbatim and count carefully.`
    }
    return base
  })

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
          `=== CANDIDATE'S FULL RESUME (context only — do NOT rewrite paragraphs that aren't in the list below) ===`,
          fullResumeText,
          ``,
          `=== BULLETS TO REWRITE (work experience only) ===`,
          lines.join('\n\n'),
          ``,
          `Return the JSON object as specified. Before responding, verify (a) each rewrite contains JD vocabulary copied verbatim from the JD text, and (b) word counts are inside the allowed range.`,
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

  // Only WORK EXPERIENCE bullets — never touch Education, Skills, Projects,
  // or anything else, even if those sections contain bullet lists.
  const bullets = args.paragraphs.filter(
    (p) => p.isBullet && p.section === 'work' && p.text.trim().length > 0,
  )
  const allowed = new Set(bullets.map((b) => b.index))

  // Plain-text dump of the entire resume — gives the model context on which
  // tech/skills the candidate has actually used across all roles, so it can
  // confidently pull JD vocabulary into bullets that are evidenced elsewhere.
  const fullResumeText = args.paragraphs
    .filter((p) => p.text.trim().length > 0)
    .map((p) => (p.isBullet ? `  - ${p.text}` : p.text))
    .join('\n')

  // Start with every bullet pending.
  let pending: PendingBullet[] = bullets.map((b) => ({
    index: b.index,
    originalText: b.text,
    targetWords: wordCount(b.text),
  }))

  const accepted: Record<number, string> = {}
  let attempts = 0

  while (pending.length > 0 && attempts < MAX_ATTEMPTS) {
    attempts++
    const responses = await callLlm(pending, args.jobDescription, fullResumeText)

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
      if (Math.abs(got - b.targetWords) <= WORD_COUNT_TOLERANCE) {
        accepted[b.index] = candidate
      } else {
        nextPending.push({
          ...b,
          lastAttempt: { text: candidate, words: got },
        })
      }
    }
    pending = nextPending
  }

  // Anything still pending after max attempts — give up cleanly.
  const wordCountMismatches: RewriteResult['wordCountMismatches'] = pending.map(
    (b) => ({
      index: b.index,
      originalWords: b.targetWords,
      newWords: b.lastAttempt?.words ?? 0,
    }),
  )

  return {
    replacements: accepted,
    wordCountMismatches,
    attemptsUsed: attempts,
  }
}
