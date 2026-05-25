import OpenAI from 'openai'
import type { DocxParagraph } from './extract-docx'

const apiKey = import.meta.env.VITE_OPENAI_API_KEY

const client = apiKey
  ? new OpenAI({ apiKey, dangerouslyAllowBrowser: true })
  : null

const MODEL = 'gpt-4o'
const MAX_ATTEMPTS = 3
/** Acceptable word-count delta. Exactly equal is ideal but unattainable in
 *  practice; ±2 words almost never shifts line wrapping in Word for a normal
 *  resume bullet. */
const WORD_COUNT_TOLERANCE = 2

const SYSTEM_PROMPT = `You are a resume editor. Your job is to rewrite WORK EXPERIENCE bullet points so they directly target a specific job description.

PROCESS (do this in your head before responding):
1. Read the job description and identify 8–15 high-signal terms: technologies, methodologies, responsibilities, domain words.
2. For each bullet, identify which of those JD terms the candidate's existing experience genuinely covers.
3. Rewrite the bullet so the JD-relevant angle leads, and the JD's exact vocabulary is mirrored where the candidate's experience supports it.

JD-ALIGNMENT RULE:
Every rewrite must visibly pull at least one term/phrase straight from the job description (when the candidate's experience supports it). A rewrite with zero JD vocabulary is a failure — try harder or leave it unchanged.

WORD-COUNT RULE:
Each rewrite must be within ±${WORD_COUNT_TOLERANCE} words of the original. Words split by whitespace. "real-time" = 1 word. "AWS/GCP" = 1 word. Count before responding.

HONESTY RULE:
Never invent experience, technologies, metrics, or skills the candidate does not have. Only reframe their existing content — keywords get pulled from the JD only when the candidate genuinely has that experience.

OUTPUT:
Return JSON: { "replacements": { "<index>": "<rewritten text>", ... } }
No commentary.`

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
): Promise<Record<number, string>> {
  if (!client) throw new Error('client not initialized')

  const lines = pending.map((b) => {
    const lo = b.targetWords - WORD_COUNT_TOLERANCE
    const hi = b.targetWords + WORD_COUNT_TOLERANCE
    const base = `[${b.index}] TARGET: ${lo}–${hi} words (original was ${b.targetWords}).\nORIGINAL: ${b.originalText}`
    if (b.lastAttempt) {
      return `${base}\nYOUR PREVIOUS ATTEMPT had ${b.lastAttempt.words} words (outside the ${lo}–${hi} range): ${b.lastAttempt.text}\nTry again. Count carefully and pull JD vocabulary.`
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
        content: `=== JOB DESCRIPTION ===\n${jobDescription}\n\n=== BULLETS TO REWRITE ===\n${lines.join('\n\n')}\n\nReturn the JSON object as specified. Verify word counts before responding.`,
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
    const responses = await callLlm(pending, args.jobDescription)

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
