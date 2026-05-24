import OpenAI from 'openai'
import type { DocxParagraph } from './extract-docx'

const apiKey = import.meta.env.VITE_OPENAI_API_KEY

const client = apiKey
  ? new OpenAI({ apiKey, dangerouslyAllowBrowser: true })
  : null

const SYSTEM_PROMPT = `You are a resume editor. You will receive a list of numbered paragraphs from a resume and a job description. Rewrite ONLY the bullet points under work experience that describe accomplishments — leave everything else exactly as-is.

WHAT TO REWRITE (include in your output):
- Bullet points under each WORK EXPERIENCE role that describe what the candidate did/achieved.
- Optional: the summary/objective paragraph at the top if there is one.

WHAT NOT TO REWRITE (do NOT include in output):
- Section headings (e.g. "WORK EXPERIENCE", "EDUCATION", "SKILLS")
- Company names
- Job titles
- Dates and locations
- Skills list, education entries, project names
- Empty paragraphs, page numbers, contact info

WORD-COUNT RULE (CRITICAL):
- Every rewritten paragraph MUST contain EXACTLY the same number of words as the original. Count carefully. Words separated by whitespace.
- If you cannot match the word count exactly, return the original text unchanged for that paragraph (omit it from your output).

HONESTY RULE:
- Never invent experience, technologies, metrics, or skills the candidate does not have. Reframe existing content to highlight JD-relevant angles.

OUTPUT FORMAT:
Return a JSON object: { "replacements": { "<paragraph_index>": "<new text>", ... } }
Only include entries for paragraphs you are rewriting. No commentary.`

export interface RewriteResult {
  replacements: Record<number, string>
  wordCountMismatches: Array<{
    index: number
    originalWords: number
    newWords: number
  }>
}

function wordCount(s: string): number {
  return s.trim().split(/\s+/).filter(Boolean).length
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

  // Skip empty paragraphs to keep the prompt focused.
  const nonEmpty = args.paragraphs.filter((p) => p.text.trim().length > 0)
  const numbered = nonEmpty
    .map((p) => `[${p.index}] (${wordCount(p.text)} words) ${p.text}`)
    .join('\n')

  const response = await client.chat.completions.create({
    model: 'gpt-4o-mini',
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      {
        role: 'user',
        content: `=== JOB DESCRIPTION ===\n${args.jobDescription}\n\n=== RESUME PARAGRAPHS ===\n${numbered}\n\nReturn the JSON object as specified.`,
      },
    ],
  })

  const raw = response.choices[0]?.message?.content ?? '{}'
  const parsed = JSON.parse(raw) as { replacements?: Record<string, string> }
  const replacementsIn = parsed.replacements ?? {}

  const replacements: Record<number, string> = {}
  const mismatches: RewriteResult['wordCountMismatches'] = []

  for (const [idxStr, newText] of Object.entries(replacementsIn)) {
    const idx = Number(idxStr)
    const original = args.paragraphs.find((p) => p.index === idx)
    if (!original) continue
    const origWc = wordCount(original.text)
    const newWc = wordCount(newText)
    // Enforce same word count — drop rewrites that don't match.
    if (origWc !== newWc) {
      mismatches.push({ index: idx, originalWords: origWc, newWords: newWc })
      continue
    }
    replacements[idx] = newText
  }

  return { replacements, wordCountMismatches: mismatches }
}
