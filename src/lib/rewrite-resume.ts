import Anthropic from '@anthropic-ai/sdk'

const apiKey = import.meta.env.VITE_ANTHROPIC_API_KEY

const client = apiKey
  ? new Anthropic({ apiKey, dangerouslyAllowBrowser: true })
  : null

const SYSTEM_PROMPT = `You are a resume editor. Rewrite the candidate's resume to be tightly aligned with the target job description.

Rules:
- Keep every fact truthful — do not invent experience, jobs, titles, dates, or skills the candidate does not have.
- Reorder, reword, and re-emphasize existing content so the most JD-relevant points come first.
- Mirror keywords and phrasing from the job description where the candidate genuinely has the experience.
- Preserve the resume's structure (summary, experience, skills, education, etc.).
- Output the full rewritten resume as plain text only. No commentary, no markdown fences.`

export async function rewriteResume(args: {
  resumeText: string
  jobDescription: string
}): Promise<string> {
  if (!client) {
    throw new Error(
      'VITE_ANTHROPIC_API_KEY is not set. Add it to .env.local and restart the dev server.',
    )
  }

  const response = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 4096,
    system: SYSTEM_PROMPT,
    messages: [
      {
        role: 'user',
        content: `=== JOB DESCRIPTION ===\n${args.jobDescription}\n\n=== CURRENT RESUME ===\n${args.resumeText}\n\nRewrite the resume above to align with the job description, following all rules.`,
      },
    ],
  })

  const text = response.content
    .filter((block) => block.type === 'text')
    .map((block) => (block as { type: 'text'; text: string }).text)
    .join('\n')

  return text.trim()
}
