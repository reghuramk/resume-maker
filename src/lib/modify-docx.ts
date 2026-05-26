import PizZip from 'pizzip'
import type { DocxDoc } from './extract-docx'

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/** Insert a new bullet by cloning the structure of an existing bullet
 *  (`templateAfterIndex`) and placing it immediately after that bullet.
 *  The new bullet inherits every paragraph property (numbering, indent,
 *  font, color) from the template, so it visually aligns with siblings. */
export interface DocxInsertion {
  templateAfterIndex: number
  text: string
}

export function buildModifiedDocx(
  doc: DocxDoc,
  replacements: Record<number, string>,
  insertions: DocxInsertion[] = [],
): Blob {
  const zip = new PizZip(doc.originalBuffer)
  const xml = zip.file(doc.documentPath)?.asText()
  if (!xml) throw new Error(`Missing ${doc.documentPath}`)

  const paraSpans = findParagraphSpans(xml)

  type Op =
    | { kind: 'edit'; idx: number; text: string }
    | { kind: 'insert'; afterIdx: number; text: string }

  const ops: Op[] = [
    ...Object.entries(replacements)
      .map(([k, v]) => ({ kind: 'edit' as const, idx: Number(k), text: v }))
      .filter((o) => paraSpans[o.idx] !== undefined),
    ...insertions
      .filter((i) => paraSpans[i.templateAfterIndex] !== undefined)
      .map((i) => ({
        kind: 'insert' as const,
        afterIdx: i.templateAfterIndex,
        text: i.text,
      })),
  ]

  // Apply highest-index ops first so earlier ops aren't invalidated by
  // length changes downstream. At the same index, run inserts AFTER the
  // edit on the same paragraph (template gets text-swapped first, then
  // cloned — they share structure either way).
  ops.sort((a, b) => {
    const ai = a.kind === 'edit' ? a.idx : a.afterIdx
    const bi = b.kind === 'edit' ? b.idx : b.afterIdx
    if (ai !== bi) return bi - ai
    if (a.kind === 'edit' && b.kind === 'insert') return -1
    if (a.kind === 'insert' && b.kind === 'edit') return 1
    return 0
  })

  let result = xml
  for (const op of ops) {
    if (op.kind === 'edit') {
      const { start, end } = paraSpans[op.idx]
      const para = result.slice(start, end)
      result =
        result.slice(0, start) +
        replaceFirstWtContent(para, op.text) +
        result.slice(end)
    } else {
      // Clone the template bullet's full <w:p>...</w:p> structure, replace
      // its text, splice in right after the template. Numbering, indent,
      // font, and color all carry over because the cloned XML keeps every
      // <w:pPr> / <w:rPr> intact.
      const { start, end } = paraSpans[op.afterIdx]
      const template = result.slice(start, end)
      const clone = replaceFirstWtContent(template, op.text)
      result = result.slice(0, end) + clone + result.slice(end)
    }
  }

  zip.file(doc.documentPath, result)
  return zip.generate({ type: 'blob', mimeType: DOCX_MIME })
}

function findParagraphSpans(
  xml: string,
): Array<{ start: number; end: number }> {
  const spans: Array<{ start: number; end: number }> = []
  const regex = /<w:p(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/w:p>)/g
  let m: RegExpExecArray | null
  while ((m = regex.exec(xml)) !== null) {
    spans.push({ start: m.index, end: m.index + m[0].length })
  }
  return spans
}

function replaceFirstWtContent(paraXml: string, newText: string): string {
  let first = true
  const regex = /<w:t(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/w:t>)/g
  return paraXml.replace(regex, () => {
    if (first) {
      first = false
      return `<w:t xml:space="preserve">${escapeXml(newText)}</w:t>`
    }
    return `<w:t></w:t>`
  })
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
