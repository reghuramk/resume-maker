import PizZip from 'pizzip'
import type { DocxDoc } from './extract-docx'

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/**
 * Replaces text inside specific paragraphs of a DOCX, returning a new Blob.
 *
 * Why string surgery instead of DOMParser/XMLSerializer:
 *   A full XML round-trip in the browser silently rewrites attribute order,
 *   whitespace, and namespace declarations across the ENTIRE document — even
 *   paragraphs we never touched. Word reads these subtle changes and re-flows
 *   layout (tab stops shift, dates wrap to new lines, right-aligned positions
 *   move). To keep the file byte-identical outside the paragraphs we're
 *   actually editing, we operate on the raw XML string.
 *
 * Algorithm:
 *   1. Locate every <w:p>...</w:p> span in the XML string (regex; <w:p>
 *      doesn't nest in OOXML, so non-greedy matching is safe).
 *   2. Apply edits in REVERSE index order so character offsets of earlier
 *      paragraphs aren't invalidated by length changes in later ones.
 *   3. For each paragraph we're editing, replace the first <w:t>...</w:t>'s
 *      content with the new text and empty all subsequent <w:t> elements in
 *      that paragraph. The <w:tab/>, <w:rPr>, and other run properties stay
 *      intact, which preserves the original formatting (font, size, color).
 */
export function buildModifiedDocx(
  doc: DocxDoc,
  replacements: Record<number, string>,
): Blob {
  // Always work from the original bytes — never mutate doc.zip. This guarantees
  // every download starts from the pristine file, even if the user clicks
  // Download multiple times or a previous run failed mid-way.
  const zip = new PizZip(doc.originalBuffer)
  const xml = zip.file(doc.documentPath)?.asText()
  if (!xml) throw new Error(`Missing ${doc.documentPath}`)

  const paraSpans = findParagraphSpans(xml)

  const edits = Object.entries(replacements)
    .map(([k, v]) => [Number(k), v] as const)
    .filter(([idx]) => paraSpans[idx] !== undefined)
    .sort((a, b) => b[0] - a[0]) // reverse — apply later edits first

  let result = xml
  for (const [idx, newText] of edits) {
    const { start, end } = paraSpans[idx]
    const before = result.slice(0, start)
    const para = result.slice(start, end)
    const after = result.slice(end)
    result = before + replaceFirstWtContent(para, newText) + after
  }

  zip.file(doc.documentPath, result)
  return zip.generate({ type: 'blob', mimeType: DOCX_MIME })
}

function findParagraphSpans(
  xml: string,
): Array<{ start: number; end: number }> {
  const spans: Array<{ start: number; end: number }> = []
  // <w:p ...> ... </w:p>  OR  <w:p .../>  (self-closing, empty paragraph)
  const regex = /<w:p(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/w:p>)/g
  let m: RegExpExecArray | null
  while ((m = regex.exec(xml)) !== null) {
    spans.push({ start: m.index, end: m.index + m[0].length })
  }
  return spans
}

function replaceFirstWtContent(paraXml: string, newText: string): string {
  let first = true
  // <w:t [attrs]>content</w:t>  OR  <w:t [attrs]/>
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
