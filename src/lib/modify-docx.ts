import type { DocxDoc } from './extract-docx'

/**
 * Replaces text inside specific paragraphs of a DOCX, returning a new Blob.
 *
 * How it preserves formatting:
 *  - A paragraph (<w:p>) can contain many runs (<w:r>) and many text nodes
 *    (<w:t>) — each run holds its own formatting (font, size, bold, color).
 *  - When we replace a paragraph's text, we put the entire new text into the
 *    first <w:t> and empty all the others. This preserves the paragraph's
 *    primary run formatting (the first run's properties usually represent the
 *    paragraph's main style). Inline emphasis (e.g. one bold word inside a
 *    bullet) is lost — for resume bullets this is almost always acceptable.
 *  - Untouched paragraphs are not modified at all, so headings, company names,
 *    job titles, and dates keep their original formatting verbatim.
 */
export function buildModifiedDocx(
  doc: DocxDoc,
  replacements: Record<number, string>,
): Blob {
  const documentXml = doc.zip.file(doc.documentPath)?.asText()
  if (!documentXml) throw new Error(`Missing ${doc.documentPath}`)

  const parser = new DOMParser()
  const xmlDoc = parser.parseFromString(documentXml, 'application/xml')
  const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const paraNodes = xmlDoc.getElementsByTagNameNS(wNS, 'p')

  for (const [indexStr, newText] of Object.entries(replacements)) {
    const idx = Number(indexStr)
    const p = paraNodes[idx]
    if (!p) continue

    const ts = p.getElementsByTagNameNS(wNS, 't')
    if (ts.length === 0) continue

    // Put all the new text in the first <w:t>, blank the rest.
    ts[0].textContent = newText
    // Preserve leading/trailing spaces in the first run.
    ts[0].setAttribute('xml:space', 'preserve')
    for (let i = 1; i < ts.length; i++) {
      ts[i].textContent = ''
    }
  }

  const serializer = new XMLSerializer()
  const newXml = serializer.serializeToString(xmlDoc)
  doc.zip.file(doc.documentPath, newXml)

  const out = doc.zip.generate({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })
  return out
}
