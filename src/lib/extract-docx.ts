import PizZip from 'pizzip'

export interface DocxParagraph {
  index: number
  text: string
  /** True if this paragraph is a list/bullet item (has <w:numPr>). Resume
   *  bullet points always have this; titles, headings, dates, and skills
   *  lists do not. We use this to restrict LLM rewrites to bullets only. */
  isBullet: boolean
  /** Which top-level section this paragraph falls under, inferred from the
   *  nearest preceding section heading. "work" means WORK EXPERIENCE /
   *  EMPLOYMENT / etc. — the only section we want the LLM to touch. */
  section: 'work' | 'other'
}

export interface DocxDoc {
  zip: PizZip
  /** Path of the main document part inside the zip (e.g. "word/document.xml"
   *  or "word/document2.xml" — varies by author tool, so we resolve from the
   *  package relationships instead of hardcoding). */
  documentPath: string
  /** Original file bytes, kept so we can build modified outputs from a fresh
   *  zip each time (avoids cross-download state pollution). */
  originalBuffer: ArrayBuffer
  paragraphs: DocxParagraph[]
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const OFFICE_DOC_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

/** Look up the main document's path inside the zip via _rels/.rels. */
function resolveDocumentPath(zip: PizZip): string {
  const rels = zip.file('_rels/.rels')?.asText()
  if (rels) {
    const doc = new DOMParser().parseFromString(rels, 'application/xml')
    const nodes = doc.getElementsByTagName('Relationship')
    for (let i = 0; i < nodes.length; i++) {
      if (nodes[i].getAttribute('Type') === OFFICE_DOC_REL) {
        const target = nodes[i].getAttribute('Target') ?? ''
        // Targets are relative to the package root; strip any leading slash.
        return target.replace(/^\/+/, '')
      }
    }
  }
  // Sensible fallback when .rels is missing or malformed.
  if (zip.file('word/document.xml')) return 'word/document.xml'
  throw new Error('Invalid DOCX: could not locate the main document part.')
}

/**
 * Reads a DOCX file. Returns:
 *  - the unzipped archive (kept around so we can write changes back later)
 *  - the path of the main document part inside the zip
 *  - one entry per paragraph: its index in document order, and the concatenated
 *    text of all <w:t> runs inside it.
 */
export async function readDocx(file: File): Promise<DocxDoc> {
  const buffer = await file.arrayBuffer()
  const zip = new PizZip(buffer)

  const documentPath = resolveDocumentPath(zip)
  const documentXml = zip.file(documentPath)?.asText()
  if (!documentXml) {
    throw new Error(`Invalid DOCX: missing ${documentPath}`)
  }

  const doc = new DOMParser().parseFromString(documentXml, 'application/xml')
  const paraNodes = doc.getElementsByTagNameNS(W_NS, 'p')

  const paragraphs: DocxParagraph[] = []
  // Track section as we walk in document order. Anything before the first
  // recognized work heading is "other"; everything between a work heading
  // and the next non-work section heading is "work".
  let currentSection: 'work' | 'other' = 'other'
  for (let i = 0; i < paraNodes.length; i++) {
    const ts = paraNodes[i].getElementsByTagNameNS(W_NS, 't')
    let text = ''
    for (let j = 0; j < ts.length; j++) {
      text += ts[j].textContent ?? ''
    }
    const numPr = paraNodes[i].getElementsByTagNameNS(W_NS, 'numPr')
    const isBullet = numPr.length > 0

    // Update section based on heading text (only for non-bullet paragraphs —
    // a bullet whose text happens to say "EDUCATION" shouldn't change scope).
    if (!isBullet) {
      const kind = classifySectionHeading(text)
      if (kind === 'work' || kind === 'other') currentSection = kind
    }

    paragraphs.push({ index: i, text, isBullet, section: currentSection })
  }

  return { zip, documentPath, originalBuffer: buffer, paragraphs }
}

/** Decide if a heading line marks a section we care about. Returns null if
 *  the line isn't a heading at all. Heuristic: short line + uppercase /
 *  title-case + matches a known section noun. */
function classifySectionHeading(text: string): 'work' | 'other' | null {
  const t = text.trim()
  if (!t || t.length > 60) return null
  const norm = t.toLowerCase().replace(/[^a-z\s]/g, '').trim()
  if (!norm) return null
  const WORK = /^(work\s+experience|professional\s+experience|experience|employment|employment\s+history)$/
  const OTHER =
    /^(education|skills|projects|certifications|achievements|awards|publications|volunteer|languages|interests|references|summary|profile|objective|contact)$/
  if (WORK.test(norm)) return 'work'
  if (OTHER.test(norm)) return 'other'
  return null
}
