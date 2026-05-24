import PizZip from 'pizzip'

export interface DocxParagraph {
  index: number
  text: string
}

export interface DocxDoc {
  zip: PizZip
  /** Path of the main document part inside the zip (e.g. "word/document.xml"
   *  or "word/document2.xml" — varies by author tool, so we resolve from the
   *  package relationships instead of hardcoding). */
  documentPath: string
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
  for (let i = 0; i < paraNodes.length; i++) {
    const ts = paraNodes[i].getElementsByTagNameNS(W_NS, 't')
    let text = ''
    for (let j = 0; j < ts.length; j++) {
      text += ts[j].textContent ?? ''
    }
    paragraphs.push({ index: i, text })
  }

  return { zip, documentPath, paragraphs }
}
