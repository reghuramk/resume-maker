import PizZip from 'pizzip'

export interface DocxParagraph {
  index: number
  text: string
}

export interface DocxDoc {
  zip: PizZip
  paragraphs: DocxParagraph[]
}

/**
 * Reads a DOCX file. Returns:
 *  - the unzipped archive (kept around so we can write changes back later)
 *  - one entry per paragraph: its index in document order, and the concatenated
 *    text of all <w:t> runs inside it.
 *
 * Paragraph order in the array matches their on-page order.
 */
export async function readDocx(file: File): Promise<DocxDoc> {
  const buffer = await file.arrayBuffer()
  const zip = new PizZip(buffer)
  const documentXml = zip.file('word/document.xml')?.asText()
  if (!documentXml) {
    throw new Error('Invalid DOCX: missing word/document.xml')
  }

  const doc = new DOMParser().parseFromString(documentXml, 'application/xml')
  // w:p = paragraph, w:t = text run content
  const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const paraNodes = doc.getElementsByTagNameNS(wNS, 'p')

  const paragraphs: DocxParagraph[] = []
  for (let i = 0; i < paraNodes.length; i++) {
    const ts = paraNodes[i].getElementsByTagNameNS(wNS, 't')
    let text = ''
    for (let j = 0; j < ts.length; j++) {
      text += ts[j].textContent ?? ''
    }
    paragraphs.push({ index: i, text })
  }

  return { zip, paragraphs }
}
