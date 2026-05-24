import { renderAsync } from 'docx-preview'

/**
 * Opens the modified DOCX in a new window, renders it to HTML using docx-preview
 * (preserves fonts, layout, spacing), and triggers the browser print dialog.
 * The user picks "Save as PDF" as the destination.
 *
 * This route is intentional: pure-JS DOCX→PDF rendering (without going through
 * the browser print pipeline) loses font fidelity. The browser's print engine
 * embeds the actual rendered fonts, so the PDF looks identical to what
 * docx-preview rendered on screen.
 */
export async function exportDocxAsPdf(docxBlob: Blob): Promise<void> {
  const win = window.open('', '_blank')
  if (!win) throw new Error('Popup blocked — allow popups and try again.')

  win.document.write(`
    <!doctype html>
    <html>
      <head>
        <title>Tailored Resume</title>
        <style>
          body { margin: 0; }
          .docx-wrapper { background: white !important; padding: 0 !important; }
          @media print {
            @page { margin: 0; }
            body { margin: 0; }
          }
        </style>
      </head>
      <body><div id="render"></div></body>
    </html>
  `)
  win.document.close()

  const target = win.document.getElementById('render')!
  await renderAsync(docxBlob, target, undefined, {
    inWrapper: false,
    ignoreWidth: false,
    ignoreHeight: false,
  })

  // Word uses Symbol/Wingdings private-use codepoints for list bullets and
  // tick marks. Browsers don't ship those fonts, so the characters render as
  // missing-glyph tofu boxes. Swap them for standard Unicode equivalents that
  // every font supports.
  replaceSymbolPuaChars(win.document.body)

  // Give layout a tick to settle before opening print dialog.
  setTimeout(() => win.print(), 200)
}

const PUA_REPLACEMENTS: Record<string, string> = {
  '': '•', // Symbol bullet → •
  '': '■', // Symbol black square → ■
  '': '□', // Symbol white square → □
  '': '✓', // Wingdings check → ✓
  '': '✓', // Wingdings check (alt) → ✓
  '': '►', // Wingdings arrow → ►
  '': '→', // Wingdings arrow right → →
}

function replaceSymbolPuaChars(root: Node) {
  const doc = root.ownerDocument
  if (!doc) return
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let node: Node | null
  while ((node = walker.nextNode())) {
    const text = node.nodeValue
    if (!text) continue
    let replaced = text
    for (const [from, to] of Object.entries(PUA_REPLACEMENTS)) {
      if (replaced.includes(from)) {
        replaced = replaced.split(from).join(to)
      }
    }
    if (replaced !== text) node.nodeValue = replaced
  }
}
