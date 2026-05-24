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

  // Give layout a tick to settle before opening print dialog.
  setTimeout(() => win.print(), 200)
}
