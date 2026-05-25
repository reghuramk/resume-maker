import { useState } from 'react'
import type { DragEvent, ChangeEvent } from 'react'
import { readDocx, type DocxDoc } from './lib/extract-docx'
import { rewriteResume, type RewriteResult } from './lib/rewrite-resume'
import { buildModifiedDocx } from './lib/modify-docx'
import './App.css'

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [doc, setDoc] = useState<DocxDoc | null>(null)
  const [jobDescription, setJobDescription] = useState('')
  const [result, setResult] = useState<RewriteResult | null>(null)
  const [isDragging, setIsDragging] = useState(false)
  const [isReading, setIsReading] = useState(false)
  const [isRewriting, setIsRewriting] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const handleFile = async (incoming: File | undefined) => {
    if (!incoming) return
    if (
      incoming.type !== DOCX_MIME &&
      !incoming.name.toLowerCase().endsWith('.docx')
    ) {
      setError('Only .docx files are supported.')
      return
    }
    setError(null)
    setFile(incoming)
    setResult(null)
    setIsReading(true)
    try {
      const parsed = await readDocx(incoming)
      setDoc(parsed)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to read DOCX.')
      setDoc(null)
    } finally {
      setIsReading(false)
    }
  }

  const onDrop = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setIsDragging(false)
    void handleFile(e.dataTransfer.files?.[0])
  }
  const onDragOver = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setIsDragging(true)
  }
  const onDragLeave = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setIsDragging(false)
  }
  const onChange = (e: ChangeEvent<HTMLInputElement>) => {
    void handleFile(e.target.files?.[0])
  }

  const onRewrite = async () => {
    if (!doc || !jobDescription.trim()) return
    setError(null)
    setIsRewriting(true)
    try {
      const res = await rewriteResume({
        paragraphs: doc.paragraphs,
        jobDescription,
      })
      setResult(res)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Rewrite failed.')
    } finally {
      setIsRewriting(false)
    }
  }

  const onDownloadDocx = () => {
    if (!doc || !result) return
    const blob = buildModifiedDocx(doc, result.replacements)
    const base = file?.name.replace(/\.docx$/i, '') ?? 'resume'
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${base}-tailored.docx`
    a.click()
    URL.revokeObjectURL(url)
  }

  const replacementCount = result ? Object.keys(result.replacements).length : 0

  return (
    <main className="page">
      <h1>Resume Maker</h1>
      <p className="subtitle">
        Drop a .docx, paste a job description, get a tailored version with
        original formatting preserved.
      </p>

      <label
        htmlFor="docx-input"
        className={`dropzone ${isDragging ? 'dropzone--active' : ''}`}
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
      >
        <input
          id="docx-input"
          type="file"
          accept=".docx"
          onChange={onChange}
          hidden
        />
        {file ? (
          <div className="file-info">
            <strong>{file.name}</strong>
            <span>{(file.size / 1024).toFixed(1)} KB</span>
            {isReading && <span>Reading…</span>}
            {doc && <span>{doc.paragraphs.length} paragraphs parsed</span>}
          </div>
        ) : (
          <div className="dropzone__prompt">
            <span className="dropzone__icon">📄</span>
            <p>Drag &amp; drop your .docx resume here</p>
            <p className="dropzone__hint">or click to browse</p>
          </div>
        )}
      </label>

      <section className="panel">
        <label className="field">
          <span>Job description</span>
          <textarea
            value={jobDescription}
            onChange={(e) => setJobDescription(e.target.value)}
            placeholder="Paste the job description here…"
            rows={8}
          />
        </label>
        <button
          type="button"
          className="primary"
          disabled={!doc || !jobDescription.trim() || isRewriting || isReading}
          onClick={onRewrite}
        >
          {isRewriting ? 'Tailoring…' : 'Tailor resume'}
        </button>
      </section>

      {result && (
        <section className="panel">
          <h2 className="section-title">
            Rewrote {replacementCount} paragraph
            {replacementCount === 1 ? '' : 's'}
          </h2>

          {result.wordCountMismatches.length > 0 && (
            <p className="warn">
              Dropped {result.wordCountMismatches.length} rewrite
              {result.wordCountMismatches.length === 1 ? '' : 's'} that didn't
              match original word count (would have broken alignment).
            </p>
          )}

          <div className="diff-list">
            {Object.entries(result.replacements).map(([idxStr, newText]) => {
              const idx = Number(idxStr)
              const original =
                doc?.paragraphs.find((p) => p.index === idx)?.text ?? ''
              return (
                <div key={idx} className="diff">
                  <div className="diff__col">
                    <span className="diff__label">Original</span>
                    <p>{original}</p>
                  </div>
                  <div className="diff__col">
                    <span className="diff__label">Tailored</span>
                    <p>{newText}</p>
                  </div>
                </div>
              )
            })}
          </div>

          <div className="actions">
            <button
              type="button"
              className="primary"
              onClick={onDownloadDocx}
            >
              Download tailored .docx
            </button>
          </div>
          <p className="hint">
            Open the downloaded file in <strong>Microsoft Word</strong>, then
            File&nbsp;→&nbsp;Save&nbsp;As&nbsp;→&nbsp;PDF for the best result.
          </p>
        </section>
      )}

      {error && <p className="error">{error}</p>}
    </main>
  )
}

export default App
