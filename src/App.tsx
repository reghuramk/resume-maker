import { useState } from 'react'
import type { DragEvent, ChangeEvent } from 'react'
import { extractPdfText } from './lib/extract-pdf'
import { rewriteResume } from './lib/rewrite-resume'
import './App.css'

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [resumeText, setResumeText] = useState('')
  const [jobDescription, setJobDescription] = useState('')
  const [rewritten, setRewritten] = useState('')
  const [isDragging, setIsDragging] = useState(false)
  const [isExtracting, setIsExtracting] = useState(false)
  const [isRewriting, setIsRewriting] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const handleFile = async (incoming: File | undefined) => {
    if (!incoming) return
    if (incoming.type !== 'application/pdf') {
      setError('Only PDF files are supported.')
      return
    }
    setError(null)
    setFile(incoming)
    setIsExtracting(true)
    try {
      const text = await extractPdfText(incoming)
      setResumeText(text)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to read PDF.')
    } finally {
      setIsExtracting(false)
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
    if (!resumeText.trim() || !jobDescription.trim()) return
    setError(null)
    setIsRewriting(true)
    try {
      const result = await rewriteResume({ resumeText, jobDescription })
      setRewritten(result)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Rewrite failed.')
    } finally {
      setIsRewriting(false)
    }
  }

  return (
    <main className="page">
      <h1>Resume Maker</h1>
      <p className="subtitle">
        Drop your resume, paste the job description, and tailor it.
      </p>

      <label
        htmlFor="pdf-input"
        className={`dropzone ${isDragging ? 'dropzone--active' : ''}`}
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
      >
        <input
          id="pdf-input"
          type="file"
          accept="application/pdf"
          onChange={onChange}
          hidden
        />
        {file ? (
          <div className="file-info">
            <strong>{file.name}</strong>
            <span>{(file.size / 1024).toFixed(1)} KB</span>
            {isExtracting && <span>Extracting…</span>}
          </div>
        ) : (
          <div className="dropzone__prompt">
            <span className="dropzone__icon">📄</span>
            <p>Drag &amp; drop your PDF here</p>
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
          disabled={
            !resumeText.trim() ||
            !jobDescription.trim() ||
            isRewriting ||
            isExtracting
          }
          onClick={onRewrite}
        >
          {isRewriting ? 'Tailoring…' : 'Tailor resume'}
        </button>
      </section>

      {resumeText && (
        <section className="panel">
          <label className="field">
            <span>Extracted resume text (editable)</span>
            <textarea
              value={resumeText}
              onChange={(e) => setResumeText(e.target.value)}
              rows={10}
            />
          </label>
        </section>
      )}

      {rewritten && (
        <section className="panel">
          <label className="field">
            <span>Tailored resume (editable)</span>
            <textarea
              value={rewritten}
              onChange={(e) => setRewritten(e.target.value)}
              rows={16}
            />
          </label>
        </section>
      )}

      {error && <p className="error">{error}</p>}
    </main>
  )
}

export default App
