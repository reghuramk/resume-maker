import { useState } from 'react'
import type { DragEvent, ChangeEvent } from 'react'
import { readDocx, type DocxDoc } from './lib/extract-docx'
import {
  rewriteResume,
  describePersona,
  type RewriteResult,
} from './lib/rewrite-resume'
import { buildModifiedDocx } from './lib/modify-docx'
import './App.css'

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [doc, setDoc] = useState<DocxDoc | null>(null)
  const [jobDescription, setJobDescription] = useState('')
  const [personaMode, setPersonaMode] = useState(false)
  const [roleTitle, setRoleTitle] = useState('')
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
        personaMode,
        roleTitle: roleTitle.trim() || undefined,
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
    const insertions = result.generatedBullets.map((g) => ({
      templateAfterIndex: g.templateAfterIndex,
      text: g.text,
    }))
    const blob = buildModifiedDocx(doc, result.replacements, insertions)
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

        <label className="checkbox-row">
          <input
            type="checkbox"
            checked={personaMode}
            onChange={(e) => setPersonaMode(e.target.checked)}
          />
          <span>
            Aggressive persona rewrite — extract a target persona from the
            JD and transmute every bullet's tech stack to fit it. Falls back
            to keyword-swap automatically if the JD is too generic.
          </span>
        </label>

        {personaMode && (
          <label className="field">
            <span>Role title (optional — helps persona extraction)</span>
            <input
              type="text"
              value={roleTitle}
              onChange={(e) => setRoleTitle(e.target.value)}
              placeholder="e.g. Senior Go Fullstack Engineer"
            />
          </label>
        )}

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
          <div
            className={`mode-badge ${
              result.mode === 'persona-transmute'
                ? 'mode-badge--persona'
                : 'mode-badge--swap'
            }`}
          >
            <strong>
              Mode:{' '}
              {result.mode === 'persona-transmute'
                ? 'Persona transmute'
                : 'Keyword swap'}
            </strong>
            {result.mode === 'persona-transmute' && result.persona && (
              <span> — {describePersona(result.persona)}</span>
            )}
            {result.personaFallbackReason && (
              <span className="mode-badge__fallback">
                {' '}
                — Requested persona mode, fell back: {result.personaFallbackReason}
              </span>
            )}
          </div>

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

          {result.bumpedForRepetition.length > 0 && (
            <p className="warn">
              Kept {result.bumpedForRepetition.length} original bullet
              {result.bumpedForRepetition.length === 1 ? '' : 's'} unchanged
              because the rewrite would have repeated a must-have keyword more
              than {2} times.
            </p>
          )}

          {result.alignmentRelaxed.length > 0 && (
            <p className="hint">
              {result.alignmentRelaxed.length} bullet
              {result.alignmentRelaxed.length === 1 ? '' : 's'} accepted with
              relaxed word-count (minor line-wrap shift possible) to preserve
              must-have coverage that would otherwise be missing.
            </p>
          )}

          {result.resurrectedForCoverage.length > 0 && (
            <p className="hint">
              {result.resurrectedForCoverage.length} rewrite
              {result.resurrectedForCoverage.length === 1 ? '' : 's'}{' '}
              resurrected to fill must-haves that would otherwise be missing.
              A keyword may appear in 3 bullets instead of 2 as a result — the
              trade-off for higher Boolean coverage.
            </p>
          )}

          {Object.values(result.skillsPaddingApplied).flat().length > 0 && (
            <p className="hint">
              Added to your Skills section:{' '}
              <strong>
                {Object.values(result.skillsPaddingApplied)
                  .flat()
                  .join(', ')}
              </strong>{' '}
              — must-haves that didn't fit in any work bullet got appended to
              the skills lines so ATS Boolean coverage holds.
            </p>
          )}

          <div
            className={`boolean-summary ${
              result.booleanMatch.passed
                ? 'boolean-summary--pass'
                : 'boolean-summary--fail'
            }`}
          >
            <div className="boolean-summary__head">
              <strong>
                Recruiter Boolean check:{' '}
                {result.booleanMatch.passed ? 'PASSED' : 'INCOMPLETE'}
              </strong>
              <span>
                {result.booleanMatch.matched.length} /{' '}
                {result.mustHaves.length} must-haves present (
                {result.booleanMatch.scorePct}%)
              </span>
            </div>
            <div className="boolean-summary__row">
              <span className="boolean-summary__label">Matched:</span>
              <span>
                {result.booleanMatch.matched.map((m) => m.term).join(', ') ||
                  '—'}
              </span>
            </div>
            {result.booleanMatch.missing.length > 0 && (
              <div className="boolean-summary__row">
                <span className="boolean-summary__label">
                  Still missing:
                </span>
                <span>
                  {result.booleanMatch.missing.map((m) => m.term).join(', ')}
                </span>
              </div>
            )}
            <p className="hint" style={{ marginTop: '0.5rem' }}>
              This simulates the Boolean AND-query a recruiter would run in
              Workday / Greenhouse / Lever. Missing must-haves likely have no
              natural home in your existing experience — consider adding a
              skills line or a project bullet manually if they're load-bearing.
            </p>
          </div>

          <div className="diff-list">
            {Object.entries(result.replacements).map(([idxStr, newText]) => {
              const idx = Number(idxStr)
              const original =
                doc?.paragraphs.find((p) => p.index === idx)?.text ?? ''
              const added = result.addedByIndex[idx] ?? []
              return (
                <div key={idx} className="diff">
                  <div className="diff__col">
                    <span className="diff__label">Original</span>
                    <p>{original}</p>
                  </div>
                  <div className="diff__col">
                    <span className="diff__label">Tailored</span>
                    <p>{newText}</p>
                    {added.length > 0 && (
                      <span className="badge badge--good">
                        +{added.length} must-have: {added.join(', ')}
                      </span>
                    )}
                  </div>
                </div>
              )
            })}
          </div>

          {result.fabricatedReplacements.length > 0 && (
            <div className="generated-block generated-block--strong">
              <h3 className="section-title">
                ⚠ {result.fabricatedReplacements.length} bullet
                {result.fabricatedReplacements.length === 1 ? '' : 's'}{' '}
                replaced with fabricated content (Tier 3)
              </h3>
              <p className="hint">
                The original bullets had zero JD vocabulary, and the
                must-haves below had no natural home in your real experience.
                Each replacement is{' '}
                <strong>completely fabricated</strong> using scale figures
                already in your resume. Read carefully before downloading and
                decide whether you're comfortable shipping it.
              </p>
              {result.fabricatedReplacements.map((f) => (
                <div key={f.index} className="generated-bullet">
                  <div className="generated-bullet__meta">
                    <span className="badge badge--good">
                      +{f.mustHaveTerm}
                    </span>
                    <span className="generated-bullet__role">
                      replaces a low-value bullet
                    </span>
                  </div>
                  <p style={{ textDecoration: 'line-through', color: '#999' }}>
                    {f.originalText}
                  </p>
                  <p>{f.newText}</p>
                </div>
              ))}
            </div>
          )}

          {result.generatedBullets.length > 0 && (
            <div className="generated-block">
              <h3 className="section-title">
                {result.generatedBullets.length} new bullet
                {result.generatedBullets.length === 1 ? '' : 's'} generated to
                cover missing must-haves
              </h3>
              <p className="hint">
                Each will be inserted into your resume under the listed role,
                matching the surrounding bullet's formatting exactly. Review
                them before downloading — these are <strong>fabricated</strong>{' '}
                to fit must-haves your existing experience didn't cover.
              </p>
              {result.generatedBullets.map((g, i) => (
                <div key={i} className="generated-bullet">
                  <div className="generated-bullet__meta">
                    <span className="badge badge--good">
                      +{g.mustHaveTerm}
                    </span>
                    <span className="generated-bullet__role">
                      attributed to: <strong>{g.roleHeader}</strong>
                    </span>
                  </div>
                  <p>{g.text}</p>
                </div>
              ))}
            </div>
          )}

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
