import { useState } from 'react'
import type { DragEvent, ChangeEvent } from 'react'
import './App.css'

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [isDragging, setIsDragging] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const handleFile = (incoming: File | undefined) => {
    if (!incoming) return
    if (incoming.type !== 'application/pdf') {
      setError('Only PDF files are supported.')
      setFile(null)
      return
    }
    setError(null)
    setFile(incoming)
  }

  const onDrop = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setIsDragging(false)
    handleFile(e.dataTransfer.files?.[0])
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
    handleFile(e.target.files?.[0])
  }

  return (
    <main className="page">
      <h1>Resume Maker</h1>
      <p className="subtitle">Drop a PDF to get started.</p>

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
          </div>
        ) : (
          <div className="dropzone__prompt">
            <span className="dropzone__icon">📄</span>
            <p>Drag &amp; drop your PDF here</p>
            <p className="dropzone__hint">or click to browse</p>
          </div>
        )}
      </label>

      {error && <p className="error">{error}</p>}
    </main>
  )
}

export default App
